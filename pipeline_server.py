#!/usr/bin/env python3
import json
import os
import subprocess
import threading
import time
import glob
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler

STATUS_DIR = '/local-files'


def write_status(file_id, data):
    """Atomically write a per-job status file so concurrent jobs can
    never clobber each other's result (unlike a single shared file)."""
    path = os.path.join(STATUS_DIR, f'STATUS_{file_id}.json')
    tmp_path = path + '.tmp'
    with open(tmp_path, 'w') as f:
        json.dump(data, f)
    os.rename(tmp_path, path)  # atomic on same filesystem


class PipelineHandler(BaseHTTPRequestHandler):
    def _handle_batch(self, data):
        """New (Session 34): handle a multi-file AO batch. Always routes to
        W1 (Agent 1+2) -- batches are only used for new AO submissions, never
        for the _approved OS-redaction trigger (which is always single-file).
        """
        files = data.get('files', [])
        if not files:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b'{"error": "files array is empty"}')
            return

        batch_id = f"batch_{int(time.time())}_{os.urandom(4).hex()}"
        file_names = ", ".join(f.get('file_name', '?') for f in files)
        print(f'Received BATCH: {len(files)} file(s) [{file_names}] -> {batch_id}', flush=True)

        batch_json_path = os.path.join(STATUS_DIR, f'BATCH_{batch_id}.json')
        with open(batch_json_path, 'w') as f:
            json.dump(files, f)

        job_label = f"AO_multifichiers_{len(files)}fichiers"
        pipeline_script = '/root/pipeline_w1.py'
        print(f'Routing BATCH to W1 (Agent 1+2 - new AO, multi-fichiers)', flush=True)

        write_status(batch_id, {
            "status": "processing",
            "file_name": job_label,
            "pipeline": pipeline_script,
            "started_at": time.time()
        })

        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "accepted", "file_name": job_label, "pipeline": pipeline_script, "job_id": batch_id}).encode())

        def run():
            print(f'Starting pipeline: {pipeline_script} for {job_label}', flush=True)
            try:
                result = subprocess.run(
                    ['python3', pipeline_script, batch_id, job_label, batch_json_path],
                    capture_output=True, text=True, timeout=2400
                )
                print(result.stdout, flush=True)
                if result.stderr:
                    print(f'STDERR: {result.stderr}', flush=True)

                if result.returncode != 0:
                    write_status(batch_id, {
                        "status": "error",
                        "file_name": job_label,
                        "pipeline": pipeline_script,
                        "error": f"exit code {result.returncode}: {result.stderr[-1000:]}",
                        "finished_at": time.time()
                    })
                    print(f'Pipeline FAILED (exit {result.returncode}) for {job_label}', flush=True)
                    return

                output = {}
                for line in result.stdout.split('\n'):
                    if line.startswith('N8N_OUTPUT:'):
                        output = json.loads(line.replace('N8N_OUTPUT:', ''))
                        break

                with open('/local-files/LATEST_N8N_OUTPUT.json', 'w') as f:
                    json.dump(output, f)

                write_status(batch_id, {
                    "status": "done",
                    "file_name": job_label,
                    "pipeline": pipeline_script,
                    "outputs": output,
                    "finished_at": time.time()
                })
                print(f'Pipeline done: {output}', flush=True)

            except subprocess.TimeoutExpired:
                write_status(batch_id, {
                    "status": "error",
                    "file_name": job_label,
                    "pipeline": pipeline_script,
                    "error": "pipeline timed out after 1800s",
                    "finished_at": time.time()
                })
                print(f'Pipeline TIMED OUT for {job_label}', flush=True)
            except Exception as e:
                write_status(batch_id, {
                    "status": "error",
                    "file_name": job_label,
                    "pipeline": pipeline_script,
                    "error": str(e),
                    "finished_at": time.time()
                })
                print(f'Pipeline EXCEPTION for {job_label}: {e}', flush=True)

        threading.Thread(target=run, daemon=True).start()

    def _handle_addenda(self, data):
        """Session 40: validate an addenda submission against an existing
        AO's persisted numero_ao, then launch pipeline_w3.py (Session 40
        build step 3, not yet written) to compute the after-score and
        regenerate the Planification document only. This route's job is
        narrow: validate the match and launch -- all regeneration logic
        lives in pipeline_w3.py, not here."""
        numero_ao = (data.get('numero_ao') or '').strip()
        addendum_file_id = data.get('file_id', '')
        addendum_file_name = data.get('file_name', '')

        if not numero_ao or not addendum_file_id:
            self.send_response(400)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({
                "status": "error",
                "message": "numero_ao et file_id sont requis"
            }, ensure_ascii=False).encode('utf-8'))
            return

        matches = []
        pattern = os.path.join(STATUS_DIR, 'STATUS_*.json')
        for path in glob.glob(pattern):
            basename = os.path.basename(path)
            job_id = basename[len('STATUS_'):-len('.json')]
            # Session 46: batch_ AOs are now eligible for addenda matching.
            # pipeline_w3.py was extended this session to re-download a batch
            # AO's full original file set via its BATCH_<id>.json manifest --
            # the exclusion below (Session 39) predated that support, and had
            # been silently blocking nearly every real AO since Session 35
            # made every Upload Form submission route through the batch path.
            if job_id.startswith('addenda_'):
                continue  # Session 43: exclude addenda-job STATUS files from
                           # the matching pool -- otherwise the most recent
                           # addendum becomes a false match for the next
                           # addendum on the same AO, since it always has the
                           # most recent finished_at and also carries the same
                           # numero_ao (copied through by pipeline_w3.py).
            try:
                with open(path, 'r') as f:
                    status_data = json.load(f)
            except Exception as e:
                print(f'Skipping unreadable status file {path}: {e}', flush=True)
                continue
            if status_data.get('status') != 'done':
                continue
            outputs = status_data.get('outputs', {})
            if outputs.get('numero_ao', '') == numero_ao:
                matches.append((job_id, status_data))

        if not matches:
            print(f'Addenda: no match found for numero_ao={numero_ao!r}', flush=True)
            self.send_response(404)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps({
                "status": "error",
                "message": f"Aucun AO trouve avec le numero '{numero_ao}'. Veuillez verifier le numero et reessayer."
            }, ensure_ascii=False).encode('utf-8'))
            return

        if len(matches) > 1:
            print(f'Addenda: {len(matches)} matches for numero_ao={numero_ao!r}, using most recent by finished_at', flush=True)

        matches.sort(key=lambda m: m[1].get('finished_at', 0), reverse=True)
        original_file_id, matched_status = matches[0]

        job_id = f"addenda_{int(time.time())}_{os.urandom(4).hex()}"
        job_label = f"Addenda_{numero_ao}"
        pipeline_script = '/root/pipeline_w3.py'

        print(f'Addenda matched: numero_ao={numero_ao!r} -> original_file_id={original_file_id} -> {job_id}', flush=True)

        write_status(job_id, {
            "status": "processing",
            "file_name": job_label,
            "pipeline": pipeline_script,
            "numero_ao": numero_ao,
            "original_file_id": original_file_id,
            "started_at": time.time()
        })

        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps({
            "status": "accepted",
            "job_id": job_id,
            "numero_ao": numero_ao,
            "original_file_id": original_file_id
        }, ensure_ascii=False).encode('utf-8'))

        def run():
            print(f'Starting pipeline: {pipeline_script} for {job_label}', flush=True)
            try:
                result = subprocess.run(
                    ['python3', pipeline_script, job_id, original_file_id, addendum_file_id, addendum_file_name, numero_ao],
                    capture_output=True, text=True, timeout=2400
                )
                print(result.stdout, flush=True)
                if result.stderr:
                    print(f'STDERR: {result.stderr}', flush=True)

                if result.returncode != 0:
                    write_status(job_id, {
                        "status": "error",
                        "file_name": job_label,
                        "pipeline": pipeline_script,
                        "error": f"exit code {result.returncode}: {result.stderr[-1000:]}",
                        "finished_at": time.time()
                    })
                    print(f'Pipeline FAILED (exit {result.returncode}) for {job_label}', flush=True)
                    return

                output = {}
                for line in result.stdout.split('\n'):
                    if line.startswith('N8N_OUTPUT:'):
                        output = json.loads(line.replace('N8N_OUTPUT:', ''))
                        break

                write_status(job_id, {
                    "status": "done",
                    "file_name": job_label,
                    "pipeline": pipeline_script,
                    "outputs": output,
                    "finished_at": time.time()
                })
                print(f'Pipeline done: {output}', flush=True)

            except subprocess.TimeoutExpired:
                write_status(job_id, {
                    "status": "error",
                    "file_name": job_label,
                    "pipeline": pipeline_script,
                    "error": "pipeline timed out after 1800s",
                    "finished_at": time.time()
                })
                print(f'Pipeline TIMED OUT for {job_label}', flush=True)
            except Exception as e:
                write_status(job_id, {
                    "status": "error",
                    "file_name": job_label,
                    "pipeline": pipeline_script,
                    "error": str(e),
                    "finished_at": time.time()
                })
                print(f'Pipeline EXCEPTION for {job_label}: {e}', flush=True)

        threading.Thread(target=run, daemon=True).start()

    def do_GET(self):
        if self.path == '/list_aos':
            self._handle_list_aos()
            return
        self.send_response(404)
        self.end_headers()

    def _handle_list_aos(self):
        """Session 39: list completed AOs (single-file AND batch) for the
        addenda-upload dropdown. Label is built from client + titre + date
        to disambiguate resubmissions of the same AO."""
        aos = []
        pattern = os.path.join(STATUS_DIR, 'STATUS_*.json')
        for path in glob.glob(pattern):
            basename = os.path.basename(path)
            try:
                with open(path, 'r') as f:
                    data = json.load(f)
            except Exception as e:
                print(f'Skipping unreadable status file {path}: {e}', flush=True)
                continue
            if data.get('status') != 'done':
                continue

            outputs = data.get('outputs', {})
            client = outputs.get('ao_client', '') or ''
            titre = outputs.get('ao_titre', '') or ''
            file_name = data.get('file_name', '')
            finished_at = data.get('finished_at', 0)

            try:
                date_str = datetime.fromtimestamp(finished_at).strftime('%d %b')
            except Exception:
                date_str = ''

            if client and titre:
                label = f"{client} — {titre}" + (f" ({date_str})" if date_str else "")
            else:
                label = file_name or basename

            # id is the batch_id for batch AOs, file_id for single-file AOs --
            # either way it's the string between 'STATUS_' and '.json'
            job_id = basename[len('STATUS_'):-len('.json')]
            is_batch = job_id.startswith('batch_')

            aos.append({
                'file_id': job_id,
                'is_batch': is_batch,
                'label': label,
                'client': client,
                'titre': titre,
                'gonogo_score': outputs.get('gonogo_score', 0),
                'finished_at': finished_at
            })

        aos.sort(key=lambda a: a['finished_at'], reverse=True)
        body = json.dumps(aos, ensure_ascii=False).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        data = json.loads(body)

        if data.get('files'):
            self._handle_batch(data)
            return

        if data.get('numero_ao'):
            self._handle_addenda(data)
            return

        file_id   = data.get('file_id', '')
        file_name = data.get('file_name', '')

        print(f'Received: {file_name} ({file_id})', flush=True)

        if not file_id:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b'{"error": "file_id required"}')
            return

        # Route: _approved -> Agent 3 (W2), else -> Agent 1+2 (W1)
        if '_approved' in file_name.lower():
            pipeline_script = '/root/pipeline_w2.py'
            print(f'Routing to W2 (Agent 3 - approved planning doc)', flush=True)
        else:
            pipeline_script = '/root/pipeline_w1.py'
            print(f'Routing to W1 (Agent 1+2 - new AO)', flush=True)

        # Write initial per-job status BEFORE responding, so n8n can never
        # poll for a status file that doesn't exist yet.
        write_status(file_id, {
            "status": "processing",
            "file_name": file_name,
            "pipeline": pipeline_script,
            "started_at": time.time()
        })

        # Respond immediately - pipeline runs in background
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "accepted", "file_name": file_name, "pipeline": pipeline_script}).encode())

        # Run pipeline in background thread
        def run():
            print(f'Starting pipeline: {pipeline_script} for {file_name}', flush=True)
            try:
                result = subprocess.run(
                    ['python3', pipeline_script, file_id, file_name],
                    capture_output=True, text=True, timeout=2400
                )
                print(result.stdout, flush=True)
                if result.stderr:
                    print(f'STDERR: {result.stderr}', flush=True)

                if result.returncode != 0:
                    write_status(file_id, {
                        "status": "error",
                        "file_name": file_name,
                        "pipeline": pipeline_script,
                        "error": f"exit code {result.returncode}: {result.stderr[-1000:]}",
                        "finished_at": time.time()
                    })
                    print(f'Pipeline FAILED (exit {result.returncode}) for {file_name}', flush=True)
                    return

                # Parse N8N_OUTPUT
                output = {}
                for line in result.stdout.split('\n'):
                    if line.startswith('N8N_OUTPUT:'):
                        output = json.loads(line.replace('N8N_OUTPUT:', ''))
                        break

                # Legacy shared file kept for backward compatibility only -
                # the per-job status file below is now the source of truth.
                with open('/local-files/LATEST_N8N_OUTPUT.json', 'w') as f:
                    json.dump(output, f)

                write_status(file_id, {
                    "status": "done",
                    "file_name": file_name,
                    "pipeline": pipeline_script,
                    "outputs": output,
                    "finished_at": time.time()
                })
                print(f'Pipeline done: {output}', flush=True)

            except subprocess.TimeoutExpired:
                write_status(file_id, {
                    "status": "error",
                    "file_name": file_name,
                    "pipeline": pipeline_script,
                    "error": "pipeline timed out after 1800s",
                    "finished_at": time.time()
                })
                print(f'Pipeline TIMED OUT for {file_name}', flush=True)
            except Exception as e:
                write_status(file_id, {
                    "status": "error",
                    "file_name": file_name,
                    "pipeline": pipeline_script,
                    "error": str(e),
                    "finished_at": time.time()
                })
                print(f'Pipeline EXCEPTION for {file_name}: {e}', flush=True)

        threading.Thread(target=run, daemon=True).start()

    def log_message(self, format, *args):
        print(format % args, flush=True)


if __name__ == '__main__':
    server = HTTPServer(('0.0.0.0', 5679), PipelineHandler)
    print('Pipeline server running on port 5679', flush=True)
    server.serve_forever()
