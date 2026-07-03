#!/usr/bin/env python3
import json
import subprocess
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

class PipelineHandler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        data = json.loads(body)

        file_id   = data.get('file_id', '')
        file_name = data.get('file_name', '')

        print(f'Received: {file_name} ({file_id})', flush=True)

        if not file_id:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b'{"error": "file_id required"}')
            return

        # Route: _approved → Agent 3 (W2), else → Agent 1+2 (W1)
        if '_approved' in file_name.lower():
            pipeline_script = '/root/pipeline_w2.py'
            print(f'Routing to W2 (Agent 3 — approved planning doc)', flush=True)
        else:
            pipeline_script = '/root/pipeline_w1.py'
            print(f'Routing to W1 (Agent 1+2 — new AO)', flush=True)

        # Respond immediately — pipeline runs in background
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({"status": "accepted", "file_name": file_name, "pipeline": pipeline_script}).encode())

        # Run pipeline in background thread
        def run():
            print(f'Starting pipeline: {pipeline_script} for {file_name}', flush=True)
            result = subprocess.run(
                ['python3', pipeline_script, file_id, file_name],
                capture_output=True, text=True, timeout=600
            )
            print(result.stdout, flush=True)
            if result.stderr:
                print(f'STDERR: {result.stderr}', flush=True)
            # Parse N8N_OUTPUT and save to fixed path
            output = {}
            for line in result.stdout.split('\n'):
                if line.startswith('N8N_OUTPUT:'):
                    output = json.loads(line.replace('N8N_OUTPUT:', ''))
                    break
            with open('/local-files/LATEST_N8N_OUTPUT.json', 'w') as f:
                json.dump(output, f)
            print(f'Pipeline done: {output}', flush=True)

        threading.Thread(target=run, daemon=True).start()

    def log_message(self, format, *args):
        print(format % args, flush=True)

if __name__ == '__main__':
    server = HTTPServer(('0.0.0.0', 5679), PipelineHandler)
    print('Pipeline server running on port 5679', flush=True)
    server.serve_forever()
