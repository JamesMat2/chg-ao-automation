#!/bin/bash
# deploy.sh — simple deploy for CHG AO automation pipeline
# Not CI/CD by design: manual trigger, plain git pull, safety check before restart.
set -e

cd /root

echo "==> Pulling latest from origin/master..."
git pull origin master

echo "==> Syntax-checking all tracked Python files..."
FAILED=0
for f in $(git ls-files '*.py'); do
    if ! python3 -m py_compile "$f" 2>/tmp/deploy_syntax_error.log; then
        echo "    SYNTAX ERROR in $f:"
        cat /tmp/deploy_syntax_error.log
        FAILED=1
    fi
done

if [ "$FAILED" -eq 1 ]; then
    echo "==> ABORTED: syntax errors found above. pipeline-server was NOT restarted."
    echo "==> Fix the error(s), commit, push, and re-run deploy.sh."
    exit 1
fi

echo "==> All files compiled cleanly."
echo "==> Restarting pipeline-server..."
systemctl restart pipeline-server
sleep 2
STATUS=$(systemctl is-active pipeline-server)

if [ "$STATUS" = "active" ]; then
    echo "==> SUCCESS: pipeline-server is active."
else
    echo "==> WARNING: pipeline-server status is '$STATUS' after restart."
    echo "==> Check: journalctl -u pipeline-server -n 50 --no-pager"
    exit 1
fi
