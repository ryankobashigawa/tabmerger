#!/bin/sh
# Serves the repo locally (with GitHub Pages-style extensionless URLs) and runs
# test/e2e.html in headless Chrome, which drives the real pages: upload, drag
# & drop, samples, downloads, merge into tabs, split, and the hand-off.
set -e
DIR="$(cd "$(dirname "$0")/.." && pwd)"
PORT="${PORT:-8765}"

(cd "$DIR" && exec python3 test/serve.py "$PORT" >/dev/null 2>&1) &
SERVER=$!
trap 'kill "$SERVER" 2>/dev/null || true' EXIT
sleep 1

OUTPUT="$(python3 "$DIR/test/run-e2e.py" "http://127.0.0.1:$PORT/test/e2e.html" 180)"
echo "$OUTPUT"
if [ -z "$OUTPUT" ] || echo "$OUTPUT" | grep -qE '^(FAIL|RUNNING)'; then
  exit 1
fi
