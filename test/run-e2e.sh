#!/bin/sh
# Serves the repo locally and runs test/e2e.html in headless Chrome, which
# drives the real index.html (upload, drag & drop, samples, downloads).
set -e
DIR="$(cd "$(dirname "$0")/.." && pwd)"
PORT="${PORT:-8765}"
CHROME="${CHROME:-/Applications/Google Chrome.app/Contents/MacOS/Google Chrome}"

STARTED_SERVER=""
if ! curl -s -o /dev/null "http://127.0.0.1:$PORT/"; then
  (cd "$DIR" && python3 -m http.server "$PORT" --bind 127.0.0.1 >/dev/null 2>&1) &
  STARTED_SERVER=$!
  sleep 1
fi

PROFILE="$(mktemp -d)"
OUTPUT="$("$CHROME" --headless=new --disable-gpu --user-data-dir="$PROFILE" \
  --virtual-time-budget=20000 --dump-dom "http://127.0.0.1:$PORT/test/e2e.html" 2>/dev/null \
  | sed -n '/<pre id="results">/,/<\/pre>/p' | sed 's/<[^>]*>//g; s/&amp;/\&/g; s/&lt;/</g; s/&gt;/>/g; s/&quot;/"/g')"
rm -rf "$PROFILE"
[ -n "$STARTED_SERVER" ] && kill "$STARTED_SERVER" 2>/dev/null || true

echo "$OUTPUT"
if [ -z "$OUTPUT" ] || echo "$OUTPUT" | grep -qE '^(FAIL|RUNNING)'; then
  exit 1
fi
