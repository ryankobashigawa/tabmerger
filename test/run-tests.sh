#!/bin/sh
# Runs test/run-tests.js against core.js via macOS's built-in JavaScriptCore
# (no Node install required on this machine). core.js itself has no DOM
# dependency and would run the same way under Node if this app is ever
# deployed to an environment that has it.
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
OUTPUT="$(osascript -l JavaScript "$DIR/run-tests.js")"
echo "$OUTPUT"
if echo "$OUTPUT" | grep -q '^FAIL'; then
  exit 1
fi
