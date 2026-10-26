"""Runs test/e2e.html in headless Chrome on real time and prints its results.

Talks to Chrome over --remote-debugging-pipe (the DevTools protocol on file
descriptors 3/4), so it needs nothing beyond Python's standard library. Real
time matters: Chrome's --virtual-time-budget mode never completes IndexedDB
requests, which the stack -> merge hand-off relies on.

Usage: python3 test/run-e2e.py <url-of-e2e.html> [timeout-seconds]
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time

CHROME = os.environ.get("CHROME", "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")
url = sys.argv[1]
deadline = time.time() + (float(sys.argv[2]) if len(sys.argv) > 2 else 180)

to_chrome_r, to_chrome_w = os.pipe()
from_chrome_r, from_chrome_w = os.pipe()


def wire_fds():
    os.dup2(to_chrome_r, 3)
    os.dup2(from_chrome_w, 4)


profile = tempfile.mkdtemp()
chrome = subprocess.Popen(
    [CHROME, "--headless=new", "--disable-gpu", "--remote-debugging-pipe",
     f"--user-data-dir={profile}", "about:blank"],
    # close_fds=False: Popen would otherwise close fds 3/4 after wire_fds runs.
    close_fds=False, preexec_fn=wire_fds,
    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
)
os.close(to_chrome_r)
os.close(from_chrome_w)
reader = os.fdopen(from_chrome_r, "rb")
writer = os.fdopen(to_chrome_w, "wb")
next_id = 0
buffered = b""


def send(method, params=None, session=None):
    global next_id, buffered
    next_id += 1
    msg = {"id": next_id, "method": method, "params": params or {}}
    if session:
        msg["sessionId"] = session
    writer.write(json.dumps(msg).encode() + b"\0")
    writer.flush()
    while True:  # read NUL-delimited messages until our reply arrives
        while b"\0" not in buffered:
            chunk = os.read(reader.fileno(), 65536)
            if not chunk:
                raise RuntimeError("Chrome exited")
            buffered += chunk
        raw, buffered = buffered.split(b"\0", 1)
        reply = json.loads(raw)
        if reply.get("id") == next_id:
            if "error" in reply:
                raise RuntimeError(reply["error"])
            return reply.get("result", {})


text = ""
try:
    target = send("Target.createTarget", {"url": url})["targetId"]
    session = send("Target.attachToTarget", {"targetId": target, "flatten": True})["sessionId"]
    while time.time() < deadline:
        time.sleep(0.5)
        res = send("Runtime.evaluate", {
            "expression": "document.getElementById('results') && document.getElementById('results').textContent",
            "returnByValue": True,
        }, session)
        text = res.get("result", {}).get("value") or ""
        if text and not text.startswith("RUNNING"):
            break
finally:
    chrome.kill()
    shutil.rmtree(profile, ignore_errors=True)

print(text or "RUNNING (no output)")
