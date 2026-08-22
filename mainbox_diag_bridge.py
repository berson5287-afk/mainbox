"""MaINbox Diagnostic Bridge -- localhost control surface for Claude Code.

ENGINE_VERSION 1.1.0

Purpose
-------
Lets an AI coding agent (Claude Code) running ON THIS MACHINE drive and inspect
a live MaINbox instance: read state, tail the ops log, trigger scans, and inject
SYNTHETIC emails through the real import pipeline -- without screenshots, file
ferrying, or touching real Outlook mail.

Security model (deliberate, hard limits)
----------------------------------------
* OFF by default. The app only starts the bridge when MAINBOX_DIAG=1.
* Binds 127.0.0.1 ONLY. Never a LAN/Tailscale interface.
* Every request must carry the per-boot token (X-Diag-Token header). The token
  is written to <data_dir>/diag_token.txt with user-only visibility; a client on
  the same machine reads it from there.
* READ-MOSTLY + SAFE TRIGGERS ONLY. There is deliberately NO endpoint that
  sends email, deletes data, or modifies Outlook. The v1.1.0 diagnostic
  endpoints (/why /rules /settings /audit) are strictly READ-ONLY -- they
  assemble facts the app already holds and mutate nothing, so this guarantee
  is unchanged. Synthetic emails injected for
  testing are tagged (entry_id prefix 'diag:') and can be purged as a group.

Threading model
---------------
The HTTP server runs on a daemon thread and NEVER touches Tk or Outlook COM
itself. Every action is marshaled to the app through two callbacks the app
registers at startup:

    run_on_main(fn)        -> schedules fn() on the Tk main thread (root.after)
    run_outlook(fn, done)  -> runs fn(outlook, namespace) on the app's Outlook
                              worker (the same run_outlook_worker path every
                              scan uses), then done(result_or_None)

The HTTP handler blocks on a threading.Event (bounded timeout) until the app
side finishes, then returns JSON. A busy Outlook worker returns 409 rather
than queueing forever.

Endpoints
---------
  GET  /state              app version, counts, flags, scan freshness
  GET  /trackers           sent/waiting placeholders (active + recently cleared)
  GET  /groups             groups + membership counts
  GET  /emails?n=&q=       loaded email rows (trimmed), newest first, optional
                           substring filter on subject/sender
  GET  /coverage?job=      quote-coverage ledger summary (or one job in full)
  GET  /opslog?n=          last n ops-log records (parsed JSON)
  GET  /why?eid=           v1.1.0 READ-ONLY decision trace for ONE row: status and
                           who set it, learned sender/domain rule AND its confidence,
                           durable type, clamp verdicts, group + how it was joined,
                           coverage linkage, and the ops-log lines naming that row
  GET  /rules?addr=        v1.1.0 learned sender/domain/follow-up rules + contact
                           registry for one address, with confidence and siblings
  GET  /settings           v1.1.0 effective settings (which toggles are actually live)
  GET  /audit              v1.1.0 live-state audit: stuck groups, replies parked in
                           the reply queue, trackers active on completed rows.
                           REPORTS ONLY -- it never repairs anything.
  POST /scan/sent          run one sent-tracker scan pass now (Outlook worker)
  POST /reconcile          run the tracker reconcile sweep now (Outlook worker)
  POST /inject_email       body: {subject, sender_email, body, source?, type?}
                           -> builds a synthetic email dict and runs the REAL
                           import path (upsert, triage hooks, grouping, tracker
                           clears) on the main thread; returns the resulting
                           row state + tracker/group effects. No COM involved.
  POST /purge_synthetic    remove every injected ('diag:') email
  GET  /ping               liveness (auth still required)

All responses: {"ok": bool, ...} JSON. Errors: {"ok": false, "error": "..."}.
"""

import json
import os
import secrets
import threading
import traceback
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

ENGINE_VERSION = "1.1.0"  # v1.1.0: phase-1 READ-ONLY diagnostic endpoints /why /rules /settings /audit

DEFAULT_PORT = 8765
ACTION_TIMEOUT_SECONDS = 150.0  # sent scan can legitimately take a while


class DiagBridge:
    def __init__(self, data_dir, callbacks, port=None, token=None):
        """callbacks: dict the APP supplies. Required keys:
             state()                    -> dict
             trackers()                 -> list[dict]
             groups()                   -> dict
             emails(n, q)               -> list[dict]
             coverage(job)              -> dict
             opslog_path()              -> str
             run_on_main(fn)            -> None (schedules fn on Tk thread)
             scan_sent(done)            -> bool (False = worker busy)
             reconcile(done)            -> bool (False = worker busy)
             inject_email(payload)      -> dict  (called ON MAIN via run_on_main)
             purge_synthetic()          -> int   (called ON MAIN via run_on_main)
        """
        self.data_dir = data_dir
        self.cb = dict(callbacks or {})
        self.port = int(port or os.environ.get("MAINBOX_DIAG_PORT", DEFAULT_PORT))
        self.token = token or os.environ.get("MAINBOX_DIAG_TOKEN") or secrets.token_hex(16)
        self.httpd = None
        self.thread = None
        self.started_at = None

    # ------------------------------------------------------------------ setup
    def write_token_file(self):
        try:
            path = os.path.join(self.data_dir, "diag_token.txt")
            with open(path, "w", encoding="utf-8") as f:
                f.write(self.token)
            return path
        except Exception:
            return ""

    def start(self):
        bridge = self

        class Handler(BaseHTTPRequestHandler):
            server_version = "MaINboxDiag/" + ENGINE_VERSION

            # silence default stderr logging; the app has its own logs
            def log_message(self, *a):
                pass

            def _send(self, code, obj):
                try:
                    body = json.dumps(obj, default=str).encode("utf-8")
                    self.send_response(code)
                    self.send_header("Content-Type", "application/json")
                    self.send_header("Content-Length", str(len(body)))
                    self.end_headers()
                    self.wfile.write(body)
                except Exception:
                    pass

            def _auth_ok(self):
                return self.headers.get("X-Diag-Token", "") == bridge.token

            def _read_json_body(self):
                try:
                    n = int(self.headers.get("Content-Length", 0) or 0)
                    if n <= 0 or n > 1_000_000:
                        return {}
                    return json.loads(self.rfile.read(n).decode("utf-8"))
                except Exception:
                    return {}

            def do_GET(self):
                if not self._auth_ok():
                    return self._send(401, {"ok": False, "error": "bad or missing X-Diag-Token"})
                u = urlparse(self.path)
                q = {k: v[0] for k, v in parse_qs(u.query).items()}
                try:
                    if u.path == "/ping":
                        return self._send(200, {"ok": True, "version": ENGINE_VERSION,
                                                "started_at": bridge.started_at})
                    if u.path == "/state":
                        return self._send(200, {"ok": True, "state": bridge.cb["state"]()})
                    if u.path == "/trackers":
                        return self._send(200, {"ok": True, "trackers": bridge.cb["trackers"]()})
                    if u.path == "/groups":
                        return self._send(200, {"ok": True, "groups": bridge.cb["groups"]()})
                    if u.path == "/emails":
                        n = max(1, min(500, int(q.get("n", 50) or 50)))
                        return self._send(200, {"ok": True,
                                                "emails": bridge.cb["emails"](n, q.get("q", ""))})
                    if u.path == "/coverage":
                        return self._send(200, {"ok": True,
                                                "coverage": bridge.cb["coverage"](q.get("job", ""))})
                    if u.path == "/opslog":
                        n = max(1, min(5000, int(q.get("n", 200) or 200)))
                        return self._send(200, {"ok": True, "records": bridge.read_opslog_tail(n)})
                    # ---- v0.2.0 phase 1: READ-ONLY diagnostic surface -------------
                    # These mutate nothing. Each is served only if the app supplied the
                    # callback, so a newer app + older bridge (or the reverse) degrades
                    # to a clean 404 instead of a crash.
                    if u.path == "/why":
                        cb = bridge.cb.get("why")
                        if not cb:
                            return self._send(404, {"ok": False, "error": "why not supported by this app build"})
                        return self._send(200, {"ok": True, "why": cb(q.get("eid", ""))})
                    if u.path == "/rules":
                        cb = bridge.cb.get("rules")
                        if not cb:
                            return self._send(404, {"ok": False, "error": "rules not supported by this app build"})
                        return self._send(200, {"ok": True, "rules": cb(q.get("addr", ""))})
                    if u.path == "/settings":
                        cb = bridge.cb.get("settings")
                        if not cb:
                            return self._send(404, {"ok": False, "error": "settings not supported by this app build"})
                        return self._send(200, {"ok": True, "settings": cb()})
                    if u.path == "/audit":
                        cb = bridge.cb.get("audit")
                        if not cb:
                            return self._send(404, {"ok": False, "error": "audit not supported by this app build"})
                        return self._send(200, {"ok": True, "audit": cb()})
                    return self._send(404, {"ok": False, "error": "unknown endpoint"})
                except Exception as e:
                    return self._send(500, {"ok": False, "error": repr(e),
                                            "trace": traceback.format_exc()[-1500:]})

            def do_POST(self):
                if not self._auth_ok():
                    return self._send(401, {"ok": False, "error": "bad or missing X-Diag-Token"})
                u = urlparse(self.path)
                try:
                    if u.path == "/scan/sent":
                        return self._send(*bridge.run_worker_action("scan_sent"))
                    if u.path == "/reconcile":
                        return self._send(*bridge.run_worker_action("reconcile"))
                    if u.path == "/inject_email":
                        payload = self._read_json_body()
                        if not str(payload.get("subject", "")).strip():
                            return self._send(400, {"ok": False, "error": "subject required"})
                        return self._send(*bridge.run_main_action("inject_email", payload))
                    if u.path == "/purge_synthetic":
                        return self._send(*bridge.run_main_action("purge_synthetic"))
                    return self._send(404, {"ok": False, "error": "unknown endpoint"})
                except Exception as e:
                    return self._send(500, {"ok": False, "error": repr(e),
                                            "trace": traceback.format_exc()[-1500:]})

        self.httpd = ThreadingHTTPServer(("127.0.0.1", self.port), Handler)
        self.httpd.daemon_threads = True
        self.thread = threading.Thread(target=self.httpd.serve_forever,
                                       name="mainbox-diag-bridge", daemon=True)
        self.thread.start()
        self.started_at = datetime.now().isoformat(timespec="seconds")
        return self.port

    def stop(self):
        try:
            if self.httpd:
                self.httpd.shutdown()
        except Exception:
            pass

    # ---------------------------------------------------------------- actions
    def run_worker_action(self, name):
        """Run an Outlook-worker action (scan/reconcile); block until done or timeout."""
        done_evt = threading.Event()
        box = {}

        def done(result):
            box["result"] = result
            done_evt.set()

        try:
            accepted = self.cb[name](done)
        except Exception as e:
            return 500, {"ok": False, "error": repr(e)}
        if not accepted:
            return 409, {"ok": False, "error": "outlook worker busy; retry shortly"}
        if not done_evt.wait(ACTION_TIMEOUT_SECONDS):
            return 504, {"ok": False, "error": f"timed out after {ACTION_TIMEOUT_SECONDS:.0f}s"}
        return 200, {"ok": True, "result": box.get("result")}

    def run_main_action(self, name, payload=None):
        """Run a main-thread action (inject/purge); block until done or timeout."""
        done_evt = threading.Event()
        box = {}

        def on_main():
            try:
                if payload is not None:
                    box["result"] = self.cb[name](payload)
                else:
                    box["result"] = self.cb[name]()
            except Exception as e:
                box["error"] = repr(e)
                box["trace"] = traceback.format_exc()[-1500:]
            finally:
                done_evt.set()

        try:
            self.cb["run_on_main"](on_main)
        except Exception as e:
            return 500, {"ok": False, "error": "run_on_main failed: " + repr(e)}
        if not done_evt.wait(30.0):
            return 504, {"ok": False, "error": "main thread did not respond within 30s"}
        if "error" in box:
            return 500, {"ok": False, "error": box["error"], "trace": box.get("trace", "")}
        return 200, {"ok": True, "result": box.get("result")}

    # ----------------------------------------------------------------- opslog
    def read_opslog_tail(self, n):
        try:
            path = self.cb["opslog_path"]()
            if not path or not os.path.exists(path):
                return []
            with open(path, "rb") as f:
                f.seek(0, os.SEEK_END)
                size = f.tell()
                f.seek(max(0, size - max(4096, n * 400)))
                chunk = f.read().decode("utf-8", "replace")
            lines = chunk.splitlines()
            if size > len(chunk):
                lines = lines[1:]  # first line may be partial
            out = []
            for ln in lines[-n:]:
                try:
                    out.append(json.loads(ln))
                except Exception:
                    continue
            return out
        except Exception:
            return []
