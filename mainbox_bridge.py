"""mainbox_bridge.py -- MaINbox <-> MaINbox Brain/Voice file bridge (adapter).

MaINbox imports ``MainboxAdapter`` from this module at startup (it searches
MAINBOX_BRIDGE_DIR, then %LOCALAPPDATA%\\MaINbox\\bridge) and, on its Tk main
thread every 4 seconds, calls:

    adapter.heartbeat(app_version=...)      -> writes heartbeat.json
    adapter.publish_followups(items)        -> writes followups.json (when changed)
    adapter.poll_commands(handlers={...})   -> runs queued commands, writes results

Everything is plain files in ONE folder so the phone side (the voice server,
same PC) needs no socket into the app, and every mutation still runs inside
MaINbox's own process on its main thread:

    bridge/
      heartbeat.json            {"ts", "app_version"}          (app -> world)
      followups.json            {"published_at", "items":[..]} (app -> world)
      commands/<id>.json        {"id","name","args","ts"}      (world -> app)
      results/<id>.json         {"id","ok",...}                (app -> world)

Commands are consumed in filename (timestamp) order and deleted after their
result is written; stale results are pruned after an hour. A handler that
raises is reported as {"ok": False, "error": ...} instead of crashing the
tick. This module is deliberately stdlib-only and never imports MaINbox.

ENGINE_VERSION 1.0.1 (2026-08-28): rename retry + in-place fallback.
1.0.0 (2026-08-26): first release, for MaINbox v4.2.99 + MaINbox Voice v0.10.
"""
from __future__ import annotations

import os
import json
import time
import traceback

ENGINE_VERSION = "1.0.2"
_RESULT_TTL_S = 3600


def _default_dir() -> str:
    env = (os.environ.get("MAINBOX_BRIDGE_DIR") or "").strip()
    if env:
        return env
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "MaINbox", "bridge")


def _atomic_write(path: str, data) -> None:
    tmp = f"{path}.{os.getpid()}.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1, default=str)
    # 1.0.1: Windows refuses the rename while another process (the voice
    # server) has the target open for reading -> "Access is denied" and a
    # logged tick error. Retry briefly, then fall back to an in-place write.
    # 1.0.2 (audit): NO in-place fallback -- truncating the file while a reader
    # holds it is exactly the torn read the atomic contract exists to prevent.
    # If the rename cannot land after retries, keep the previous version on disk
    # (stale beats torn; heartbeat/snapshot refresh 4 s later anyway).
    for attempt in range(8):
        try:
            os.replace(tmp, path)
            return
        except PermissionError:
            time.sleep(0.05 * (attempt + 1))
    try:
        os.remove(tmp)
    except OSError:
        pass


class MainboxAdapter:
    def __init__(self, bridge_dir: str | None = None):
        self.dir = bridge_dir or _default_dir()
        self.cmd_dir = os.path.join(self.dir, "commands")
        self.res_dir = os.path.join(self.dir, "results")
        for d in (self.dir, self.cmd_dir, self.res_dir):
            os.makedirs(d, exist_ok=True)
        self._last_prune = 0.0

    # -- app -> world -------------------------------------------------------
    def heartbeat(self, app_version: str = "") -> None:
        _atomic_write(os.path.join(self.dir, "heartbeat.json"),
                      {"ts": time.time(), "app_version": str(app_version or ""),
                       "bridge_version": ENGINE_VERSION})

    def publish_followups(self, items) -> None:
        _atomic_write(os.path.join(self.dir, "followups.json"),
                      {"published_at": time.time(), "items": list(items or [])})

    # -- world -> app -------------------------------------------------------
    def poll_commands(self, handlers: dict | None = None) -> int:
        """Run every queued command through ``handlers[name](args)``. Returns
        the number of commands processed. Unknown names get an error result so
        the caller never waits on a command that can't be served."""
        handlers = handlers or {}
        try:
            names = sorted(f for f in os.listdir(self.cmd_dir) if f.endswith(".json"))
        except OSError:
            return 0
        n = 0
        for fn in names:
            path = os.path.join(self.cmd_dir, fn)
            try:
                with open(path, encoding="utf-8") as f:
                    cmd = json.load(f)
            except (OSError, ValueError):
                # half-written or unreadable: leave it for the next tick, but
                # never forever -- drop after 60s
                try:
                    if time.time() - os.path.getmtime(path) > 60:
                        os.remove(path)
                except OSError:
                    pass
                continue
            cid = str(cmd.get("id") or os.path.splitext(fn)[0])
            name = str(cmd.get("name") or "")
            args = cmd.get("args") or {}
            fnc = handlers.get(name)
            if fnc is None:
                result = {"ok": False, "error": f"MaINbox doesn't know the command {name!r}"}
            else:
                try:
                    out = fnc(args)
                    if isinstance(out, dict):
                        result = dict(out)
                        result.setdefault("ok", True)
                    else:
                        result = {"ok": True, "result": out}
                except Exception as e:  # noqa: BLE001
                    result = {"ok": False, "error": f"{type(e).__name__}: {e}",
                              "trace": traceback.format_exc()[-1500:]}
            # the handler's own "id" (e.g. the follow-up id) must survive;
            # the command id goes under cmd_id
            result.setdefault("id", cid)
            result["cmd_id"] = cid
            result["name"] = name
            result["done_ts"] = time.time()
            try:
                _atomic_write(os.path.join(self.res_dir, cid + ".json"), result)
            except OSError:
                pass
            try:
                os.remove(path)
            except OSError:
                pass
            n += 1
        self._prune_results()
        return n

    def _prune_results(self) -> None:
        now = time.time()
        if now - self._last_prune < 300:
            return
        self._last_prune = now
        try:
            for fn in os.listdir(self.res_dir):
                p = os.path.join(self.res_dir, fn)
                try:
                    if now - os.path.getmtime(p) > _RESULT_TTL_S:
                        os.remove(p)
                except OSError:
                    pass
        except OSError:
            pass
