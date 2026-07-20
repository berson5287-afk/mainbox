# ============================================================
# mainbox_watchdog.py  --  external supervisor for MaINbox
#
# This is the OPTIONAL second half of crash hunting. crash_guard.py captures
# the crash from INSIDE the process; this watchdog watches from OUTSIDE.
#
# It launches MaINbox as a child process, waits for it to exit, and logs a
# timestamped line with the OS EXIT CODE -- which independently confirms a
# native crash. On Windows an unhandled access violation makes the process exit
# with 0xC0000005 (3221225477). Seeing that code is hard proof that the death
# was a native fault (Outlook/COM), not a Python exception.
#
# It can also relaunch MaINbox automatically after a crash, with a per-hour cap
# so a crash-on-startup cannot spin forever.
#
# RUN IT FROM A CONSOLE so you can watch the status live:
#     python mainbox_watchdog.py
# or point it explicitly at the .pyw:
#     python mainbox_watchdog.py "C:\path\to\MaINbox_v3_9_83_AI_Assistant.pyw"
#
# Pure standard library.
# ============================================================

import os
import sys
import time
import datetime
import subprocess

WATCHDOG_VERSION = "v1.0"

# ---- config -------------------------------------------------------------
TARGET = ""                  # path to the MaINbox .pyw. Empty -> auto-detect
                             # a *.pyw with "mainbox" in the name next to this
                             # file, or take it from the command line argument.
AUTO_RESTART = True          # relaunch MaINbox after a non-clean exit
RESTART_DELAY_SECONDS = 3
MAX_RESTARTS_PER_HOUR = 6    # safety cap against a crash-on-startup loop
# -------------------------------------------------------------------------

# Common Windows process exit codes worth translating on sight.
_EXIT_MEANINGS = {
    0:           "clean exit (normal shutdown / window closed)",
    1:           "Python exited with error (uncaught exception reached top level)",
    3221225477:  "0xC0000005 ACCESS VIOLATION -- NATIVE crash (Outlook/COM/native lib). "
                 "Not catchable by Python try/except. This is the classic 'just closed' cause.",
    3221225725:  "0xC00000FD STACK OVERFLOW -- runaway recursion or deep native call chain.",
    3221226505:  "0xC0000409 STACK BUFFER OVERRUN / native fast-fail.",
    3221225786:  "0xC000013A -- terminated by Ctrl+C / console close.",
    1073807364:  "0x40010004 DBG_TERMINATE_PROCESS -- process was killed (e.g. Task Manager).",
}


def describe(code):
    if code in _EXIT_MEANINGS:
        return _EXIT_MEANINGS[code]
    u = code & 0xFFFFFFFF          # normalize a possibly-signed value
    if u in _EXIT_MEANINGS:
        return _EXIT_MEANINGS[u]
    return f"abnormal exit (code {code} / 0x{u:08X})"


def python_launcher():
    """Prefer pythonw.exe so the child stays windowless like a real .pyw run."""
    exe = sys.executable or "python"
    d = os.path.dirname(exe)
    pyw = os.path.join(d, "pythonw.exe")
    return pyw if os.path.exists(pyw) else exe


def find_target():
    if len(sys.argv) > 1 and sys.argv[1].strip():
        return os.path.abspath(sys.argv[1].strip())
    if TARGET.strip():
        return os.path.abspath(TARGET.strip())
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        cands = [f for f in os.listdir(here)
                 if f.lower().endswith(".pyw") and "mainbox" in f.lower()]
        cands.sort()                       # v3_9_83 sorts after v3_9_8, etc.
        if cands:
            return os.path.join(here, cands[-1])
    except Exception:
        pass
    return ""


def main():
    target = find_target()
    if not target or not os.path.exists(target):
        print("Watchdog: could not locate the MaINbox .pyw.")
        print("Pass the path explicitly:")
        print('    python mainbox_watchdog.py "C:\\path\\to\\'
              'MaINbox_v3_9_83_AI_Assistant.pyw"')
        return

    log_dir = os.path.dirname(target) or os.getcwd()
    log_path = os.path.join(log_dir, "mainbox_watchdog.log")
    launcher = python_launcher()

    def log(line):
        stamp = datetime.datetime.now().isoformat(sep=" ", timespec="seconds")
        msg = f"[{stamp}] {line}"
        try:
            print(msg, flush=True)
        except Exception:
            pass
        try:
            with open(log_path, "a", encoding="utf-8") as fh:
                fh.write(msg + "\n")
        except Exception:
            pass

    log(f"watchdog {WATCHDOG_VERSION} starting")
    log(f"target   = {target}")
    log(f"launcher = {launcher}")
    log(f"log file = {log_path}")
    log(f"auto-restart = {AUTO_RESTART}  (max {MAX_RESTARTS_PER_HOUR}/hour)")

    restart_times = []
    run = 0
    while True:
        run += 1
        log(f"--- launching MaINbox (run #{run}) ---")
        start = time.time()
        try:
            proc = subprocess.Popen([launcher, target], cwd=log_dir)
        except Exception as e:
            log(f"FAILED to launch MaINbox: {e!r}")
            return

        proc.wait()
        elapsed = time.time() - start
        code = proc.returncode
        log(f"MaINbox exited after {elapsed:.0f}s -- exit code {code} -> "
            f"{describe(code)}")

        if code == 0:
            log("clean exit; watchdog stopping.")
            return

        if not AUTO_RESTART:
            log("AUTO_RESTART is off; watchdog stopping. "
                "Check mainbox_crash.log for the in-process traceback.")
            return

        now = time.time()
        restart_times = [t for t in restart_times if now - t < 3600]
        if len(restart_times) >= MAX_RESTARTS_PER_HOUR:
            log(f"hit MAX_RESTARTS_PER_HOUR ({MAX_RESTARTS_PER_HOUR}); stopping "
                f"to avoid a crash loop. The cause is in mainbox_crash.log.")
            return
        restart_times.append(now)

        log(f"restarting in {RESTART_DELAY_SECONDS}s...")
        time.sleep(RESTART_DELAY_SECONDS)


if __name__ == "__main__":
    main()
