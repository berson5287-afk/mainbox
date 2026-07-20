# ============================================================
# crash_guard.py  --  silent-exit / crash capture for MaINbox
#
# WHY THIS EXISTS
# MaINbox runs as a .pyw, which has NO console. That means sys.stderr /
# sys.stdout go nowhere, so any traceback Python prints when it dies is
# discarded -- the app just "closes by itself" with nothing to look at.
# Worse, an Outlook/COM access violation is a *native* crash: it is NOT a
# Python exception, so no amount of try/except can see it and the process
# vanishes instantly. The only tool that records native crashes from Python
# is faulthandler, which MaINbox does not currently use.
#
# This module installs, in one call, every capture that was missing:
#   * faulthandler            -> writes the Python stack of ALL threads at the
#                                moment of a native fault (access violation,
#                                stack overflow, etc.). This is what finally
#                                names a COM/Outlook crash.
#   * sys.excepthook          -> logs uncaught main-thread exceptions.
#   * threading.excepthook    -> logs uncaught background-thread exceptions.
#   * stderr/stdout redirect  -> routes anything Python would have printed
#                                (incl. your traceback.print_exc) into the log.
#   * Tk report_callback_exception (optional, via patch_tk_class) -> logs Tk
#                                callback errors instead of dropping them.
#   * atexit marker           -> writes "CLEAN EXIT" on a graceful shutdown.
#                                If the log just STOPS with no such marker, the
#                                death was a hard crash (faulthandler dump above
#                                it tells you where).
#   * hang detector (v1.1, via start_hang_detector) -> for a UI freeze that
#                                RECOVERS (locks up, then comes back), which
#                                faulthandler never sees because nothing died:
#                                dumps ALL thread stacks when the Tk main thread
#                                stops processing for > N seconds, then logs how
#                                long the freeze lasted. Distinguishes a real UI
#                                hang from a laptop-sleep / whole-process suspend.
#
# IMPORTANT: this CAPTURES crashes, it does not PREVENT them. On a native
# fault the process still dies; you just finally get a log. (Use the separate
# mainbox_watchdog.py if you also want auto-relaunch + the OS exit code.)
#
# INTEGRATION  (one line, as early as possible -- right after APP_DATA_DIR is
# computed, around line 463 in MaINbox):
#
#     import crash_guard; crash_guard.install(APP_DATA_DIR)
#
# and once, anywhere before windows are built, to also catch Tk callbacks
# across every tk.Tk() root the app creates:
#
#     crash_guard.patch_tk_class()
#
# and once, right after the long-lived MAIN Tk root is built and just before
# its root.mainloop(), to also capture UI freezes that recover:
#
#     crash_guard.start_hang_detector(main_root)
#
# Pure standard library. install() is idempotent and never raises.
# ============================================================

import os
import sys
import time
import atexit
import threading
import traceback
import datetime
import platform

try:
    import faulthandler
except Exception:
    faulthandler = None

CRASH_GUARD_VERSION = "v1.1"

_installed = False
_lock = threading.Lock()
_LOG_FH = None            # real file object, kept alive for faulthandler's fd
_LOG_PATH = None
_orig_stdout = None
_orig_stderr = None
_orig_excepthook = None
_orig_threadhook = None
_hb_thread = None
_hb_stop = None

# --- main/UI-thread hang detector state (v1.1) ---
_hangd_started = False
_hangd_stop_evt = None
_hangd_state = {
    "root": None,
    "last_beat": 0.0,       # time.monotonic() of the last UI-thread heartbeat
    "threshold": 10.0,      # seconds of UI silence before we call it a hang
    "check_interval": 2.0,  # watchdog poll period
    "heartbeat": 1.0,       # UI-thread heartbeat reschedule period
    "redump": 30.0,         # re-dump cadence while a hang persists
    "stopped": False,       # root destroyed / shutting down -> stop reporting
    "hanging": False,
    "hang_anchor": 0.0,
    "last_dump": 0.0,
}


def _now():
    return datetime.datetime.now().isoformat(sep=" ", timespec="milliseconds")


def _raw_write(text):
    """Write straight to the crash log and flush. Never raises."""
    fh = _LOG_FH
    if fh is None:
        return
    try:
        fh.write(text)
        fh.flush()
    except Exception:
        pass


class _Tee:
    """Forwards writes to the original stream (if any) AND the crash log.

    In a .pyw the original stream is usually None; teeing still works because
    we always have the log file. Swallows every error so logging can never be
    the thing that crashes the app.
    """

    def __init__(self, original):
        self._original = original

    def write(self, s):
        try:
            if self._original is not None:
                self._original.write(s)
                self._original.flush()
        except Exception:
            pass
        try:
            if _LOG_FH is not None:
                _LOG_FH.write(s)
                _LOG_FH.flush()
        except Exception:
            pass
        return len(s) if s else 0

    def flush(self):
        for t in (self._original, _LOG_FH):
            try:
                if t is not None:
                    t.flush()
            except Exception:
                pass

    def isatty(self):
        try:
            return bool(self._original is not None and self._original.isatty())
        except Exception:
            return False

    def fileno(self):
        # Some libraries probe fileno(); hand back the real log fd so they have
        # a valid descriptor (in a .pyw the original is None and has none).
        if _LOG_FH is not None:
            return _LOG_FH.fileno()
        raise OSError("crash_guard tee has no fileno")


def _excepthook(exc_type, exc_value, exc_tb):
    try:
        _raw_write(
            "\n" + "=" * 72 + "\n"
            + f"[{_now()}] UNHANDLED EXCEPTION (main thread)\n"
            + "=" * 72 + "\n"
        )
        _raw_write("".join(traceback.format_exception(exc_type, exc_value, exc_tb)))
        _raw_write("=" * 72 + "\n")
    except Exception:
        pass
    try:
        if _orig_excepthook is not None:
            _orig_excepthook(exc_type, exc_value, exc_tb)
    except Exception:
        pass


def _threadhook(args):
    # Mirror the stdlib default: a thread that exits via SystemExit is normal.
    if getattr(args, "exc_type", None) is SystemExit:
        return
    try:
        tname = getattr(getattr(args, "thread", None), "name", "?")
        _raw_write(
            "\n" + "-" * 72 + "\n"
            + f"[{_now()}] UNHANDLED EXCEPTION (thread: {tname})\n"
            + "-" * 72 + "\n"
        )
        _raw_write("".join(traceback.format_exception(
            args.exc_type, args.exc_value, args.exc_traceback)))
        _raw_write("-" * 72 + "\n")
    except Exception:
        pass
    try:
        if _orig_threadhook is not None:
            _orig_threadhook(args)
    except Exception:
        pass


def _tk_log_callback_exception(exc_type, exc_value, exc_tb):
    try:
        _raw_write(
            "\n" + "." * 72 + "\n"
            + f"[{_now()}] TK CALLBACK EXCEPTION\n"
            + "." * 72 + "\n"
        )
        _raw_write("".join(traceback.format_exception(exc_type, exc_value, exc_tb)))
        _raw_write("." * 72 + "\n")
    except Exception:
        pass


def patch_tk_class(show_dialog=False):
    """Route Tk callback exceptions for EVERY tk.Tk() root to the crash log.

    MaINbox creates several roots over its lifetime (picker -> main -> picker);
    patching the class covers all of them with one call. By default this only
    logs. Pass show_dialog=True while hunting a bug to also pop a messagebox so
    you SEE the error the instant it happens.
    """
    try:
        import tkinter
    except Exception:
        return

    def handler(self, exc_type, exc_value, exc_tb):
        _tk_log_callback_exception(exc_type, exc_value, exc_tb)
        if show_dialog:
            try:
                from tkinter import messagebox
                messagebox.showerror(
                    "MaINbox error (logged)",
                    f"{getattr(exc_type, '__name__', exc_type)}: {exc_value}\n\n"
                    f"Full traceback written to:\n{_LOG_PATH}",
                )
            except Exception:
                pass

    try:
        tkinter.Tk.report_callback_exception = handler
        _raw_write(f"[{_now()}] Tk callback-exception capture installed "
                   f"(show_dialog={show_dialog})\n")
    except Exception:
        pass


def _on_exit():
    # NOTE: atexit handlers DO run on a normal/exception shutdown, but do NOT
    # run on a hard native crash (access violation) or os._exit(). So:
    #   marker present  -> graceful shutdown.
    #   marker ABSENT (log just stops) -> hard crash; look just above for the
    #                                     faulthandler dump.
    try:
        _raw_write(f"\n[{_now()}] CLEAN EXIT (atexit) pid={os.getpid()}\n")
    except Exception:
        pass


def start_heartbeat(interval_seconds=30):
    """Optional: write a 'heartbeat' line every N seconds.

    Lets you bound the time of death -- the last heartbeat before the log stops
    tells you the process was alive until ~then. Off by default (chatty); turn
    it on while actively reproducing the crash.
    """
    global _hb_thread, _hb_stop
    if _hb_thread is not None:
        return
    _hb_stop = threading.Event()

    def _loop():
        while not _hb_stop.wait(interval_seconds):
            _raw_write(f"[{_now()}] heartbeat pid={os.getpid()}\n")

    _hb_thread = threading.Thread(target=_loop, name="crash-guard-heartbeat",
                                  daemon=True)
    _hb_thread.start()
    _raw_write(f"[{_now()}] heartbeat started (every {interval_seconds}s)\n")


# ============================================================
# Main/UI-thread HANG detector (v1.1)
#
# faulthandler only fires when the process DIES. A UI freeze that recovers
# (the app "locks up for 30 seconds" then comes back) leaves no trace. This
# detector catches exactly that: a stamp refreshed on the Tk main thread (so it
# only advances while the event loop is alive) plus a daemon watchdog that, when
# the stamp goes stale, dumps EVERY thread's stack into the crash log -- naming
# what the UI is stuck on AND what each worker is doing (e.g. holding a lock).
# It logs how long each freeze lasted, and does NOT misreport a laptop-sleep
# (whole-process suspend) as a UI hang.
# ============================================================

def _hangd_dump(reason):
    """Write a labeled all-thread stack dump into the crash log. Never raises."""
    try:
        _raw_write(
            "\n" + "!" * 72 + "\n"
            + f"[{_now()}] {reason}\n"
            + "    UI event loop has not run -- stacks of ALL threads follow. The\n"
            + "    main-thread frame is what the UI is stuck on; a worker frame in a\n"
            + "    COM/Outlook call or a lock.acquire() is usually the cause.\n"
            + "!" * 72 + "\n"
        )
    except Exception:
        pass
    # Prefer faulthandler: it marks the current thread and walks every thread.
    if faulthandler is not None and _LOG_FH is not None:
        try:
            faulthandler.dump_traceback(file=_LOG_FH, all_threads=True)
            try:
                _LOG_FH.flush()
            except Exception:
                pass
            _raw_write("!" * 72 + "\n")
            return
        except Exception:
            pass
    # Fallback: manual stacks of every live thread.
    try:
        frames = sys._current_frames()
        names = {}
        try:
            for t in threading.enumerate():
                names[t.ident] = t.name
        except Exception:
            pass
        try:
            main_id = threading.main_thread().ident
        except Exception:
            main_id = None
        for tid, frame in frames.items():
            tag = " (MAIN/UI THREAD)" if tid == main_id else ""
            _raw_write(f"\n--- Thread {names.get(tid, '?')}{tag} (id={tid}) ---\n")
            try:
                _raw_write("".join(traceback.format_stack(frame)))
            except Exception:
                pass
        _raw_write("!" * 72 + "\n")
    except Exception:
        pass


def _hangd_heartbeat():
    """Runs ON the Tk main thread via root.after; only fires while the event
    loop is processing, so a blocked UI thread stops refreshing the stamp."""
    try:
        _hangd_state["last_beat"] = time.monotonic()
    except Exception:
        pass
    root = _hangd_state.get("root")
    if root is None or _hangd_state.get("stopped"):
        return
    try:
        root.after(int(_hangd_state["heartbeat"] * 1000), _hangd_heartbeat)
    except Exception:
        # Root destroyed (window closed / logout). Stop quietly so the watchdog
        # does not misread shutdown as a freeze.
        _hangd_state["stopped"] = True


def _hangd_watchdog():
    wd_prev = time.monotonic()
    while not (_hangd_stop_evt is not None and _hangd_stop_evt.wait(_hangd_state["check_interval"])):
        now = time.monotonic()
        wd_gap = now - wd_prev
        wd_prev = now
        try:
            threshold = float(_hangd_state["threshold"])
        except Exception:
            threshold = 10.0
        if _hangd_state.get("stopped"):
            _hangd_state["hanging"] = False
            continue
        last = _hangd_state.get("last_beat", 0.0)
        if last <= 0:
            continue
        # Whole-process suspend guard: if the watchdog THREAD itself lost as much
        # wall time as our hang threshold, the entire process was frozen (laptop
        # sleep, severe swap) -- the stale heartbeat is expected, not a UI hang.
        if wd_gap >= threshold:
            _hangd_state["hanging"] = False
            continue
        delta = now - last
        if delta >= threshold:
            if not _hangd_state.get("hanging"):
                _hangd_state["hanging"] = True
                _hangd_state["hang_anchor"] = last
                _hangd_state["last_dump"] = now
                _hangd_dump(f"MAIN/UI THREAD UNRESPONSIVE for ~{delta:.0f}s "
                            f"(no UI event processed since last heartbeat)")
            else:
                try:
                    redump = float(_hangd_state["redump"])
                except Exception:
                    redump = 30.0
                if (now - _hangd_state.get("last_dump", now)) >= redump:
                    _hangd_state["last_dump"] = now
                    stuck = now - _hangd_state.get("hang_anchor", now)
                    _hangd_dump(f"MAIN/UI THREAD STILL UNRESPONSIVE for ~{stuck:.0f}s")
        else:
            if _hangd_state.get("hanging"):
                _hangd_state["hanging"] = False
                try:
                    dur = max(0.0, last - _hangd_state.get("hang_anchor", last))
                except Exception:
                    dur = 0.0
                _raw_write(f"\n[{_now()}] >>> UI THREAD RESPONSIVE AGAIN after "
                           f"~{dur:.1f}s freeze (pid={os.getpid()}) <<<\n")


def start_hang_detector(root, threshold_seconds=10.0, check_interval_seconds=2.0,
                        heartbeat_seconds=1.0, redump_interval_seconds=30.0):
    """Detect and record main/UI-thread freezes -- hangs that RECOVER, which
    faulthandler never sees because the process does not die.

    A tiny stamp is refreshed on the Tk main thread via root.after (so it only
    advances while the event loop is alive). A daemon watchdog dumps EVERY
    thread's stack to the crash log if that stamp goes stale for longer than
    threshold_seconds -- naming exactly what the UI is stuck on, and what any
    worker thread is doing (e.g. a COM call or a lock.acquire()). When the UI
    recovers, it logs how long the freeze lasted. A whole-process suspend
    (laptop sleep) is distinguished from a real UI hang and is not reported.

    Call once, AFTER the long-lived main Tk root exists (right before
    root.mainloop()). Safe to call again for a replacement root; never raises.
    """
    global _hangd_started, _hangd_stop_evt
    try:
        _hangd_state["root"] = root
        _hangd_state["threshold"] = float(threshold_seconds)
        _hangd_state["check_interval"] = float(check_interval_seconds)
        _hangd_state["heartbeat"] = float(heartbeat_seconds)
        _hangd_state["redump"] = float(redump_interval_seconds)
        _hangd_state["stopped"] = False
        _hangd_state["hanging"] = False
        _hangd_state["last_beat"] = time.monotonic()
    except Exception:
        return where()
    # (Re)arm the heartbeat on this root's event loop.
    try:
        root.after(int(_hangd_state["heartbeat"] * 1000), _hangd_heartbeat)
    except Exception:
        pass
    # Start the watchdog thread exactly once.
    if not _hangd_started:
        try:
            _hangd_stop_evt = threading.Event()
            t = threading.Thread(target=_hangd_watchdog,
                                 name="mainbox-hang-detector", daemon=True)
            t.start()
            _hangd_started = True
            _raw_write(f"[{_now()}] hang detector ON (UI-freeze threshold "
                       f"{float(threshold_seconds):.0f}s; on freeze dumps all thread "
                       f"stacks, logs recovery + duration)\n")
        except Exception as e:
            try:
                _raw_write(f"[{_now()}] hang detector start FAILED: {e!r}\n")
            except Exception:
                pass
    return where()


def log(msg):
    """Public breadcrumb helper. Drop crash_guard.log('about to scan inbox')
    anywhere to narrow down where death occurs relative to your own steps."""
    _raw_write(f"[{_now()}] {msg}\n")


def where():
    """Return the crash log path (so you can show it in an About box, etc.)."""
    return _LOG_PATH


def install(log_dir=None, *, redirect_std=True, enable_faulthandler=True,
            install_excepthooks=True, heartbeat=False, heartbeat_seconds=30):
    """Install all capture hooks. Idempotent. Returns the crash log path.

    log_dir: where to write mainbox_crash.log. Pass APP_DATA_DIR so the log
             lands next to coverage_debug.log. If omitted, defaults to this
             file's folder, then the temp dir as a last resort.
    """
    global _installed, _LOG_FH, _LOG_PATH, _orig_stdout, _orig_stderr
    global _orig_excepthook, _orig_threadhook

    with _lock:
        if _installed:
            return _LOG_PATH

        if not log_dir:
            try:
                log_dir = os.path.dirname(os.path.abspath(__file__))
            except Exception:
                log_dir = os.getcwd()
        try:
            os.makedirs(log_dir, exist_ok=True)
        except Exception:
            pass

        _LOG_PATH = os.path.join(log_dir, "mainbox_crash.log")
        try:
            # Line-buffered text append; real fd so faulthandler can write to it.
            _LOG_FH = open(_LOG_PATH, "a", buffering=1, encoding="utf-8",
                           errors="replace")
        except Exception:
            import tempfile
            _LOG_PATH = os.path.join(tempfile.gettempdir(), "mainbox_crash.log")
            _LOG_FH = open(_LOG_PATH, "a", buffering=1, encoding="utf-8",
                           errors="replace")

        _raw_write(
            "\n" + "#" * 72 + "\n"
            + f"# MaINbox crash_guard {CRASH_GUARD_VERSION}  --  session start {_now()}\n"
            + f"# pid={os.getpid()}  python={platform.python_version()}  "
            + f"{platform.system()} {platform.release()} ({platform.machine()})\n"
            + f"# executable={sys.executable}\n"
            + f"# argv={sys.argv}\n"
            + "#" * 72 + "\n"
        )

        if enable_faulthandler and faulthandler is not None:
            try:
                faulthandler.enable(file=_LOG_FH, all_threads=True)
                _raw_write(f"[{_now()}] faulthandler enabled "
                           f"(NATIVE-crash capture ON -- catches COM/Outlook "
                           f"access violations)\n")
            except Exception as e:
                _raw_write(f"[{_now()}] faulthandler enable FAILED: {e!r}\n")
        elif faulthandler is None:
            _raw_write(f"[{_now()}] faulthandler unavailable on this Python; "
                       f"native crashes will NOT be captured\n")

        if redirect_std:
            _orig_stdout = sys.stdout
            _orig_stderr = sys.stderr
            try:
                sys.stdout = _Tee(_orig_stdout)
                sys.stderr = _Tee(_orig_stderr)
                _raw_write(f"[{_now()}] stdout/stderr redirected into this log\n")
            except Exception:
                pass

        if install_excepthooks:
            _orig_excepthook = sys.excepthook
            sys.excepthook = _excepthook
            if hasattr(threading, "excepthook"):
                _orig_threadhook = threading.excepthook
                threading.excepthook = _threadhook
            _raw_write(f"[{_now()}] sys/threading excepthooks installed\n")

        atexit.register(_on_exit)

        if heartbeat:
            start_heartbeat(heartbeat_seconds)

        _installed = True
        _raw_write(f"[{_now()}] crash_guard ready. Log: {_LOG_PATH}\n")
        return _LOG_PATH


if __name__ == "__main__":
    # Tiny self-test: install, then deliberately crash to prove the log works.
    p = install()
    print("crash_guard installed. Log at:", p)
    print("Raising a test exception now (should be captured, then re-raised)...")
    raise RuntimeError("crash_guard self-test exception -- this is expected")
