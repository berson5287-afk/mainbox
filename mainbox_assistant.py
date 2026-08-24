# ============================================================
# MaINbox - AI Assistant Engine  (experimental, toggleable)
# Copyright (c) 2026 Stephen Berson. All Rights Reserved.
#
# A virtual-assistant layer that sits ON TOP of MaINbox's existing
# triage / follow-up / reminder decisions. It does not replace any
# scan logic. It only decides, per action:
#
#     AUTO    -> high confidence: do it automatically, just log it
#     REVIEW  -> medium confidence (low-stakes only): do it, flag it
#     ASK      -> low confidence OR high-stakes medium: surface to the
#                user instead of acting (this "bypasses" silent mode)
#
# It ships with:
#   - AssistantConfig    : thresholds + on/off + log-visible flags
#   - AssistantEngine    : the brain + a persisted activity log
#   - AssistantLogWindow : a LIVE change log (Tk) the user can review,
#                          undo, and correct; can be hidden (silent)
#   - a built-in demo     : run this file directly to see the UX with
#                          fake emails (no Outlook / Ollama needed)
#
# Integrating into MaINbox is additive and surgical -- see
# mainbox_assistant_integration.md.
# ============================================================

import json
import os
import uuid
import time
from datetime import datetime, timedelta

# -- Decision outcomes ---------------------------------------------------------
AUTO   = "auto"     # confident enough -> execute automatically
REVIEW = "review"   # borderline, low-stakes -> execute but flag for a glance
ASK    = "ask"      # not confident, or high-stakes borderline -> ask the user

# -- Action stakes --------------------------------------------------------------
# Low-stakes actions are cheap to undo (relabel, regroup). High-stakes actions
# touch the outside world or are hard to reverse (send/deploy a draft, delete,
# close a thread) and therefore demand a higher bar before going fully automatic.
LOW  = "low"
HIGH = "high"


class AssistantConfig:
    """All thresholds in one place. Defaults are deliberately conservative:
    the assistant only acts silently when it is genuinely sure, and always
    falls back to asking when it is not."""

    def __init__(self, **overrides):
        # Master switches
        self.enabled = True            # is the assistant doing anything at all?
        self.show_log = True           # is the live activity log window shown?
                                       #   False = "silent" mode. Uncertain (ASK)
                                       #   items STILL surface to the user.

        # Low-stakes thresholds (category / urgency / status / grouping)
        self.low_auto_at = 0.75        # >= this -> AUTO
        self.low_ask_below = 0.55      # <  this -> ASK ; between -> REVIEW

        # Hard auto-line: when True, the auto-file bar is the ONLY line that acts.
        # Anything below it is HELD for the user's one-click Approve instead of the
        # old "do it, but flag it amber" middle band. This makes the single visible
        # slider mean exactly what it says: nothing is set up below the bar.
        self.hard_auto_line = True

        # High-stakes thresholds (deploy a draft, delete, close a thread)
        self.high_auto_at = 0.88       # >= this -> AUTO
        self.high_ask_below = 0.65     # <  this -> ASK ; between -> ASK (staged)

        # If False, even high-stakes actions never go fully automatic; the most
        # the assistant will do is stage them for one-click approval.
        self.allow_auto_high_stakes = True

        # Auto-send safety: when an auto-send WOULD fire, hold it for a short
        # countdown (default 2 min) with a Cancel, so the user can edit/stop it
        # before it actually goes out. delay_enabled toggles the hold.
        self.autosend_delay_enabled = True
        self.autosend_delay_seconds = 120

        # Auto-archive: when ON, a detected customer cancellation ("no longer
        # needed") or a no-quote decision archives the group on its own instead of
        # waiting for the user's Approve. Default OFF (the safe, ask-first behavior).
        self.autoarchive_enabled = False

        for k, v in overrides.items():
            if hasattr(self, k):
                setattr(self, k, v)

    # ---- mapping to/from MaINbox's existing settings dict -----------------------
    @classmethod
    def from_mainbox_settings(cls, settings):
        """Build config from MaINbox's settings dict. Falls back to defaults for
        any key that isn't present, so existing installs keep working."""
        s = settings or {}
        return cls(
            enabled=bool(s.get("assistant_enabled", True)),
            show_log=bool(s.get("assistant_show_log", True)),
            low_auto_at=float(s.get("assistant_low_auto_at", 0.75)),
            low_ask_below=float(s.get("assistant_low_ask_below", 0.55)),
            hard_auto_line=bool(s.get("assistant_hard_auto_line", True)),
            high_auto_at=float(s.get("assistant_high_auto_at", 0.88)),
            high_ask_below=float(s.get("assistant_high_ask_below", 0.65)),
            allow_auto_high_stakes=bool(s.get("assistant_allow_auto_high_stakes", True)),
            autosend_delay_enabled=bool(s.get("assistant_autosend_delay_enabled", True)),
            autosend_delay_seconds=int(s.get("assistant_autosend_delay_seconds", 120)),
            autoarchive_enabled=bool(s.get("assistant_autoarchive_enabled", False)),
        )

    def to_mainbox_settings(self):
        return {
            "assistant_enabled": self.enabled,
            "assistant_show_log": self.show_log,
            "assistant_low_auto_at": self.low_auto_at,
            "assistant_low_ask_below": self.low_ask_below,
            "assistant_hard_auto_line": self.hard_auto_line,
            "assistant_high_auto_at": self.high_auto_at,
            "assistant_high_ask_below": self.high_ask_below,
            "assistant_allow_auto_high_stakes": self.allow_auto_high_stakes,
            "assistant_autosend_delay_enabled": self.autosend_delay_enabled,
            "assistant_autosend_delay_seconds": self.autosend_delay_seconds,
            "assistant_autoarchive_enabled": self.autoarchive_enabled,
        }


# -- Confidence helpers -------------------------------------------------------
def confidence_from_triage(result):
    """Collapse a MaINbox triage/quick-intent result dict into a single 0..1
    confidence. Uses the SAME fields the app already produces:
    business_relevance, spam_likelihood, and (if present) an explicit model
    confidence. No new AI calls."""
    if not isinstance(result, dict):
        return 0.0
    # NB: use an explicit None check, not `or`, so a legitimate 0.0
    # (e.g. spam_likelihood == 0.0 meaning "definitely not spam") is kept.
    def _num(key, default):
        v = result.get(key, default)
        if v is None:
            v = default
        try:
            return float(v)
        except Exception:
            return default
    br = _num("business_relevance", 0.0)   # missing -> no evidence -> low
    sl = _num("spam_likelihood", 1.0)      # missing -> assume spammy (safe)
    base = max(0.0, min(1.0, br)) * (1.0 - max(0.0, min(1.0, sl)))
    # Blend in an explicit confidence if the model gave one.
    for key in ("confidence", "triage_confidence"):
        if key in result:
            try:
                mc = float(result.get(key) or 0.0)
                return max(0.0, min(1.0, 0.5 * base + 0.5 * mc))
            except Exception:
                pass
    return max(0.0, min(1.0, base))


def confidence_from_group(result):
    """Grouping confidence comes straight from the model's group_confidence."""
    if not isinstance(result, dict):
        return 0.0
    try:
        return max(0.0, min(1.0, float(result.get("group_confidence", 0.0) or 0.0)))
    except Exception:
        return 0.0


# -- The engine ---------------------------------------------------------------
class AssistantEngine:
    """UI-agnostic brain. You give it callbacks; it decides and records.

    Callbacks (all optional):
        do_apply(action)  -> actually perform a low/high-stakes action that the
                             engine decided to AUTO/REVIEW. Return True on success.
        do_ask(action)    -> surface an action to the user (prompt / staged
                             approval). The host decides how.
        on_log_change()   -> notify the UI that the activity list changed.
        do_undo(action)   -> reverse an applied action. Return True on success.

    In MaINbox these map to: do_apply = "write the label/group/draft",
    do_ask = "show the existing prompt or stage it", do_undo = "restore the
    previous category/urgency/status/group" (you already snapshot these)."""

    def __init__(self, config=None, log_path=None, config_path=None,
                 do_apply=None, do_ask=None, do_undo=None, on_log_change=None,
                 on_activity=None):
        self.config = config or AssistantConfig()
        self.config_path = config_path
        self.log_path = log_path
        self.do_apply = do_apply
        self.do_ask = do_ask
        self.do_undo = do_undo
        self.on_log_change = on_log_change
        # Fires once per NEW activity row (not on refresh/approve/undo). The host
        # uses it to auto-open the log window per the "Show this log" preference.
        self.on_activity = on_activity
        self.on_row_activate = None   # host sets this: double-click a log row -> edit/manage it
        self.open_email_cb = None     # host sets this: "Open email" jumps to the row in the app
        self.activity = []          # newest-first list of action records
        self._load_config()
        self._load()

    # ---- persistence ------------------------------------------------------------
    def _load_config(self):
        """If a saved config file exists, it remembers the user's last slider /
        toggle positions and overrides the defaults passed in from settings."""
        if not self.config_path:
            return
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    d = json.load(f)
                if isinstance(d, dict):
                    self.config = AssistantConfig.from_mainbox_settings(d)
        except Exception:
            pass

    def save_config(self):
        if not self.config_path:
            return
        try:
            tmp = self.config_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.config.to_mainbox_settings(), f, indent=2)
            os.replace(tmp, self.config_path)
        except Exception:
            pass

    def _load(self):
        if not self.log_path:
            return
        try:
            if os.path.exists(self.log_path):
                with open(self.log_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, list):
                    self.activity = data
        except Exception:
            self.activity = []
        # Heal any legacy data written before ids were collision-proof: drop
        # records whose id duplicates an earlier one, and backfill missing ids.
        seen, clean = set(), []
        for r in self.activity:
            if not isinstance(r, dict):
                continue
            rid = r.get("id")
            if not rid or rid in seen:
                rid = f"act_{datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:8]}"
                r["id"] = rid
            seen.add(rid)
            clean.append(r)
        self.activity = clean

    def _save(self):
        if not self.log_path:
            return
        try:
            tmp = self.log_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.activity[:2000], f, indent=2)
            os.replace(tmp, self.log_path)
        except Exception:
            pass

    # ---- core decision ----------------------------------------------------------
    def classify(self, confidence, stakes):
        """Return AUTO / REVIEW / ASK for a given confidence + stakes."""
        c = self.config
        if stakes == HIGH:
            if c.allow_auto_high_stakes and confidence >= c.high_auto_at:
                return AUTO
            if confidence >= c.high_ask_below:
                return ASK        # staged: prepared, but the user confirms
            return ASK
        # low stakes
        if confidence >= c.low_auto_at:
            return AUTO
        if confidence >= c.low_ask_below:
            return REVIEW         # do it, but flag it for an easy glance
        return ASK

    # States that still represent a "live" item the user hasn't resolved yet.
    # A new handle() with the same dedup_key as one of these is a duplicate and
    # is collapsed onto the existing record instead of stacking a new row.
    #
    # v4.2.82: "rejected" belongs here too. Dismiss sets state="rejected", and
    # with rejected NOT counted as live, the next re-triage pass (every ~15 min)
    # found no live record for the same dedup_key and minted a brand-new pending
    # ask -- so a dismissed follow-up came back forever (reported live: the
    # Dennis McCloskey "Re: Material list 6531" ask). The collapse branch in
    # handle() deliberately never changes an existing record's state, so
    # re-scans now land silently on the dismissed row (repeat_count ticks) and
    # Dismiss finally means dismissed.
    _LIVE_STATES = ("pending", "auto", "flagged", "corrected", "armed", "rejected")

    def _find_live_by_dedup(self, dedup_key):
        """Return the most recent still-live record carrying this dedup_key, or
        None. Used to make handle() idempotent so a periodic re-scan can't log
        the same follow-up / group suggestion over and over."""
        if not dedup_key:
            return None
        for r in self.activity:   # activity is newest-first
            if r.get("dedup_key") == dedup_key and r.get("state") in self._LIVE_STATES:
                return r
        return None

    def handle(self, *, kind, stakes, confidence, subject="", sender="",
               summary="", before=None, after=None, reason="", payload=None,
               dedup_key=None, hold_for_user=False, force_auto=False):
        """The single entry point. Decides, acts (or asks), logs, returns the
        record. `kind` is a short label like 'Category', 'Group', 'Follow-up',
        'Reply draft'. `summary` is the human one-liner ('General -> Quote /
        Estimate'). before/after let the user undo. payload carries whatever the
        host needs to actually execute (e.g. the email + target group id).

        `dedup_key` (optional) makes the call idempotent: if a still-live record
        with the same key already exists, no new row is created -- the existing
        one is refreshed and returned. Hosts pass e.g. 'followup:<entry_id>' so a
        re-triage of the same email every few minutes can't flood the log."""

        # Idempotency gate: collapse repeated suggestions for the same thing onto
        # the existing live record instead of stacking duplicates.
        if dedup_key:
            existing = self._find_live_by_dedup(dedup_key)
            if existing is not None:
                existing["repeat_count"] = int(existing.get("repeat_count", 1) or 1) + 1
                existing["last_seen"] = datetime.now().isoformat()
                # Refresh the rationale/confidence in case the latest read differs,
                # but DON'T move the row or change its original surfaced time/state.
                if reason:
                    existing["reason"] = reason
                try:
                    existing["confidence"] = round(float(confidence), 3)
                except Exception:
                    pass
                self._save()
                self._notify()
                return existing

        if not self.config.enabled:
            # Assistant is off entirely -> behave like the legacy app: ask.
            decision = ASK
        else:
            decision = self.classify(confidence, stakes)
            # Hard auto-line: anything below the auto bar is held for the user
            # (one-click Approve) instead of the old "do it but flag it" band, so
            # the visible slider is the single line that governs taking action.
            if getattr(self.config, "hard_auto_line", True) and decision == REVIEW:
                decision = ASK
            # Caller can force a hold (e.g. a destructive cleanup must always be
            # confirmed, regardless of stakes/confidence or the auto-send setting).
            if hold_for_user:
                decision = ASK
            # Caller can force auto-execution (e.g. Auto-Archive is ON, so a
            # detected cancellation/no-quote archives on its own). Independent of
            # the auto-send setting -- archiving sends nothing and is reversible.
            if force_auto:
                decision = AUTO

        rec = {
            # uuid suffix so rapid-fire actions can't collide on coarse-resolution
            # clocks (Windows datetime.now() ticks ~15ms, so microseconds repeat).
            "id": f"act_{datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex[:8]}",
            "at": datetime.now().isoformat(),
            "kind": kind,
            "stakes": stakes,
            "confidence": round(float(confidence), 3),
            "subject": subject,
            "sender": sender,
            "summary": summary,
            "before": before or {},
            "after": after or {},
            "reason": reason,
            "decision": decision,
            "state": "",            # auto | flagged | pending | done | undone | rejected
            "payload": payload or {},
            "dedup_key": dedup_key or "",
            "repeat_count": 1,
        }

        executed = False
        if decision == AUTO:
            if (stakes == HIGH and kind != "Cleanup"
                    and getattr(self.config, "autosend_delay_enabled", True)
                    and int(getattr(self.config, "autosend_delay_seconds", 0) or 0) > 0):
                # Arm a delayed auto-send: don't fire now -- hold it with a visible
                # countdown and a Cancel, then send when the timer elapses (the log
                # window ticks it). Gives the user a window to edit or stop it before
                # it actually goes out. Approve = send now; Dismiss = cancel.
                rec["state"] = "armed"
                rec["fire_at"] = (datetime.now()
                                  + timedelta(seconds=int(self.config.autosend_delay_seconds))).isoformat()
            else:
                executed = self._execute(rec)
                rec["state"] = "auto" if executed else "pending"
        elif decision == REVIEW:
            # Only reached when hard_auto_line is False: act, but flag amber.
            executed = self._execute(rec)
            rec["state"] = "flagged" if executed else "pending"
        else:  # ASK
            rec["state"] = "pending"
            self._ask(rec)

        self.activity.insert(0, rec)
        self._save()
        self._notify()
        if self.on_activity:
            try:
                self.on_activity(rec)
            except Exception:
                pass
        return rec

    def _execute(self, rec):
        if not self.do_apply:
            return False
        try:
            return bool(self.do_apply(rec))
        except Exception:
            return False

    def _ask(self, rec):
        # Uncertain items always reach the user, even when the log is hidden.
        if self.do_ask:
            try:
                self.do_ask(rec)
            except Exception:
                pass

    # ---- user actions from the log ----------------------------------------------
    def approve(self, rec_id):
        rec = self._find(rec_id)
        # "armed" = a delayed auto-send still counting down; approving = send now.
        if not rec or rec.get("state") not in ("pending", "armed"):
            return False
        if self._execute(rec):
            rec["state"] = "auto"
            rec["approved_at"] = datetime.now().isoformat()
            rec.pop("fire_at", None)
            self._save(); self._notify()
            return True
        return False

    def fire_due_auto_sends(self):
        """Execute any armed auto-sends whose countdown has elapsed. Called on a
        ~1s timer by the log window. Returns the number fired."""
        now = datetime.now()
        fired = 0
        for rec in self.activity:
            if rec.get("state") != "armed":
                continue
            try:
                due = datetime.fromisoformat(rec.get("fire_at", "")) <= now
            except Exception:
                due = True
            if due:
                ok = self._execute(rec)
                rec["state"] = "auto" if ok else "pending"
                rec.pop("fire_at", None)
                fired += 1
        if fired:
            self._save(); self._notify()
        return fired

    def reject(self, rec_id):
        rec = self._find(rec_id)
        if not rec:
            return False
        rec["state"] = "rejected"
        rec["rejected_at"] = datetime.now().isoformat()
        self._save(); self._notify()
        return True

    def undo(self, rec_id):
        rec = self._find(rec_id)
        if not rec or rec.get("state") not in ("auto", "flagged"):
            return False
        ok = True
        if self.do_undo:
            try:
                ok = bool(self.do_undo(rec))
            except Exception:
                ok = False
        if ok:
            rec["state"] = "undone"
            rec["undone_at"] = datetime.now().isoformat()
            self._save(); self._notify()
        return ok

    def correct(self, rec_id, new_after):
        """Record a user correction (and let the host re-apply + learn)."""
        rec = self._find(rec_id)
        if not rec:
            return False
        rec["corrected_from"] = rec.get("after", {})
        rec["after"] = new_after
        rec["state"] = "corrected"
        rec["corrected_at"] = datetime.now().isoformat()
        # Re-run the apply with the corrected target so the change takes effect.
        self._execute(rec)
        self._save(); self._notify()
        return True

    def mark_undone(self, rec_id):
        """Flip a record to 'undone' without re-running do_undo. Used when the host
        app already reversed the underlying action (e.g. cancelled a follow-up)."""
        rec = self._find(rec_id)
        if not rec:
            return False
        rec["state"] = "undone"
        rec["undone_at"] = datetime.now().isoformat()
        self._save(); self._notify()
        return True

    def mark_corrected(self, rec_id, new_after=None):
        """Flip a record to 'corrected' WITHOUT re-running do_apply. Used when the
        host already applied the corrected action itself (e.g. moved the email to
        a different group and taught the model). `new_after`, if given, updates the
        displayed 'after' snapshot so the log row reflects the user's choice."""
        rec = self._find(rec_id)
        if not rec:
            return False
        if new_after is not None:
            rec["corrected_from"] = rec.get("after", {})
            rec["after"] = new_after
        rec["state"] = "corrected"
        rec["corrected_at"] = datetime.now().isoformat()
        self._save(); self._notify()
        return True

    # ---- stats / helpers --------------------------------------------------------
    def counts(self):
        auto = sum(1 for r in self.activity if r.get("state") in ("auto", "flagged"))
        pending = sum(1 for r in self.activity if r.get("state") == "pending")
        return {"auto": auto, "pending": pending, "total": len(self.activity)}

    def pending(self):
        return [r for r in self.activity if r.get("state") == "pending"]

    def clear_finished(self):
        self.activity = [r for r in self.activity if r.get("state") == "pending"]
        self._save(); self._notify()

    def _find(self, rec_id):
        for r in self.activity:
            if r.get("id") == rec_id:
                return r
        return None

    def _notify(self):
        if self.on_log_change:
            try:
                self.on_log_change()
            except Exception:
                pass


# -- Live activity log window (Tkinter) ---------------------------------------
# Imported lazily so the engine can be used headless / unit-tested without Tk.
def open_assistant_log_window(parent, engine, theme=None):
    """Open (or focus) the live change-log window. `engine` is an AssistantEngine.
    `theme` is an optional dict of colors so it can match MaINbox's palette."""
    import tkinter as tk
    from tkinter import ttk, simpledialog

    t = {
        "bg":     "#15181e",
        "bg2":    "#1c2027",
        "bg3":    "#222831",
        "fg":     "#e8eaed",
        "fg_dim": "#9aa3ad",
        "accent": "#1f9cff",   # electric blue
        "auto":   "#1f9cff",   # auto actions
        "flag":   "#f2b134",   # borderline / flagged
        "ask":    "#ff7043",   # needs you
        "undone": "#6b7280",   # reversed
        "done":   "#3ddc84",   # corrected / approved
    }
    if theme:
        t.update(theme)

    # Reuse an existing window if it's still alive.
    existing = getattr(engine, "_log_win", None)
    if existing is not None:
        try:
            if existing.winfo_exists():
                existing.deiconify(); existing.lift(); existing.focus_force()
                return existing
        except Exception:
            pass

    win = tk.Toplevel(parent)
    win.title("MaINbox AI Assistant \u2014 Live Activity")
    win.configure(bg=t["bg"])
    win.geometry("1180x620")
    try:
        win.minsize(940, 480)
    except Exception:
        pass
    engine._log_win = win

    # --- top control bar (two rows so nothing is cramped) -----------------------
    bar = tk.Frame(win, bg=t["bg2"])
    bar.pack(fill="x", side="top")
    row_top = tk.Frame(bar, bg=t["bg2"]); row_top.pack(fill="x")
    row_bot = tk.Frame(bar, bg=t["bg2"]); row_bot.pack(fill="x")
    row_bot2 = tk.Frame(bar, bg=t["bg2"]); row_bot2.pack(fill="x")

    enabled_var     = tk.BooleanVar(value=engine.config.enabled)
    showlog_var     = tk.BooleanVar(value=engine.config.show_log)
    file_var        = tk.DoubleVar(value=engine.config.low_auto_at)    # low-stakes: file/label/group
    send_var        = tk.DoubleVar(value=engine.config.high_auto_at)   # high-stakes: send/deploy a draft
    autosend_var    = tk.BooleanVar(value=engine.config.allow_auto_high_stakes)
    autoarchive_var = tk.BooleanVar(value=getattr(engine.config, "autoarchive_enabled", False))

    def _persist():
        if hasattr(engine, "save_config"):
            engine.save_config()

    def on_enabled():
        engine.config.enabled = enabled_var.get()
        _persist(); _refresh_status()

    def on_showlog():
        # This is an auto-OPEN preference, not a close button. Checked = pop this
        # window for ANY new activity; unchecked = pop only when something needs
        # attention. Toggling it just saves the preference -- it never opens or
        # closes the window (use the X to hide it).
        engine.config.show_log = showlog_var.get()
        _persist()

    def on_file(_=None):
        # Governs ONLY low-stakes filing: category, urgency, status, grouping.
        engine.config.low_auto_at = round(file_var.get(), 2)
        file_lbl.config(text=f"Auto-file labels/status/groups \u2265 {int(engine.config.low_auto_at*100)}%")
        _persist()

    def on_send(_=None):
        # Governs ONLY high-stakes actions: sending/deploying a draft. Independent
        # of the file bar so raising one never silently moves the other.
        engine.config.high_auto_at = round(send_var.get(), 2)
        send_lbl.config(text=f"Auto-send drafts \u2265 {int(engine.config.high_auto_at*100)}%")
        _persist()

    def on_autosend():
        # CHECKED = auto-send ON: high-stakes sends fire on their own, but always
        # with the built-in ~2 min delay (an armed countdown you can cancel/edit).
        # UNCHECKED = stage only: sends wait for your Approve.
        engine.config.allow_auto_high_stakes = autosend_var.get()
        engine.config.autosend_delay_enabled = True  # the delay is part of auto-send
        send_scale.config(state="normal" if autosend_var.get() else "disabled")
        on_send()

    def on_autoarchive():
        # CHECKED = a detected cancellation / no-quote archives the group on its own.
        # UNCHECKED = it's staged as "Needs you" for your one-click Approve.
        engine.config.autoarchive_enabled = autoarchive_var.get()
        _persist()

    # Row 1: master switches + live counter
    tk.Checkbutton(row_top, text="Assistant ON", variable=enabled_var, command=on_enabled,
                   bg=t["bg2"], fg=t["fg"], selectcolor=t["bg3"],
                   activebackground=t["bg2"], activeforeground=t["accent"],
                   font=("Segoe UI", 10, "bold")).pack(side="left", padx=(12, 18), pady=(8, 2))
    tk.Checkbutton(row_top, text="Show this log", variable=showlog_var, command=on_showlog,
                   bg=t["bg2"], fg=t["fg_dim"], selectcolor=t["bg3"],
                   activebackground=t["bg2"], activeforeground=t["accent"]).pack(side="left", padx=(0, 18), pady=(8, 2))
    status_lbl = tk.Label(row_top, text="", bg=t["bg2"], fg=t["accent"], font=("Segoe UI", 10, "bold"))
    status_lbl.pack(side="right", padx=14)

    # Row 2: the two independent confidence bars
    file_lbl = tk.Label(row_bot, text=f"Auto-file labels/status/groups \u2265 {int(engine.config.low_auto_at*100)}%",
                        bg=t["bg2"], fg=t["fg_dim"])
    file_lbl.pack(side="left", padx=(12, 4), pady=(0, 8))
    tk.Scale(row_bot, from_=0.40, to=0.95, resolution=0.01, orient="horizontal",
             variable=file_var, command=on_file, showvalue=False, length=140,
             bg=t["bg2"], fg=t["fg"], troughcolor=t["bg3"], highlightthickness=0,
             activebackground=t["accent"]).pack(side="left", pady=(0, 8))

    send_lbl = tk.Label(row_bot, text=f"Auto-send drafts \u2265 {int(engine.config.high_auto_at*100)}%",
                        bg=t["bg2"], fg=t["fg_dim"])
    send_lbl.pack(side="left", padx=(22, 4), pady=(0, 8))
    send_scale = tk.Scale(row_bot, from_=0.40, to=0.99, resolution=0.01, orient="horizontal",
                          variable=send_var, command=on_send, showvalue=False, length=140,
                          bg=t["bg2"], fg=t["fg"], troughcolor=t["bg3"], highlightthickness=0,
                          activebackground=t["accent"])
    send_scale.pack(side="left", pady=(0, 8))
    _dly = int(getattr(engine.config, "autosend_delay_seconds", 120) or 120)
    _dly_txt = (f"{_dly // 60} min" if _dly % 60 == 0 else f"{_dly}s")
    # Both stacked checkboxes share one width + left anchor so their boxes line up
    # vertically (right-aligned to the same edge), matching the layout request.
    _cbw = 36
    autosend_chk = tk.Checkbutton(row_bot, text=f"Auto-send ({_dly_txt} delay)",
                   variable=autosend_var, command=on_autosend, width=_cbw, anchor="w",
                   bg=t["bg2"], fg=t["fg_dim"], selectcolor=t["bg3"],
                   activebackground=t["bg2"], activeforeground=t["accent"])
    autosend_chk.pack(side="right", padx=(0, 14), pady=(0, 8))
    if not autosend_var.get():
        send_scale.config(state="disabled")

    # Row 3 (right-aligned, sits directly under the auto-send box): auto-archive a
    # confidently-detected CLOSED quote -- a won order, a cancellation, a no-quote,
    # or a clear customer decline -- on its own instead of asking first. Soft/AI-read
    # closings are always held for review and never ride this setting.
    tk.Checkbutton(row_bot2, text="Auto-archive closed quotes", width=_cbw, anchor="w",
                   variable=autoarchive_var, command=on_autoarchive,
                   bg=t["bg2"], fg=t["fg_dim"], selectcolor=t["bg3"],
                   activebackground=t["bg2"], activeforeground=t["accent"]).pack(side="right", padx=(0, 14), pady=(0, 8))

    # --- legend -----------------------------------------------------------------
    legend = tk.Frame(win, bg=t["bg"])
    legend.pack(fill="x")
    for label, color in (("\u25cf Auto", t["auto"]), ("\u25cf Flagged", t["flag"]),
                         ("\u25cf Needs you", t["ask"]), ("\u25cf Done", t["done"]),
                         ("\u25cf Undone", t["undone"])):
        tk.Label(legend, text=label, fg=color, bg=t["bg"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(12, 0), pady=(4, 2))

    # --- activity table ---------------------------------------------------------
    style = ttk.Style(win)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Assist.Treeview", background=t["bg3"], fieldbackground=t["bg3"],
                    foreground=t["fg"], rowheight=26, borderwidth=0)
    style.configure("Assist.Treeview.Heading", background=t["bg2"], foreground=t["fg_dim"])
    # Dark-theme the correction dropdowns (combobox field + its popup listbox).
    try:
        style.configure("Assist.TCombobox", fieldbackground=t["bg3"], background=t["bg3"],
                        foreground=t["fg"], arrowcolor=t["fg"], borderwidth=0)
        style.map("Assist.TCombobox",
                  fieldbackground=[("readonly", t["bg3"])],
                  foreground=[("readonly", t["fg"])],
                  selectbackground=[("readonly", t["bg3"])],
                  selectforeground=[("readonly", t["fg"])])
    except Exception:
        pass
    try:
        win.option_add("*TCombobox*Listbox.background", t["bg3"])
        win.option_add("*TCombobox*Listbox.foreground", t["fg"])
        win.option_add("*TCombobox*Listbox.selectBackground", t["accent"])
        win.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
        win.option_add("*TCombobox*Listbox.font", "{Segoe UI} 10")
    except Exception:
        pass

    # `checked` holds the iids whose checkbox is ticked. The leading [ ]/[x] column
    # makes multi-select obvious and tap-friendly; selectmode="extended" also
    # allows ctrl/shift range-selection. Actions (Approve/Dismiss/Undo/Correct...)
    # operate on every checked row, or - when nothing is checked - the one
    # highlighted row, so single-item use is unchanged.
    checked = set()
    cols = ("chk", "time", "kind", "what", "target", "conf", "state")
    tree = ttk.Treeview(win, columns=cols, show="headings", style="Assist.Treeview", selectmode="extended")
    tree.heading("chk", text="\u2610", command=lambda: _toggle_all())
    tree.column("chk", width=34, anchor="center", stretch=False)
    for col, label, w in (("time", "Time", 80), ("kind", "Action", 110),
                          ("what", "What changed", 300), ("target", "Email", 250),
                          ("conf", "Sure", 60), ("state", "Status", 90)):
        tree.heading(col, text=label)
        tree.column(col, width=w, anchor="w")
    tree.pack(fill="both", expand=True, padx=10, pady=(4, 6))

    tree.tag_configure("auto",     foreground=t["auto"])
    tree.tag_configure("flagged",  foreground=t["flag"])
    tree.tag_configure("pending",  foreground=t["ask"], background="#2a2230")
    tree.tag_configure("undone",   foreground=t["undone"])
    tree.tag_configure("rejected", foreground=t["undone"])
    tree.tag_configure("done",     foreground=t["done"])
    tree.tag_configure("armed",    foreground=t["ask"], background="#33271a")

    state_label = {"auto": "Auto", "flagged": "Flagged", "pending": "Needs you",
                   "undone": "Undone", "rejected": "Dismissed", "corrected": "Corrected",
                   "armed": "Sending\u2026"}

    def _row_tag(state):
        if state == "corrected":
            return "done"
        return state if state in ("auto", "flagged", "pending", "undone", "rejected", "armed") else "auto"

    def _refresh_status():
        c = engine.counts()
        on = "ON" if engine.config.enabled else "OFF"
        status_lbl.config(text=f"Assistant {on}   \xb7   {c['auto']} handled   \xb7   {c['pending']} need you")

    def refresh():
        try:
            if not tree.winfo_exists():
                return
        except Exception:
            return
        sel = tree.selection()
        keep = sel[0] if sel else None
        tree.delete(*tree.get_children())
        seen_iids = set()
        for r in engine.activity[:600]:
            rid = r.get("id")
            # Belt-and-suspenders: never insert the same iid twice in one render.
            if not rid or rid in seen_iids or tree.exists(rid):
                continue
            seen_iids.add(rid)
            try:
                ts = datetime.fromisoformat(r.get("at", "")).strftime("%I:%M %p")
            except Exception:
                ts = ""
            conf = f"{int(round(r.get('confidence', 0)*100))}%"
            glyph = "\u2611" if rid in checked else "\u2610"
            st_txt = state_label.get(r.get("state", ""), r.get("state", ""))
            # Repeated identical suggestions are now collapsed onto one row; show
            # how many times it recurred so the user still sees it's been noisy.
            rc = int(r.get("repeat_count", 1) or 1)
            if rc > 1:
                st_txt = f"{st_txt} \xd7{rc}"
            tree.insert("", "end", iid=rid,
                        values=(glyph, ts, r.get("kind", ""), r.get("summary", ""),
                                f"{r.get('sender','')} \u2014 {r.get('subject','')}"[:70],
                                conf, st_txt),
                        tags=(_row_tag(r.get("state", "")),))
        # Drop checks for rows that no longer exist, then sync the header box.
        for iid in list(checked):
            if not tree.exists(iid):
                checked.discard(iid)
        try:
            _update_header_check()
        except Exception:
            pass
        if keep and tree.exists(keep):
            tree.selection_set(keep)
        _refresh_status()
        _sync_buttons()

    # --- "why this decision" line -----------------------------------------------
    why_lbl = tk.Label(win, text="Select a row to see why the assistant did or didn't act.",
                       bg=t["bg"], fg=t["fg_dim"], anchor="w", justify="left", wraplength=960)
    why_lbl.pack(fill="x", padx=12, pady=(0, 4))

    def _rationale(rec):
        if not rec:
            return "Select a row to see why the assistant did or didn't act."
        c = engine.config
        conf = int(round(rec.get("confidence", 0) * 100))
        reason = (rec.get("reason", "") or "").strip()
        tail = f"  ({reason})" if reason else ""
        if rec.get("state") == "armed":
            try:
                rem = int(max(0, (datetime.fromisoformat(rec.get("fire_at", "")) - datetime.now()).total_seconds()))
                mmss = f"{rem // 60}:{rem % 60:02d}"
            except Exception:
                mmss = "soon"
            return (f"Auto-sending in {mmss} \u2014 Dismiss to cancel, or Approve to send it "
                    f"now. Nothing leaves until the timer reaches 0.{tail}")
        if rec.get("kind") == "Cleanup":
            if rec.get("state") in ("auto", "flagged"):
                return ("Auto-archived: a cancellation / no-quote was detected, so the "
                        "group was archived on its own (emails kept and hidden; "
                        f"unarchive to bring them back).{tail}")
            return ("Cancellation / no-quote detected \u2014 staged archiving the group "
                    "(emails kept and hidden; unarchiving brings them all back) for "
                    f"your one-click Approve. Dismiss to leave it as is.{tail}")
        if rec.get("stakes") == HIGH:
            bar = int(c.high_auto_at * 100)
            if not c.allow_auto_high_stakes:
                return f"High-stakes (sends/deploys). Auto-send is OFF, so it's staged for your one-click Approve \u2014 {conf}% sure.{tail}"
            if rec.get("decision") == "auto":
                return f"High-stakes, and {conf}% met your {bar}% auto-send bar, so it was sent automatically.{tail}"
            return f"High-stakes (sends a draft). {conf}% is below your {bar}% auto-send bar, so it's waiting for you. Lower the send bar to automate it.{tail}"
        bar = int(c.low_auto_at * 100)
        if rec.get("decision") == "auto":
            return f"Low-stakes filing, and {conf}% met your {bar}% auto-file bar, so it was done automatically.{tail}"
        if rec.get("decision") == "review":
            return f"Low-stakes but borderline at {conf}% \u2014 applied and flagged amber for a quick glance.{tail}"
        return f"Low-stakes, but {conf}% is below your {bar}% auto-file bar, so it's waiting for you.{tail}"

    # --- action buttons ---------------------------------------------------------
    btns = tk.Frame(win, bg=t["bg"])
    btns.pack(fill="x", padx=10, pady=(0, 10))

    def _selected():
        sel = tree.selection()
        if not sel:
            return None
        return engine._find(sel[0])

    # ---- checkbox / multi-select plumbing --------------------------------------
    def _visible_iids():
        return list(tree.get_children())

    def _update_header_check():
        vis = set(_visible_iids())
        all_checked = bool(vis) and vis.issubset(checked)
        tree.heading("chk", text=("\u2611" if all_checked else "\u2610"))

    def _toggle_all():
        vis = _visible_iids()
        if checked and set(vis).issubset(checked):
            for i in vis:
                checked.discard(i)
        else:
            checked.update(vis)
        for i in vis:
            try:
                tree.set(i, "chk", "\u2611" if i in checked else "\u2610")
            except Exception:
                pass
        _update_header_check()
        _sync_buttons()

    def _clear_checks():
        checked.clear()
        for i in _visible_iids():
            try:
                tree.set(i, "chk", "\u2610")
            except Exception:
                pass
        _update_header_check()

    def _targets():
        """Records the buttons act on: every checked row, or -- if none are
        checked -- every highlighted row. So highlighting a range (or ticking
        the boxes) both drive bulk actions; a single click still acts on one."""
        ids = [i for i in checked if tree.exists(i)]
        if not ids:
            ids = list(tree.selection())
        out = []
        for i in ids:
            r = engine._find(i)
            if r:
                out.append(r)
        return out

    def do_approve():
        for r in _targets():
            if r.get("state") == "pending":
                engine.approve(r["id"])
        _clear_checks()

    def do_reject():
        for r in _targets():
            if r.get("state") == "pending":
                engine.reject(r["id"])
        _clear_checks()

    def do_undo():
        for r in _targets():
            if r.get("state") in ("auto", "flagged"):
                engine.undo(r["id"])
        _clear_checks()

    # ---- correction via dropdowns (single + guided bulk) -----------------------
    def _dropdown_popup(title, prompt, choices, current=None, apply_label="Apply"):
        """Modal combobox picker. `choices` is a list of (value, display) tuples.
        Returns (value, display), or None if cancelled."""
        dlg = tk.Toplevel(win)
        dlg.title(title)
        dlg.configure(bg=t["bg"])
        dlg.transient(win)
        try:
            dlg.grab_set()
        except Exception:
            pass
        tk.Label(dlg, text=prompt, bg=t["bg"], fg=t["fg"], font=("Segoe UI", 10),
                 wraplength=440, justify="left").pack(anchor="w", padx=16, pady=(16, 10))
        displays = [d for (_v, d) in choices]
        var = tk.StringVar()
        cur_display = ""
        for v, d in choices:
            if current is not None and v == current:
                cur_display = d
                break
        var.set(cur_display or (displays[0] if displays else ""))
        combo = ttk.Combobox(dlg, textvariable=var, values=displays, state="readonly",
                             width=48, style="Assist.TCombobox")
        combo.pack(padx=16, pady=(0, 16))
        result = {"picked": None}

        def _apply():
            d = var.get()
            for v, dd in choices:
                if dd == d:
                    result["picked"] = (v, dd)
                    break
            dlg.destroy()

        def _cancel():
            result["picked"] = None
            dlg.destroy()

        brow = tk.Frame(dlg, bg=t["bg"]); brow.pack(fill="x", padx=16, pady=(0, 16))
        tk.Button(brow, text=apply_label, command=_apply, bg=t["accent"], fg="#ffffff",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=16, pady=6,
                  cursor="hand2", activebackground=t["done"], activeforeground="#ffffff").pack(side="left")
        tk.Button(brow, text="Cancel", command=_cancel, bg=t["bg3"], fg=t["fg"],
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=16, pady=6,
                  cursor="hand2", activebackground=t["ask"], activeforeground="#ffffff").pack(side="right")
        dlg.bind("<Return>", lambda e: _apply())
        dlg.bind("<Escape>", lambda e: _cancel())
        try:
            dlg.update_idletasks()
            px, py = win.winfo_rootx(), win.winfo_rooty()
            pw, ph = win.winfo_width(), win.winfo_height()
            w_, h_ = dlg.winfo_width(), dlg.winfo_height()
            dlg.geometry(f"+{px + max(0,(pw - w_)//2)}+{py + max(0,(ph - h_)//2)}")
        except Exception:
            pass
        dlg.wait_window()
        return result["picked"]

    def _correction_options(rec):
        cb = getattr(engine, "correction_options_cb", None)
        if callable(cb):
            try:
                return cb(rec)
            except Exception:
                return None
        return None

    def _correct_one(rec, field, value, display):
        cb = getattr(engine, "apply_correction_cb", None)
        if callable(cb):
            try:
                return bool(cb(rec, field, value, display))
            except Exception:
                return False
        # Fallback: record the corrected field on the record itself.
        after = dict(rec.get("after", {}))
        after[field] = value
        return engine.correct(rec["id"], after)

    def _single_correct(rec):
        opts = _correction_options(rec)
        if not opts or not opts.get("choices"):
            # Legacy free-text fallback (host didn't supply dropdown options).
            after = rec.get("after", {})
            field = "status" if "status" in after else (list(after.keys())[0] if after else "value")
            current = after.get(field, "")
            newval = simpledialog.askstring("Correct", f"Set {field} to:", initialvalue=str(current), parent=win)
            if newval is not None:
                corrected = dict(after); corrected[field] = newval
                engine.correct(rec["id"], corrected)
            return
        picked = _dropdown_popup(opts.get("title", "Correct"),
                                 opts.get("prompt", "Choose the correct value:"),
                                 opts["choices"], opts.get("current"), apply_label="Apply")
        if picked is not None:
            _correct_one(rec, opts["field"], picked[0], picked[1])

    def _bulk_correct(recs):
        """Guided multi-step correction for a mixed selection. Follow-ups get one
        popup (pick an action, Apply to all / Cancel); group rows get the next
        popup (pick a group, Apply to all / Cancel); any other kinds get one
        popup per field after that."""
        followups = [r for r in recs if r.get("kind") in ("Follow-up", "Reply draft")]
        groups    = [r for r in recs if r.get("kind") == "Group"]
        others    = [r for r in recs if r not in followups and r not in groups]

        # Popup 1 - the follow-ups: one action for all of them.
        if followups:
            o = _correction_options(followups[0])
            if o and o.get("choices"):
                picked = _dropdown_popup(
                    "Correct follow-ups",
                    f"Choose what to do with the {len(followups)} selected follow-up(s):",
                    o["choices"], o.get("current"), apply_label="Apply to all")
                if picked is not None:
                    for r in followups:
                        _correct_one(r, o["field"], picked[0], picked[1])

        # Popup 2 - the group rows: one target group for all of them.
        if groups:
            o = _correction_options(groups[0])
            if o and o.get("choices"):
                picked = _dropdown_popup(
                    "Re-group emails",
                    f"Choose the group for the {len(groups)} selected email(s):",
                    o["choices"], o.get("current"), apply_label="Apply to all")
                if picked is not None:
                    for r in groups:
                        _correct_one(r, o["field"], picked[0], picked[1])

        # Any other kinds (Category / Status / Urgency): group by field, one each.
        if others:
            by_field = {}
            for r in others:
                o = _correction_options(r)
                if not o or not o.get("choices"):
                    continue
                if o["field"] not in by_field:
                    by_field[o["field"]] = (o, [])
                by_field[o["field"]][1].append(r)
            for field, (o, rs) in by_field.items():
                picked = _dropdown_popup(
                    f"Correct {field}",
                    f"Choose the correct {field} for the {len(rs)} selected item(s):",
                    o["choices"], o.get("current"), apply_label="Apply to all")
                if picked is not None:
                    for r in rs:
                        _correct_one(r, field, picked[0], picked[1])

    def do_correct():
        tgts = _targets()
        if not tgts:
            return
        if len(tgts) == 1:
            _single_correct(tgts[0])
        else:
            _bulk_correct(tgts)
        _clear_checks()

    def open_email():
        tgts = _targets()
        r = _selected() or (tgts[0] if tgts else None)
        if r and callable(getattr(engine, "open_email_cb", None)):
            try:
                engine.open_email_cb(r)
            except Exception:
                pass

    mkbtn = lambda txt, cmd, fg=t["fg"]: tk.Button(
        btns, text=txt, command=cmd, bg=t["bg3"], fg=fg, relief="flat",
        activebackground=t["accent"], activeforeground="#ffffff",
        font=("Segoe UI", 9, "bold"), padx=12, pady=5, cursor="hand2")

    b_approve = mkbtn("\u2713 Approve", do_approve, t["done"]);  b_approve.pack(side="left", padx=(0, 6))
    b_reject  = mkbtn("\u2715 Dismiss", do_reject, t["ask"]);    b_reject.pack(side="left", padx=6)
    b_undo    = mkbtn("\u21b6 Undo", do_undo);                   b_undo.pack(side="left", padx=6)
    b_correct = mkbtn("\u270e Correct\u2026", do_correct);            b_correct.pack(side="left", padx=6)
    mkbtn("Open email", open_email, t["accent"]).pack(side="left", padx=6)
    mkbtn("Clear handled", lambda: engine.clear_finished()).pack(side="right")

    def _sync_buttons():
        tgts = _targets()
        any_pending  = any(x.get("state") == "pending" for x in tgts)
        any_undoable = any(x.get("state") in ("auto", "flagged") for x in tgts)
        b_approve.config(state="normal" if any_pending else "disabled")
        b_reject.config(state="normal" if any_pending else "disabled")
        b_undo.config(state="normal" if any_undoable else "disabled")
        b_correct.config(state="normal" if tgts else "disabled")
        n = len(tgts)
        if n > 1:
            why_lbl.config(text=(f"{n} items selected \u2014 Approve / Dismiss / Undo apply to all of them. "
                                 "Correct\u2026 opens a guided fix: an action for the follow-ups, then a group for the emails."))
        else:
            why_lbl.config(text=_rationale(_selected()))

    def _on_tree_click(event):
        # A click in the leading [ ]/[x] column toggles just that row's checkbox; it
        # does not move the highlight or change the "why" line. Returning "break"
        # suppresses the default row-select for checkbox clicks only.
        region = tree.identify("region", event.x, event.y)
        if region == "heading":
            return            # the heading's own command does select-all/none
        if region != "cell" or tree.identify_column(event.x) != "#1":
            return            # normal click elsewhere -> default selection runs
        row = tree.identify_row(event.y)
        if not row:
            return
        if row in checked:
            checked.discard(row)
        else:
            checked.add(row)
        try:
            tree.set(row, "chk", "\u2611" if row in checked else "\u2610")
        except Exception:
            pass
        _update_header_check()
        _sync_buttons()
        return "break"
    tree.bind("<Button-1>", _on_tree_click)

    tree.bind("<<TreeviewSelect>>", lambda e: _sync_buttons())

    def _on_row_activate(_=None):
        r = _selected()
        cb = getattr(engine, "on_row_activate", None)
        if r and callable(cb):
            try:
                cb(r)
            except Exception:
                pass
    tree.bind("<Double-1>", _on_row_activate)

    # The engine notifies us on every change; marshal onto the Tk thread and
    # coalesce bursts (e.g. "Feed all") into a single refresh. Guard every hop so
    # a notification arriving just after the window closes can't fire against a
    # destroyed Treeview ("invalid command name").
    _pending = {"id": None}
    def _on_change():
        try:
            if not win.winfo_exists():
                return
            if _pending["id"] is not None:
                try:
                    win.after_cancel(_pending["id"])
                except Exception:
                    pass
            _pending["id"] = win.after(40, refresh)
        except Exception:
            pass
    engine.on_log_change = _on_change

    def _on_close():
        # Closing the window just hides it; the assistant keeps running and will
        # re-open per the "Show this log" preference. Closing does NOT change that
        # preference (the checkbox is the only thing that does).
        engine.on_log_change = None     # stop the engine poking a dead widget
        engine._log_win = None
        try:
            win.destroy()
        except Exception:
            pass
    win.protocol("WM_DELETE_WINDOW", _on_close)

    def _tick():
        # ~1s heartbeat: live-update the countdown on any armed (delayed) auto-send
        # rows, then fire the ones whose timer has elapsed. Cheap no-op when there
        # are no armed rows (the normal case).
        try:
            if not win.winfo_exists():
                return
        except Exception:
            return
        try:
            now = datetime.now()
            for iid in tree.get_children():
                rec = engine._find(iid)
                if rec and rec.get("state") == "armed":
                    try:
                        rem = int(max(0, (datetime.fromisoformat(rec["fire_at"]) - now).total_seconds()))
                        tree.set(iid, "state", f"Send {rem // 60}:{rem % 60:02d}")
                    except Exception:
                        pass
            engine.fire_due_auto_sends()
        except Exception:
            pass
        try:
            win.after(1000, _tick)
        except Exception:
            pass
    win.after(1000, _tick)

    refresh()
    return win


# -- Built-in demo (no Outlook / Ollama required) -----------------------------
def _demo():
    """Run this file directly to feel the whole UX with fake emails."""
    import tkinter as tk
    from tkinter import messagebox

    root = tk.Tk()
    root.title("MaINbox AI Assistant \u2014 Demo host")
    root.configure(bg="#15181e")
    root.geometry("560x300")

    # A fake "email store" so undo/correct have something to act on.
    store = {}

    cfg = AssistantConfig()
    engine = AssistantEngine(config=cfg, log_path=os.path.join(os.getcwd(), "assistant_demo_activity.json"))

    def do_apply(rec):
        # Pretend to write the change into the email/store.
        store[rec["subject"]] = dict(rec.get("after", {}))
        return True

    def do_undo(rec):
        store[rec["subject"]] = dict(rec.get("before", {}))
        return True

    def do_ask(rec):
        # Uncertain item: if the log is hidden, pop it to the front so the user
        # still sees it (this is the "bypass silent mode" behaviour).
        if not engine.config.show_log:
            win = open_assistant_log_window(root, engine)
            try:
                win.deiconify(); win.lift(); win.bell()
            except Exception:
                pass

    def open_email_cb(rec):
        messagebox.showinfo("Open email (demo)",
                            f"Would open in Outlook:\n\n{rec.get('sender')}\n{rec.get('subject')}")

    engine.do_apply = do_apply
    engine.do_undo = do_undo
    engine.do_ask = do_ask
    engine.open_email_cb = open_email_cb

    # A handful of realistic decisions spanning the confidence range. These dicts
    # use the SAME fields MaINbox triage already produces.
    samples = [
        dict(kind="Category", stakes=LOW, subject="RE: Quote 8841 - 500ft MC cable",
             sender="buyer@accmech.com", summary="General \u2192 Quote / Estimate",
             before={"category": "General"}, after={"category": "Quote / Estimate"},
             reason="Customer request phrase matched: 'please quote'",
             result=dict(business_relevance=0.91, spam_likelihood=0.03)),
        dict(kind="Status", stakes=LOW, subject="Re: PO coming Monday for job 22-104",
             sender="pm@harriselectric.com", summary="Needs Reply \u2192 Waiting on Customer",
             before={"status": "Needs Reply"}, after={"status": "Waiting on Customer"},
             reason="Customer says PO/approval pending",
             result=dict(business_relevance=0.90, spam_likelihood=0.02)),
        dict(kind="Group", stakes=LOW, subject="Re: Switchgear lead time?",
             sender="sales@nationalbreaker.com", summary="Add to group 'Harris \u2014 Bldg 4 gear'",
             before={"group": ""}, after={"group": "Harris \u2014 Bldg 4 gear"},
             reason="Same conversation as 3 existing emails",
             result=dict(group_confidence=0.81), is_group=True),
        dict(kind="Category", stakes=LOW, subject="Your invoice is ready \u2014 Verizon",
             sender="no-reply@verizon.com", summary="General \u2192 Junk / Ignore",
             before={"category": "General"}, after={"category": "Junk / Ignore"},
             reason="Automated/marketing sender",
             result=dict(business_relevance=0.05, spam_likelihood=0.92)),
        dict(kind="Group", stakes=LOW, subject="quick question",
             sender="dave@gmail.com", summary="Maybe group 'Misc RFQs'?",
             before={"group": ""}, after={"group": "Misc RFQs"},
             reason="Weak token overlap only",
             result=dict(group_confidence=0.58), is_group=True),
        dict(kind="Follow-up", stakes=HIGH, subject="Re: Still need pricing on the gear",
             sender="buyer@accmech.com", summary="Send vendor follow-up draft (60 min)",
             before={}, after={"action": "send_followup_draft"},
             reason="Vendor went quiet; customer waiting",
             result=dict(business_relevance=0.90, spam_likelihood=0.03)),
        dict(kind="Reply draft", stakes=HIGH, subject="RE: can you confirm ship date",
             sender="pm@harriselectric.com", summary="Auto-send 'confirmed, ships Friday' reply",
             before={}, after={"action": "send_reply"},
             reason="Borderline: needs a human eye on wording",
             result=dict(business_relevance=0.72, spam_likelihood=0.05)),
    ]

    idx = {"i": 0}

    def feed_one():
        if idx["i"] >= len(samples):
            idx["i"] = 0
        s = samples[idx["i"]]; idx["i"] += 1
        conf = confidence_from_group(s["result"]) if s.get("is_group") else confidence_from_triage(s["result"])
        engine.handle(kind=s["kind"], stakes=s["stakes"], confidence=conf,
                      subject=s["subject"], sender=s["sender"], summary=s["summary"],
                      before=s["before"], after=s["after"], reason=s["reason"])

    def feed_all():
        for _ in range(len(samples)):
            feed_one()

    tk.Label(root, text="MaINbox AI Assistant \u2014 demo", bg="#15181e", fg="#e8eaed",
             font=("Segoe UI", 14, "bold")).pack(pady=(18, 4))
    tk.Label(root, text="Feed fake emails through the assistant and watch the live log.\n"
                        "Confident actions apply automatically; unsure ones wait for you.",
             bg="#15181e", fg="#9aa3ad", justify="center").pack(pady=(0, 14))

    mk = lambda txt, cmd, fg="#e8eaed": tk.Button(
        root, text=txt, command=cmd, bg="#222831", fg=fg, relief="flat",
        activebackground="#1f9cff", activeforeground="#ffffff",
        font=("Segoe UI", 10, "bold"), padx=14, pady=8, cursor="hand2")

    row = tk.Frame(root, bg="#15181e"); row.pack(pady=4)
    mk("\u25b6 Feed one email", feed_one, "#1f9cff").pack(side="left", padx=6)
    mk("\u23e9 Feed all", feed_all).pack(side="left", padx=6)
    mk("Open live log", lambda: open_assistant_log_window(root, engine), "#3ddc84").pack(side="left", padx=6)

    open_assistant_log_window(root, engine)
    root.mainloop()


# -- Headless self-test (runs anywhere, no display) ---------------------------
def _selftest():
    cfg = AssistantConfig()
    applied = []
    eng = AssistantEngine(config=cfg, do_apply=lambda r: applied.append(r["id"]) or True)

    # high confidence, low stakes -> AUTO
    r1 = eng.handle(kind="Category", stakes=LOW,
                    confidence=confidence_from_triage(dict(business_relevance=0.91, spam_likelihood=0.03)),
                    subject="quote", summary="x")
    assert r1["decision"] == AUTO and r1["state"] == "auto", r1

    # spammy -> very low conf -> ASK
    r2 = eng.handle(kind="Category", stakes=LOW,
                    confidence=confidence_from_triage(dict(business_relevance=0.05, spam_likelihood=0.92)),
                    subject="spam", summary="x")
    assert r2["decision"] == ASK and r2["state"] == "pending", r2

    # borderline low stakes -> held for approval (hard_auto_line default ON)
    r3 = eng.handle(kind="Group", stakes=LOW,
                    confidence=confidence_from_group(dict(group_confidence=0.60)),
                    subject="grp", summary="x")
    assert r3["decision"] == ASK and r3["state"] == "pending", r3

    # ...but with hard_auto_line OFF, the same borderline case acts + flags
    eng_soft = AssistantEngine(config=AssistantConfig(hard_auto_line=False),
                               do_apply=lambda r: True)
    r3b = eng_soft.handle(kind="Group", stakes=LOW,
                          confidence=confidence_from_group(dict(group_confidence=0.60)),
                          subject="grp", summary="x")
    assert r3b["decision"] == REVIEW and r3b["state"] == "flagged", r3b

    # high stakes borderline -> ASK (staged), not auto-sent
    r4 = eng.handle(kind="Reply draft", stakes=HIGH,
                    confidence=confidence_from_triage(dict(business_relevance=0.72, spam_likelihood=0.05)),
                    subject="reply", summary="x")
    assert r4["decision"] == ASK and r4["state"] == "pending", r4

    # high stakes high conf -> AUTO; but the send delay is ON by default, so the
    # send is ARMED (held with a countdown), not fired immediately.
    r5 = eng.handle(kind="Reply draft", stakes=HIGH,
                    confidence=confidence_from_triage(dict(business_relevance=0.97, spam_likelihood=0.0)),
                    subject="fu", summary="x")
    assert r5["decision"] == AUTO and r5["state"] == "armed" and r5.get("fire_at"), r5
    assert r5["id"] not in applied, "armed send must NOT have executed yet"
    # let the countdown elapse, then fire it
    r5["fire_at"] = (datetime.now() - timedelta(seconds=1)).isoformat()
    assert eng.fire_due_auto_sends() == 1
    assert eng._find(r5["id"])["state"] == "auto" and r5["id"] in applied

    # Approve on an armed row = send now
    r5b = eng.handle(kind="Reply draft", stakes=HIGH,
                     confidence=confidence_from_triage(dict(business_relevance=0.98, spam_likelihood=0.0)),
                     subject="fu2", summary="x", dedup_key="t1")
    assert r5b["state"] == "armed"
    assert eng.approve(r5b["id"]) is True and eng._find(r5b["id"])["state"] == "auto"

    # Dismiss on an armed row = cancel; a cancelled send must never fire
    r5c = eng.handle(kind="Reply draft", stakes=HIGH,
                     confidence=confidence_from_triage(dict(business_relevance=0.98, spam_likelihood=0.0)),
                     subject="fu3", summary="x", dedup_key="t2")
    assert r5c["state"] == "armed"
    assert eng.reject(r5c["id"]) is True and eng._find(r5c["id"])["state"] == "rejected"
    assert eng.fire_due_auto_sends() == 0, "a cancelled armed send must never fire"

    # v4.2.82 regression (live bug: the Dennis McCloskey follow-up): a DISMISSED
    # record must swallow re-scans of the same dedup_key. "rejected" was missing
    # from _LIVE_STATES, so every re-triage re-created the ask as a fresh pending
    # row and Dismiss never stuck.
    n_before = len(eng.activity)
    r5c2 = eng.handle(kind="Reply draft", stakes=HIGH,
                      confidence=confidence_from_triage(dict(business_relevance=0.98, spam_likelihood=0.0)),
                      subject="fu3", summary="x", dedup_key="t2")
    assert r5c2["id"] == r5c["id"], "re-scan must collapse onto the dismissed record"
    assert r5c2["state"] == "rejected", "collapse must not resurrect a dismissed record"
    assert len(eng.activity) == n_before, "no new row for a dismissed dedup_key"
    assert eng.fire_due_auto_sends() == 0
    # ...and the pending-ask variant, the exact reported shape: ask -> Dismiss ->
    # re-scan must NOT bring it back.
    p1 = eng.handle(kind="Follow-up", stakes=LOW, confidence=0.5,
                    subject="fu5", summary="ask me", dedup_key="t3", hold_for_user=True)
    assert p1["state"] == "pending", p1
    assert eng.reject(p1["id"]) is True
    n_before = len(eng.activity)
    p2 = eng.handle(kind="Follow-up", stakes=LOW, confidence=0.5,
                    subject="fu5", summary="ask me", dedup_key="t3", hold_for_user=True)
    assert p2["id"] == p1["id"] and p2["state"] == "rejected", p2
    assert len(eng.activity) == n_before, "dismissed follow-up came back"

    # with the delay OFF, the same high-stakes high-conf send fires immediately
    eng_now = AssistantEngine(config=AssistantConfig(autosend_delay_enabled=False),
                              do_apply=lambda r: True)
    r5d = eng_now.handle(kind="Reply draft", stakes=HIGH,
                         confidence=confidence_from_triage(dict(business_relevance=0.97, spam_likelihood=0.0)),
                         subject="fu4", summary="x")
    assert r5d["decision"] == AUTO and r5d["state"] == "auto", r5d

    # approve a pending, then undo an auto
    assert eng.approve(r2["id"]) is True
    assert eng.undo(r1["id"]) is True and eng._find(r1["id"])["state"] == "undone"

    # disabling the engine forces ASK
    cfg.enabled = False
    r6 = eng.handle(kind="Category", stakes=LOW, confidence=0.99, subject="off", summary="x")
    assert r6["decision"] == ASK, r6
    cfg.enabled = True

    # force_auto (Auto-Archive ON): a Cleanup executes immediately (state auto) and
    # is NOT armed/delayed -- archiving sends nothing, so the send-delay is skipped.
    rcl = eng.handle(kind="Cleanup", stakes=HIGH, confidence=0.9, subject="c",
                     summary="archive", force_auto=True, dedup_key="cl1")
    assert rcl["decision"] == AUTO and rcl["state"] == "auto" and "fire_at" not in rcl, rcl
    # ...and with Auto-Archive OFF (hold_for_user), the same Cleanup stays pending
    rcl2 = eng.handle(kind="Cleanup", stakes=HIGH, confidence=0.9, subject="c2",
                      summary="archive", hold_for_user=True, dedup_key="cl2")
    assert rcl2["state"] == "pending", rcl2

    # on_activity fires exactly once per NEW row (host uses it to auto-open the log)
    seen = []
    eng2 = AssistantEngine(config=AssistantConfig(), do_apply=lambda r: True,
                           on_activity=lambda r: seen.append(r["id"]))
    a1 = eng2.handle(kind="Category", stakes=LOW, confidence=0.99, subject="x", summary="x")
    assert seen == [a1["id"]], seen
    eng2.approve(a1["id"]); eng2.undo(a1["id"])   # state changes must NOT re-fire it
    assert seen == [a1["id"]], seen

    c = eng.counts()
    print("self-test OK \xb7", c, "\xb7 classify bands:",
          {"low_auto>=": cfg.low_auto_at, "low_ask<": cfg.low_ask_below,
           "high_auto>=": cfg.high_auto_at, "high_ask<": cfg.high_ask_below})

    # config persistence round-trip
    import tempfile
    cpath = os.path.join(tempfile.gettempdir(), "assistant_cfg_test.json")
    try:
        e2 = AssistantEngine(config=AssistantConfig(low_auto_at=0.71, high_auto_at=0.91), config_path=cpath)
        e2.save_config()
        e3 = AssistantEngine(config=AssistantConfig(low_auto_at=0.50), config_path=cpath)  # different default
        assert abs(e3.config.low_auto_at - 0.71) < 1e-6, e3.config.low_auto_at  # saved value wins
        assert abs(e3.config.high_auto_at - 0.91) < 1e-6, e3.config.high_auto_at
        print("config round-trip OK \xb7 low_auto persisted =", e3.config.low_auto_at)
    finally:
        try:
            os.remove(cpath)
        except Exception:
            pass


if __name__ == "__main__":
    import sys
    if "--selftest" in sys.argv:
        _selftest()
    else:
        _demo()
