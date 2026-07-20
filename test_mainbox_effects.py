"""
Tests for the production EffectsPort -- the parts that DON'T need Outlook:
the SmartScan row->item mapper, the JSONL event log, the injected extraction
hook, the no-hook guard, popup safety off-Windows, and protocol coverage.
The two draft methods need live Outlook and are exercised in the layer-2
walkthrough on your machine, not here.  Run: python test_mainbox_effects.py
"""
import json
import os
import shutil
import tempfile

from mainbox_effects import MaINboxEffectsPort, smartscan_rows_to_items
from quote_jobs_controller import ExtractedItem

PASS, FAIL = "  ok  ", "  FAIL"
_results = []


def check(name, cond):
    _results.append(bool(cond))
    print(f"[{PASS if cond else FAIL}] {name}")


class _Item:
    def __init__(self, **k): self.__dict__.update(k)


class _Job:
    """Minimal stand-in for the dataclass job (only what the logged effects touch)."""
    def __init__(self):
        self.job_id = "JOB-1"
        self.customer = _Item(email="buyer@acme.test", source_entry_id=None)
        self.line_items = {}


# ----------------------------- the mapper -----------------------------
def test_mapper_basic_fields():
    rows = [{"qty": "300", "unit": "FT", "description": "500MCM THHN",
             "part_number": "THHN-500", "status": "green"}]
    items = smartscan_rows_to_items(rows)
    it = items[0]
    check("maps description", it.description == "500MCM THHN")
    check("part_number -> canonical_key (lowercased)", it.canonical_key == "thhn-500")
    check("qty parsed to float", it.qty == 300.0)
    check("uom carried from 'unit'", it.uom == "FT")
    check("green status -> high confidence (clears 0.85 gate)", it.extraction_confidence >= 0.85)


def test_mapper_status_and_edges():
    rows = [
        {"description": "Item Y", "status": "yellow"},                 # below gate
        {"description": "Item R", "status": "red"},                    # below gate
        {"description": "Item U", "status": "weird-unknown-status"},   # default 0.8
        {"description": "", "part_number": ""},                        # empty -> skipped
        "not-a-dict",                                                  # junk -> skipped
        {"qty": "12 boxes", "description": "Box thing"},               # qty from messy string
    ]
    items = smartscan_rows_to_items(rows)
    by_desc = {i.description: i for i in items}
    check("empty + junk rows skipped (4 valid remain)", len(items) == 4)
    check("yellow below gate", by_desc["Item Y"].extraction_confidence < 0.85)
    check("red below gate", by_desc["Item R"].extraction_confidence < 0.85)
    check("unknown status -> 0.8 default", abs(by_desc["Item U"].extraction_confidence - 0.8) < 1e-9)
    check("qty extracted from messy string", by_desc["Box thing"].qty == 12.0)
    check("no part_number -> canonical_key None", by_desc["Item Y"].canonical_key is None)
    check("custom status map honored",
          smartscan_rows_to_items([{"description": "x", "status": "A"}],
                                  status_confidence={"a": 0.42})[0].extraction_confidence == 0.42)


# --------------------------- the event log ---------------------------
def test_event_log_writes_jsonl():
    d = tempfile.mkdtemp()
    try:
        log = os.path.join(d, "events.jsonl")
        port = MaINboxEffectsPort(event_log_path=log)
        job = _Job()
        port.flag_item_quoted(job, "LI-1")
        port.request_item_review(job, ["LI-2", "LI-3"])
        lines = [json.loads(x) for x in open(log, encoding="utf-8").read().splitlines()]
        check("two effects -> two JSONL lines", len(lines) == 2)
        check("line carries effect + job_id", lines[0]["effect"] == "flag_item_quoted" and lines[0]["job_id"] == "JOB-1")
        check("details captured (item id)", lines[0]["item"] == "LI-1")
        check("each line has a timestamp", all("ts" in l for l in lines))
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_log_without_path_is_console_only():
    port = MaINboxEffectsPort(event_log_path=None)
    port.flag_item_quoted(_Job(), "LI-9")     # must not raise, just prints
    check("no log path -> console only, no crash", True)


# ----------------------- extraction hook path ------------------------
def test_run_extraction_with_hook():
    rows = [{"qty": "2", "unit": "EA", "description": "Breaker", "part_number": "BR-2", "status": "green"},
            {"qty": "10", "description": "Conduit", "status": "yellow"}]
    port = MaINboxEffectsPort(extraction_hook=lambda job: rows)
    items = port.run_extraction(_Job())
    check("hook rows -> mapped ExtractedItems", len(items) == 2 and isinstance(items[0], ExtractedItem))
    check("hook output flows through the mapper", items[0].canonical_key == "br-2")


def test_run_extraction_without_hook_raises():
    port = MaINboxEffectsPort(extraction_hook=None)
    try:
        port.run_extraction(_Job())
        check("missing hook raises NotImplementedError", False)
    except NotImplementedError:
        check("missing hook raises NotImplementedError", True)


# --------------------- safety + protocol surface ---------------------
def test_alerts_headless_log_without_blocking():
    # enable_popups=False -> no modal dialog; alerts must still log. (Popups-on is
    # verified manually on Windows; an automated modal box would block the suite.)
    d = tempfile.mkdtemp()
    try:
        log = os.path.join(d, "events.jsonl")
        port = MaINboxEffectsPort(event_log_path=log, enable_popups=False)
        job = _Job()
        port.notify_response_needed(job, "RFQ-1", ["LI-1"])
        port.alert_overdue(job, "RFQ-1")
        port.suggest_customer_quote(job)
        effects = [json.loads(x)["effect"] for x in open(log, encoding="utf-8").read().splitlines()]
        check("headless alerts log without popping a modal (no block, no crash)",
              effects == ["notify_response_needed", "alert_overdue", "suggest_customer_quote"])
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_implements_full_port_surface():
    methods = ["draft_ack", "draft_vendor_rfq", "run_extraction", "request_item_review",
               "request_vendor_selection", "flag_item_quoted", "notify_response_needed",
               "alert_overdue", "suggest_customer_quote"]
    port = MaINboxEffectsPort()
    check("all nine EffectsPort methods present and callable",
          all(callable(getattr(port, m, None)) for m in methods))


if __name__ == "__main__":
    test_mapper_basic_fields()
    test_mapper_status_and_edges()
    test_event_log_writes_jsonl()
    test_log_without_path_is_console_only()
    test_run_extraction_with_hook()
    test_run_extraction_without_hook_raises()
    test_alerts_headless_log_without_blocking()
    test_implements_full_port_surface()
    print("-" * 56)
    total, ok = len(_results), sum(_results)
    print(f"{ok}/{total} checks passed")
    raise SystemExit(0 if ok == total else 1)
