"""
LAYER 2 -- overdue-alert proof, end to end with the REAL clock.

Run on Windows with Outlook desktop open:

    python layer2_overdue_test.py

Uses a 1-minute editable timeout (so you don't wait 30 real minutes), sends an
RFQ, waits PAST the deadline with no reply, then ticks. You should get exactly
ONE "Vendor Overdue" popup; a second tick stays silent (the fire-once guard).

Needs the same 5 files in this folder as layer2_walkthrough.py.
"""
import os
import tempfile
import time
from datetime import datetime, timezone

from quote_jobs_store import QuoteJobStore
from quote_jobs_controller import QuoteJobController, ControllerConfig
from quote_jobs_model import Customer, Vendor
from mainbox_port import OutlookEffectsPort


def main() -> None:
    store_path = os.path.join(tempfile.gettempdir(), "qj_layer2_overdue.json")
    if os.path.exists(store_path):
        os.remove(store_path)

    # 1-minute editable timeout -- this is the Fork-D knob
    ctl = QuoteJobController(
        QuoteJobStore(store_path),
        OutlookEffectsPort(display_drafts=False),   # don't pop drafts; we only care about the alert
        ControllerConfig(default_timeout_minutes=1))

    job, follow = ctl.ingest_customer_rfq(Customer(name="Acme", email="customer@example.com"))
    for ev in follow:
        ctl.dispatch(ev)
    iid = list(ctl.store.get_job(job.job_id).line_items)[0]
    ctl.select_vendor(job.job_id, Vendor(name="Graybar", email="vendor@example.com"), [iid])
    rfq_id = list(ctl.store.get_job(job.job_id).vendor_rfqs)[-1]

    ctl.mark_rfq_sent(job.job_id, rfq_id, "CONVO-OVERDUE")   # deadline = now + 1 min
    rfq = ctl.store.get_job(job.job_id).vendor_rfqs[rfq_id]
    print(f"RFQ sent.  deadline_at = {rfq.deadline_at}")
    print("Waiting 65s for the deadline to pass (no reply will arrive)...")
    for s in range(65, 0, -5):
        print(f"  {s:>2}s remaining ...", flush=True)
        time.sleep(5)

    print("\nTick #1 (now past the deadline) -> expect ONE 'Vendor Overdue' popup:")
    ctl.tick(datetime.now(timezone.utc))

    print("\nTick #2 -> expect SILENCE (alert already fired, guard held):")
    ctl.tick(datetime.now(timezone.utc))

    print(f"\nDONE. Store file: {store_path}")


if __name__ == "__main__":
    main()
