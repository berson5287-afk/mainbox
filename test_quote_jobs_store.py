"""
Standalone verification for the quote-job store. No Outlook, no MaINbox import,
no network -- this is the "test the brain in isolation" harness the architecture
was meant to enable. Run:  python test_quote_jobs_store.py
"""
import os
import shutil
import tempfile

from quote_jobs_model import (
    Customer, JobLifecycle, RequestStatus, RFQState, StatusSource, Vendor,
)
from quote_jobs_store import QuoteJobStore, normalize_subject

PASS, FAIL = "  ok  ", "  FAIL"
_results = []


def check(name, cond):
    _results.append(cond)
    print(f"[{PASS if cond else FAIL}] {name}")


def fresh_path():
    d = tempfile.mkdtemp(prefix="qjtest_")
    return d, os.path.join(d, "quote_jobs.json")


def seed(store):
    """A realistic shape: one job, two items, two vendors, overlapping asks."""
    job = store.add_job(Customer(name="Acme Co", email="buyer@acme.test",
                                 conversation_id="CUST-CONVO-1",
                                 source_entry_id="EID-CUST-1"))
    li1 = store.add_line_item(job.job_id, "500MCM THHN CU, black",
                              canonical_key="thhn-500-cu-blk", qty=300, uom="FT",
                              extraction_confidence=0.94)
    li2 = store.add_line_item(job.job_id, '2" EMT conduit',
                              canonical_key="emt-2in", qty=50, uom="EA",
                              extraction_confidence=0.71)
    # vendor A asked for both items; vendor B asked for item 2 only (many-to-many)
    rfqA = store.add_vendor_rfq(job.job_id, Vendor(name="Graybar",
                                email="quotes@graybar.test"),
                                [li1.item_id, li2.item_id])
    rfqB = store.add_vendor_rfq(job.job_id, Vendor(name="WESCO",
                                email="rfq@wesco.test"), [li2.item_id])
    return job, li1, li2, rfqA, rfqB


def test_roundtrip_and_enums():
    d, path = fresh_path()
    try:
        s1 = QuoteJobStore(path)
        job, li1, li2, rfqA, rfqB = seed(s1)
        s1.mark_rfq_sent(job.job_id, rfqA.rfq_id, "VEND-CONVO-A",
                         sent_at="2026-06-05T12:00:00+00:00",
                         deadline_at="2026-06-05T12:30:00+00:00")
        s1.record_item_response(job.job_id, rfqA.rfq_id, li1.item_id,
                                RequestStatus.PRICED, price=4210.00,
                                lead_time="in stock", status_confidence=0.9,
                                raw_excerpt="500MCM ... $14.03/ft, stock")

        s2 = QuoteJobStore(path)  # reload from disk
        j = s2.get_job(job.job_id)
        check("reload finds the job", j is not None)
        check("dict counts survive round-trip",
              len(j.line_items) == 2 and len(j.vendor_rfqs) == 2)
        req = j.vendor_rfqs[rfqA.rfq_id].requests[li1.item_id]
        check("status deserializes to the enum, not a string",
              isinstance(req.status, RequestStatus) and req.status is RequestStatus.PRICED)
        check("rfq state deserializes to enum",
              j.vendor_rfqs[rfqA.rfq_id].state is RFQState.SENT)
        check("lifecycle deserializes to enum", j.lifecycle is JobLifecycle.OPEN)
        check("deadline snapshot persisted",
              j.vendor_rfqs[rfqA.rfq_id].deadline_at == "2026-06-05T12:30:00+00:00")
        check("price + excerpt persisted",
              req.price == 4210.00 and req.raw_excerpt.startswith("500MCM"))
        check("item-2 ask exists under BOTH vendors (many-to-many)",
              li2.item_id in j.vendor_rfqs[rfqA.rfq_id].requests
              and li2.item_id in j.vendor_rfqs[rfqB.rfq_id].requests)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_indexes_survive_reload():
    d, path = fresh_path()
    try:
        s1 = QuoteJobStore(path)
        job, li1, li2, rfqA, rfqB = seed(s1)
        s1.mark_rfq_sent(job.job_id, rfqA.rfq_id, "VEND-CONVO-A")
        hit = s1.find_rfq_by_conversation("VEND-CONVO-A")
        check("vendor convo -> rfq in memory", hit is not None and hit[1].rfq_id == rfqA.rfq_id)
        check("customer convo -> job",
              s1.find_job_by_customer_conversation("CUST-CONVO-1").job_id == job.job_id)

        s2 = QuoteJobStore(path)  # indexes were NOT persisted; must rebuild on load
        hit2 = s2.find_rfq_by_conversation("VEND-CONVO-A")
        check("vendor convo index rebuilt on load",
              hit2 is not None and hit2[1].rfq_id == rfqA.rfq_id)
        check("unsent rfq is absent from convo index",
              s2.find_rfq_by_conversation("nope") is None)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_batch_collapses_writes():
    d, path = fresh_path()
    try:
        store = QuoteJobStore(path)
        writes = {"n": 0}
        real = store._write_to_disk

        def counting():
            writes["n"] += 1
            real()
        store._write_to_disk = counting

        with store.batch():
            job = store.add_job(Customer(name="X", email="x@y.test"))
            for i in range(5):
                store.add_line_item(job.job_id, f"item {i}")
        check("5+ mutations in a batch produce exactly ONE disk write", writes["n"] == 1)

        # exception inside a batch: work-so-far flushes, depth never gets stuck
        writes["n"] = 0
        try:
            with store.batch():
                store.add_line_item(job.job_id, "before boom")
                raise RuntimeError("boom")
        except RuntimeError:
            pass
        check("batch flushes on exception", writes["n"] == 1)
        check("batch depth not stuck after exception", store._batch_depth == 0)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_id_monotonicity():
    d, path = fresh_path()
    try:
        s1 = QuoteJobStore(path)
        a = s1.add_job(Customer(name="A", email="a@a.test"))
        b = s1.add_job(Customer(name="B", email="b@b.test"))
        check("job ids are zero-padded sequential",
              a.job_id == "QJ-00001" and b.job_id == "QJ-00002")
        i1 = s1.add_line_item(a.job_id, "one")
        i2 = s1.add_line_item(a.job_id, "two")
        check("line item ids sequential per job",
              i1.item_id == "LI-1" and i2.item_id == "LI-2")

        s2 = QuoteJobStore(path)  # next id derived from existing data, survives restart
        c = s2.add_job(Customer(name="C", email="c@c.test"))
        check("next job id continues after reload (no separate counter)",
              c.job_id == "QJ-00003")
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_correction_preserves_ai_guess():
    d, path = fresh_path()
    try:
        store = QuoteJobStore(path)
        job, li1, li2, rfqA, rfqB = seed(store)
        store.mark_rfq_sent(job.job_id, rfqA.rfq_id, "VEND-CONVO-A")
        # AI's read
        store.record_item_response(job.job_id, rfqA.rfq_id, li2.item_id,
                                   RequestStatus.PARTIAL, status_confidence=0.4,
                                   raw_excerpt="conduit avail, no price given")
        # user corrects status + adds price, doesn't resend the excerpt
        store.correct_item_response(job.job_id, rfqA.rfq_id, li2.item_id,
                                    status=RequestStatus.PRICED, price=212.50)
        req = store.get_job(job.job_id).vendor_rfqs[rfqA.rfq_id].requests[li2.item_id]
        check("user correction applied", req.status is RequestStatus.PRICED and req.price == 212.50)
        check("status_source flips to USER", req.status_source is StatusSource.USER)
        check("AI's original guess preserved for learning DB",
              req.ai_proposed_status is RequestStatus.PARTIAL)
        check("vendor excerpt not clobbered by correction",
              req.raw_excerpt == "conduit avail, no price given")
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_atomic_write_leaves_no_temp_files():
    d, path = fresh_path()
    try:
        store = QuoteJobStore(path)
        store.add_job(Customer(name="Z", email="z@z.test"))
        leftovers = [f for f in os.listdir(d) if f.endswith(".tmp")]
        check("no .tmp leftovers after atomic save", leftovers == [])
        check("target file exists", os.path.exists(path))
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_newer_schema_refused():
    d, path = fresh_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write('{"schema_version": 999, "jobs": {}}')
        raised = False
        try:
            QuoteJobStore(path)
        except ValueError:
            raised = True
        check("newer schema_version is refused, not silently loaded", raised)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def test_normalizer():
    check("normalizer strips stacked RE/FWD + collapses ws + lowercases",
          normalize_subject("RE: FW:  RFQ  500MCM ") == "rfq 500mcm")
    check("normalizer is total on None", normalize_subject(None) == "")


if __name__ == "__main__":
    test_roundtrip_and_enums()
    test_indexes_survive_reload()
    test_batch_collapses_writes()
    test_id_monotonicity()
    test_correction_preserves_ai_guess()
    test_atomic_write_leaves_no_temp_files()
    test_newer_schema_refused()
    test_normalizer()
    print("-" * 56)
    total, ok = len(_results), sum(_results)
    print(f"{ok}/{total} checks passed")
    raise SystemExit(0 if ok == total else 1)
