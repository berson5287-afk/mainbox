"""
Test MaINboxContactLookup (now symmetric: both sides take an email iterable) and
the category-based vendor/customer builders, with FAKE data shaped like your real
triage history. No app, no DB, no Outlook. Run: python test_mainbox_contacts.py
"""
from mainbox_contacts import (MaINboxContactLookup, build_customer_emails_from_categorized,
                              build_vendor_emails_from_categorized)

PASS, FAIL = "  ok  ", "  FAIL"
_results = []


def check(name, cond):
    _results.append(bool(cond))
    print(f"[{PASS if cond else FAIL}] {name}")


# triage-history shape: rows with sender_email + category (as in the real file)
TRIAGE_ROWS = [
    {"sender_email": "buyer@acme.test", "category": "Quote / Estimate"},          # customer
    {"sender_email": "po@biggencorp.test", "category": "Quote / Estimate"},        # customer
    {"sender_email": "markh@brazill.com", "category": "Vendor / Supplier"},        # vendor
    {"sender_email": "sales@cooper.test", "category": "Vendor / Supplier"},        # vendor
    {"sender_email": "sales@americanpoweresc.com", "category": "Quote / Estimate"},# OWN domain -> excluded
    {"sender_email": "noise@whoever.test", "category": "General"},                 # neither
    {"no_email": True},                                                            # malformed -> skipped
]


def vendors():
    return build_vendor_emails_from_categorized(TRIAGE_ROWS, exclude_domains=("americanpoweresc.com",))


def customers():
    return build_customer_emails_from_categorized(TRIAGE_ROWS, exclude_domains=("americanpoweresc.com",))


def test_builders_split_by_category():
    v, c = vendors(), customers()
    check("vendor set = 'Vendor / Supplier' rows", v == {"markh@brazill.com", "sales@cooper.test"})
    check("customer set = 'Quote / Estimate' rows (own-domain excluded)", c == {"buyer@acme.test", "po@biggencorp.test"})
    check("'General' rows are neither", "noise@whoever.test" not in v and "noise@whoever.test" not in c)
    check("own-domain excluded from customers", "sales@americanpoweresc.com" not in c)


def test_lookup_recognizes_both_sides():
    lk = MaINboxContactLookup(vendor_provider=vendors, customer_provider=customers)
    check("known vendor by exact email", lk.is_known_vendor("markh@brazill.com", "brazill.com"))
    check("known customer by exact email", lk.is_known_customer("buyer@acme.test", "acme.test"))
    check("vendor is not a customer", not lk.is_known_customer("markh@brazill.com", "brazill.com"))
    check("customer is not a vendor", not lk.is_known_vendor("buyer@acme.test", "acme.test"))
    check("unknown sender is neither",
          not lk.is_known_vendor("x@nowhere.test", "nowhere.test")
          and not lk.is_known_customer("x@nowhere.test", "nowhere.test"))


def test_domain_matching():
    on = MaINboxContactLookup(vendors, customers, match_domains=True)
    off = MaINboxContactLookup(vendors, customers, match_domains=False)
    check("new rep at known vendor domain recognized (match_domains=True)",
          on.is_known_vendor("newguy@brazill.com", "brazill.com"))
    check("not recognized with match_domains=False",
          not off.is_known_vendor("newguy@brazill.com", "brazill.com"))


def test_robust_to_empty():
    lk = MaINboxContactLookup(lambda: [], lambda: [])
    check("empty providers -> nobody known, no crash",
          not lk.is_known_vendor("a@b.test", "b.test") and not lk.is_known_customer("a@b.test", "b.test"))
    check("None providers -> no crash",
          not MaINboxContactLookup(lambda: None, lambda: None).is_known_customer("a@b.test", "b.test"))


def test_interface_matches_resolver():
    import os, tempfile, shutil
    from quote_jobs_store import QuoteJobStore
    from inbound_router import InboundMail, resolve_features, KeywordRfqScorer
    d = tempfile.mkdtemp()
    try:
        store = QuoteJobStore(os.path.join(d, "q.json"))
        lk = MaINboxContactLookup(vendors, customers)
        mail = InboundMail("M", "buyer@acme.test", subject="Request for quote",
                           body="please quote price and availability")
        feats = resolve_features(mail, store, lk, KeywordRfqScorer())
        check("lookup drives resolver: customer flagged, not vendor",
              feats.sender_is_known_customer and not feats.sender_is_known_vendor)
        mail2 = InboundMail("M2", "markh@brazill.com", subject="Re: P&A", body="here is pricing")
        feats2 = resolve_features(mail2, store, lk, KeywordRfqScorer())
        check("lookup drives resolver: vendor flagged, not customer",
              feats2.sender_is_known_vendor and not feats2.sender_is_known_customer)
    finally:
        shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    test_builders_split_by_category()
    test_lookup_recognizes_both_sides()
    test_domain_matching()
    test_robust_to_empty()
    test_interface_matches_resolver()
    print("-" * 56)
    total, ok = len(_results), sum(_results)
    print(f"{ok}/{total} checks passed")
    raise SystemExit(0 if ok == total else 1)
