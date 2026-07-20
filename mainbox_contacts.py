"""
Real ContactLookup for MaINbox -- answers "known customer? / known vendor?" from
your own data, WITHOUT importing your app and WITHOUT guessing your storage layer.

Two injected providers (dependency injection), because real-data inspection showed
the two questions have DIFFERENT best sources:

  * vendor_provider()   -> {"vendors": {email: ...}, ...}   (your vendor_history;
                           confirmed email-keyed and working)
  * customer_provider() -> iterable of customer email addresses

Why split them: vendor_history is email-keyed and reliable for vendors, but its
"customers" dict is NOT email-keyed (it was keyed by subject in real data), so the
customer side must come from your category-tagged mail instead -- e.g. senders you
classified "Quote / Estimate". build_customer_emails_from_categorized() below is a
reference that derives that set from rows carrying (sender_email, category), with
your own domain(s) excluded so internal mail never counts as a customer.

Wiring at integration (inside MaINbox, self in scope), for example:

    from mainbox_contacts import MaINboxContactLookup, build_customer_emails_from_categorized
    lookup = MaINboxContactLookup(
        vendor_provider=lambda: self.vendor_history,
        customer_provider=lambda: build_customer_emails_from_categorized(
            self.load_triage_history(),            # any rows with sender_email + category
            customer_categories=("Quote / Estimate",),
            exclude_domains=("americanpoweresc.com",)),
    )

The matching logic here is pure and fully testable with fakes (see the test).
"""
from __future__ import annotations

from typing import Callable, Dict, Iterable, Set


def _domains(emails: Iterable[str]) -> Set[str]:
    out = set()
    for e in emails:
        e = (e or "").lower()
        if "@" in e:
            out.add(e.split("@")[-1])
    return out


def build_customer_emails_from_categorized(rows, customer_categories=("Quote / Estimate",),
                                           exclude_domains=()) -> Set[str]:
    """Reference helper: derive a set of customer emails from category-tagged rows.

    `rows` is any iterable of dicts that carry a sender email and a category
    (e.g. your triage history). A row counts as a customer if its category is in
    customer_categories; rows whose domain is in exclude_domains (your own /
    internal addresses) are dropped."""
    return _emails_for_categories(rows, customer_categories, exclude_domains)


def build_vendor_emails_from_categorized(rows, vendor_categories=("Vendor / Supplier",),
                                         exclude_domains=()) -> Set[str]:
    """Same as the customer builder, for vendors -- derives known vendor emails
    from rows tagged 'Vendor / Supplier'. Use this when vendor_history is empty
    (the durable history only fills as you transact through MaINbox)."""
    return _emails_for_categories(rows, vendor_categories, exclude_domains)


def _emails_for_categories(rows, categories, exclude_domains) -> Set[str]:
    wanted = {c.lower() for c in categories}
    excl = {d.lower() for d in exclude_domains}
    out: Set[str] = set()
    for r in rows or []:
        if not isinstance(r, dict):
            continue
        email = str(r.get("sender_email") or r.get("sender") or "").lower().strip()
        category = str(r.get("category") or "").lower().strip()
        if not email or "@" not in email:
            continue
        if email.split("@")[-1] in excl:
            continue
        if category in wanted:
            out.add(email)
    return out


class MaINboxContactLookup:
    def __init__(self, vendor_provider: Callable[[], Iterable[str]],
                 customer_provider: Callable[[], Iterable[str]],
                 match_domains: bool = True):
        # both providers return an iterable of known email addresses
        # (vendors and customers respectively)
        self._vendor_provider = vendor_provider
        self._customer_provider = customer_provider
        self.match_domains = match_domains

    def _vendor_set(self) -> Set[str]:
        return {str(e).lower() for e in (self._vendor_provider() or [])}

    def _customer_set(self) -> Set[str]:
        return {str(e).lower() for e in (self._customer_provider() or [])}

    def is_known_vendor(self, email: str, domain: str) -> bool:
        vendors = self._vendor_set()
        e, d = (email or "").lower(), (domain or "").lower()
        if e in vendors:
            return True
        return self.match_domains and bool(d) and d in _domains(vendors)

    def is_known_customer(self, email: str, domain: str) -> bool:
        customers = self._customer_set()
        e, d = (email or "").lower(), (domain or "").lower()
        if e in customers:
            return True
        return self.match_domains and bool(d) and d in _domains(customers)
