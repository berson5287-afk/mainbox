"""
Recording fake ContactLookup -- construct with known customer/vendor emails and
domains; matches on either. Used by the inbound test harness and handy for REPL
dry-runs (pair it with KeywordRfqScorer for a fully offline front door).
"""
from typing import Iterable, Set


class RecordingContactLookup:
    def __init__(self, customers: Iterable[str] = (), vendors: Iterable[str] = ()):
        self.customers: Set[str] = {c.lower() for c in customers}
        self.vendors: Set[str] = {v.lower() for v in vendors}

    def is_known_customer(self, email: str, domain: str) -> bool:
        return (email or "").lower() in self.customers or (domain or "").lower() in self.customers

    def is_known_vendor(self, email: str, domain: str) -> bool:
        return (email or "").lower() in self.vendors or (domain or "").lower() in self.vendors
