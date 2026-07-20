"""MaINbox PO-Match Engine.

A conservative, dependency-free detector that answers two questions about an
inbound document/email:

  1. Is this a customer PURCHASE ORDER (an award), as opposed to an RFQ, a vendor
     quote, an invoice, a submittal, or ordinary mail?
  2. If so, WHICH open Quote Coverage job does it belong to?

The engine never touches Outlook, Tk, the filesystem, or any MaINbox state. It
takes plain data in and returns a plain verdict out, so it can be unit-tested
against real PO PDFs and reused as a standalone module.

Matching is best-match-wins across four independent signal families, in
descending trust:

    quote_number  - the seller's own quote number printed on the PO
                    (e.g. "PER QUOTE#S100099955"). This is the strongest link
                    because the number originated from US and appears on the PO
                    only because the customer is ordering against our quote.
    job_number    - a shared job/project number (e.g. "JOB# 260205") that also
                    appears in the job's group name / RFQ subject.
    po_number     - a PO number the job already recorded (rare, but if a prior
                    message mentioned "PO 12345" and this PO carries the same
                    number, that is a strong tie).
    material      - overlap between the PO's line items and the job's requested
                    items (part numbers first, then normalized descriptions).

A job is returned as a confident match only when the winning evidence clears a
high bar AND is clearly ahead of the runner-up job, so a PO is never auto-applied
to the wrong ledger. Everything below that bar is reported as review/none for a
human to resolve. The engine fails safe: any internal error yields a NONE verdict
(the caller then does nothing automatic, exactly as if the PO had not been
recognized).

Designed for MaINbox v4.1.38+ but usable standalone.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
import re
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

ENGINE_VERSION = "1.8.3"
# v1.8.3 (tenth audit): future-PO veto NARROWED -- will-be-sent/provided/released
# require explicit authorization context; operational language ("released for
# shipment/to production", "sent via UPS", "provided in two releases") on a real
# PO no longer creates a missed-PO false negative. issued/to-follow/forthcoming
# unchanged.
# v1.8.2 (ninth audit): subject/title interaction fix -- the title negation scan
# now inspects the first line that MATCHES a title regex instead of breaking on
# any line beginning with PO, so "Subject: PO attached" no longer shields an
# attachment saying "PO #12345 will be issued next week" / "to follow" from the
# future-PO suppression; object-bound will-be-issued/to-follow/forthcoming forms
# also join the title-independent state vetoes as defense in depth.
# v1.8.1: STATUS separator accepts "=" ("STATUS = CANCELLED"); "not issued" /
# "not placed" join the title-independent not-yet-authorized family so
# "PO #12345 NOT ISSUED" is not an award.
# v1.8.0 (eighth external audit): (1) TWO-TIER TITLE PRIORITY -- a spelled-out
# PURCHASE ORDER title keeps its full authority, but an abbreviated "PO #12345"
# line is reference-grade: it is demoted by any explicit non-PO document-title
# line (INVOICE / PACKING SLIP / ORDER ACKNOWLEDGEMENT / SALES ORDER / QUOTATION /
# SUBMITTAL ...), never bypasses the non-PO heading or role-identity vetoes, and
# the quote-family PO-number rescue now requires a full title. A "pure title
# line" test (phrase + at most cue and one identifier) keeps instruction text
# ("ORDER ACKNOWLEDGEMENT REQUIRED WITHIN 24 HOURS") from demoting, preserving
# the v1.7.0 abbreviated-PO fix. (2) TERMINAL STATUS fields and stamps --
# "STATUS: CANCELLED/VOID/REJECTED/CLOSED/EXPIRED/..." and standalone terminal
# stamps veto even under a full PURCHASE ORDER title.
# v1.7.0 (seventh external audit): (1) TITLE RECOGNIZER accepts abbreviated forms
# -- "PO #12345", "P.O. #12345", "P/O #12345", "PO 12345" are authoritative titles
# (with an explicit cue or digit-bearing token required, so PO STATUS / PO REQUEST
# / PO ACKNOWLEDGEMENT / PO CANCELLATION NOTICE are never titles and keep their
# role vetoes); the negation line-scan ("PO #12345 to follow") covers abbreviated
# lines too. (2) AUTHORIZATION-STATE VETOES are now TITLE-INDEPENDENT: the limbo
# states (pending signature/funding/execution/release, on hold, not released/
# authorized/approved) moved out of the title-bypassed soft tier into
# _PO_STATE_VETO_LANGUAGE plus STATUS-field and standalone-stamp forms in the
# heading zone -- "PURCHASE ORDER 12345 / STATUS: PENDING SIGNATURE" no longer
# matches; generic terms about other objects ("submittals pending approval") are
# untouched. (3) PO REQUEST/REQUISITION/INQUIRY/ACKNOWLEDGEMENT added to the
# abbreviated role-title vetoes.
# v1.6.0 (sixth external audit): (1) TITLE RECOGNIZER accepts an identifier that
# follows the words directly -- "PURCHASE ORDER 12345" is now a title, so a real PO
# whose terms say "ORDER ACKNOWLEDGEMENT REQUIRED WITHIN 24 HOURS" is no longer
# re-roled as an acknowledgment (false negative; a missed PO gets negative-
# checkpointed). The lookahead requires a digit-bearing token, so PURCHASE ORDER
# REQUEST/STATUS/ACKNOWLEDGEMENT keep their role vetoes. (2) The "P/O" spelling is
# accepted everywhere the PO abbreviation appears -- extraction, headings, and every
# lifecycle expression -- so "P/O #12345" extracts and a duplicate P/O of an already-
# recorded number now hits the existing-PO follow-up gate instead of auto-matching.
# (3) Lifecycle states completed: hard object-bound denied/expired/closed/deleted/
# terminated/nullified, and soft pending-or-awaiting signature/funding/release/
# execution plus not-funded/released/signed/executed/authorized.
# v1.5.0 (fifth external audit): (1) EXISTING-PO FOLLOW-UP GATE -- when every PO
# number on the document is already recorded on the winning job, the verdict is
# forced to REVIEW ("follow-up about an already-recorded PO" or, with
# revision/change-order language, "PO revision/change order"), never an automatic
# MATCH: status requests, receipt confirmations, and duplicate copies of a PO the
# job already has no longer create a second award; a genuinely NEW PO number
# still matches normally. (2) Soft non-award language for PO chatter: "status
# on/of PO", "confirm receipt of PO", "copy of PO", "PO acknowledgement" -- a
# body-only "Can you provide status on PO #12345?" is no longer a PO (an attached
# real PO with a PURCHASE ORDER title still classifies via the title bypass).
# (3) Lifecycle roles completed: PO CANCELLATION NOTICE, (PURCHASE) ORDER
# ACKNOWLEDGEMENT, PO REJECTION/REVOCATION/WITHDRAWAL notices, PO RECEIPT
# CONFIRMATION, PO STATUS REQUEST/UPDATE/REPORT title headings, plus hard
# object-bound "PO #x has been rejected/revoked/withdrawn/rescinded/declined"
# and "purchase order cancellation".
# v1.4.0 (fourth external audit): (1) free-floating hard vetoes (do-not-process /
# not-approved / pro-forma / for-review-only) are OBJECT-SCOPED to the sentence/line
# that contains them -- they no longer reject a valid PO whose terms say "do not
# process INVOICES without this PO number" or "SUBSTITUTIONS are not approved";
# (2) lifecycle auxiliary coverage (was/have-been cancelled, has been voided,
# superseded, for-reference-only) and object-bound soft states (PO pending approval /
# on hold / awaiting release); (3) role/state TITLE headings veto: ORDER
# ACKNOWLEDGEMENT, ORDER CONFIRMATION, SALES ORDER, PURCHASE REQUISITION,
# PRELIMINARY/PROPOSED PURCHASE ORDER, PURCHASE ORDER REQUEST/STATUS/INQUIRY --
# a PO number plus quote reference on those documents is not a new customer PO;
# (4) "P.O. Box 12345" is no longer extracted as PO number "BOX12345".


class PODecision(str, Enum):
    MATCH = "match"          # confident: this PO belongs to exactly this job
    REVIEW = "review"        # looks like a PO but the job link is not certain
    NOT_PO = "not_po"        # the document is not a customer purchase order
    NONE = "none"            # nothing actionable (also the fail-safe result)


# ---------------------------------------------------------------------------
# Inputs
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class POContext:
    """Everything the engine needs to know about the inbound message + document.

    text combines the email subject/body and the extracted attachment text; the
    engine does its own field extraction from it, but callers may also pass
    pre-extracted values (they win over regex extraction when present).
    """
    subject: str = ""
    body: str = ""
    attachment_text: str = ""
    sender_email: str = ""
    sender_name: str = ""
    sender_type: str = ""            # durable C/V/E/M if MaINbox knows it
    internal_domains: Tuple[str, ...] = ()
    has_attachments: bool = False
    # Optional pre-extracted signals (a caller that already parsed the PDF may
    # supply these; empty means "let the engine find them in the text").
    stated_quote_numbers: Tuple[str, ...] = ()
    stated_po_numbers: Tuple[str, ...] = ()
    stated_job_numbers: Tuple[str, ...] = ()
    line_parts: Tuple[str, ...] = ()      # item codes on the PO, if pre-parsed
    line_descriptions: Tuple[str, ...] = ()


@dataclass(frozen=True)
class OpenJob:
    """A single open Quote Coverage ledger the PO might belong to.

    All fields are optional; the engine matches on whatever is present. text
    should carry any free-text the job is known by (group name, RFQ subject,
    original request subject) so quote/job numbers embedded there are found.
    """
    job_id: str = ""                 # opaque handle the caller uses to act (e.g. "grp:g12")
    label: str = ""                  # human name for logging (group/subject)
    text: str = ""                   # group name + subjects + any RFQ text
    customer_email: str = ""
    quote_numbers: Tuple[str, ...] = ()
    job_numbers: Tuple[str, ...] = ()
    po_numbers: Tuple[str, ...] = ()
    requested_parts: Tuple[str, ...] = ()
    requested_descriptions: Tuple[str, ...] = ()
    # v4.1.39a: True when this job already has a recorded PO. It stays eligible so a
    # change order / partial release / revised PO can still match, but the engine
    # applies a mild penalty so it never outranks a genuinely fresh open job.
    prior_po: bool = False


# ---------------------------------------------------------------------------
# Outputs
# ---------------------------------------------------------------------------

@dataclass
class POVerdict:
    decision: PODecision
    job_id: str = ""
    job_label: str = ""
    confidence: float = 0.0
    is_po: bool = False
    po_numbers: List[str] = field(default_factory=list)
    quote_numbers: List[str] = field(default_factory=list)
    job_numbers: List[str] = field(default_factory=list)
    reasons: List[str] = field(default_factory=list)
    runner_up_label: str = ""
    runner_up_score: float = 0.0
    customer_conflict: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return {
            "decision": self.decision.value,
            "job_id": self.job_id,
            "job_label": self.job_label,
            "confidence": round(self.confidence, 4),
            "is_po": self.is_po,
            "po_numbers": list(self.po_numbers),
            "quote_numbers": list(self.quote_numbers),
            "job_numbers": list(self.job_numbers),
            "reasons": list(self.reasons),
            "runner_up_label": self.runner_up_label,
            "runner_up_score": round(self.runner_up_score, 4),
            "customer_conflict": self.customer_conflict,
            "engine_version": ENGINE_VERSION,
        }


# ---------------------------------------------------------------------------
# Tunables
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class POPolicy:
    # A document must reach this WEIGHTED po-ness score to be considered a PO
    # (v4.1.39). A strong anchor -- a "purchase order" heading (5.0) or a labeled PO
    # number (4.0) -- clears it alone; generic ship-to/bill-to/buyer/terms fields
    # (0.5 each, ~2.5 max) never do. This is the audit's weighted-model fix.
    min_po_document_score: float = 4.0
    # Score awarded per evidence family when it matches a job.
    weight_quote_number: float = 1.00
    weight_po_number: float = 0.80
    weight_job_number: float = 0.70
    weight_material_full: float = 0.55       # all PO lines found in the job
    # v4.1.42 strong-identifier gate: without a quote/PO/job-number link, material
    # evidence is capped so it cannot reach the auto-match threshold on its own --
    # a single exact part (0.35 + 0.15 customer = 0.50) or description-only overlap
    # (0.28 + 0.15 = 0.43) both land in Review; >= 2 exact parts keep full weight.
    weight_material_single_part: float = 0.35
    weight_material_desc_only: float = 0.28
    # v4.1.39: a verified customer-address agreement is mild positive corroboration
    # (never a sole basis for a match); a confirmed customer CONFLICT does not
    # subtract score but forces the winning job to Review (see _evaluate).
    weight_customer_exact: float = 0.15
    weight_customer_domain: float = 0.08
    # v4.1.39a: subtracted from a candidate that already has a recorded PO so it
    # ranks below fresh open jobs but can still match a change order / partial release.
    penalty_prior_po: float = 0.20
    # Confidence needed to auto-apply (MATCH). Below this but with any positive
    # evidence -> REVIEW.
    min_match_confidence: float = 0.70
    # The winner must beat the runner-up job by at least this margin to auto-apply.
    min_lead_over_runner_up: float = 0.30
    reject_internal_senders: bool = True


# ---------------------------------------------------------------------------
# Text helpers (pure)
# ---------------------------------------------------------------------------

def _norm(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip()


def _lower(s: Any) -> str:
    return _norm(s).lower()


def get_email_domain(addr: str) -> str:
    a = str(addr or "").strip().lower()
    if "@" in a:
        return a.rsplit("@", 1)[-1]
    return ""


def is_internal_sender(sender_email: str, internal_domains: Sequence[str]) -> bool:
    dom = get_email_domain(sender_email)
    if not dom:
        return False
    return dom in {str(d or "").strip().lower() for d in internal_domains if d}


# Free/public email domains -- a shared public domain (both gmail.com) is NOT
# evidence that two addresses are the same customer, so domain-level corroboration
# is suppressed for these.
_PUBLIC_EMAIL_DOMAINS = frozenset({
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "aol.com", "icloud.com",
    "live.com", "msn.com", "comcast.net", "verizon.net", "att.net", "me.com",
    "protonmail.com", "ymail.com", "mail.com", "gmx.com",
})


def _email_identity_relation(sender_email: str, job_customer_email: str) -> str:
    """Compare a PO's sender to a job's known customer address. Returns one of:
        "exact"    - same address (strong corroboration)
        "domain"   - same NON-public company domain (moderate corroboration)
        "conflict" - both known, different NON-public domains (a real mismatch)
        "unknown"  - one or both missing, or only a shared public domain
    A forwarded PO (contractor/coworker relays the buyer's order) legitimately
    produces "unknown" or "conflict"; the caller must not treat conflict as proof
    the match is wrong, only as a reason to confirm rather than auto-apply."""
    s = str(sender_email or "").strip().lower()
    j = str(job_customer_email or "").strip().lower()
    if not s or not j:
        return "unknown"
    if s == j:
        return "exact"
    sd = get_email_domain(s)
    jd = get_email_domain(j)
    if not sd or not jd:
        return "unknown"
    if sd == jd:
        return "unknown" if sd in _PUBLIC_EMAIL_DOMAINS else "domain"
    # Different domains. Only a CONFLICT when neither is a public mailbox (two
    # distinct company domains is a genuine customer mismatch; a personal gmail
    # forwarding a corporate PO is not).
    if sd in _PUBLIC_EMAIL_DOMAINS or jd in _PUBLIC_EMAIL_DOMAINS:
        return "unknown"
    return "conflict"


# Quote numbers: alphanumeric tokens that follow a quote-reference cue. Kept
# deliberately specific -- a quote number is only trusted when it is explicitly
# labelled as one ("per quote# S100099955", "quote no S100099955"), never a bare
# number floating in the text.
_QUOTE_CUE_RE = re.compile(
    r"(?:per\s+)?(?:quote|quotation|qte|qt|ref(?:erence)?)\s*(?:#|no\.?|number|:)*\s*"
    r"([A-Za-z]{0,3}\d[A-Za-z0-9]{3,})",
    re.I)

# PO numbers: labelled purchase-order numbers. v4.1.41: the spaced letter-prefix
# (P.O. # A 12345) is only consumed when a DIGIT follows it, and the final token
# must contain a digit (enforced in extract_po_numbers) -- v4.1.40's version
# captured prose ("po for job 123" -> phantom "FORJOB123"). Slash/period
# separators inside the number (123/45) are kept. The generic word "order" only
# counts with an explicit #/no./number/colon so ordinary prose never matches.
_PO_CUE_RE = re.compile(
    r"(?:"
    # v1.4.0: "(?!\s*box\b)" -- "P.O. Box 12345" is a mailing address, not a PO
    # number; it previously extracted as "BOX12345" and contaminated PO history,
    # dedupe, and job matching.
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)(?!\s*box\b)\s*(?:#|no\.?|number|:)*\s*"
    r"((?:[A-Za-z]{1,3}\s(?=\d))?(?=[A-Za-z0-9\-/.]*\d)[A-Za-z0-9][A-Za-z0-9\-/.]*(?:\s?\d[A-Za-z0-9\-/.]*)?)"
    r"|"
    r"\border\s*(?:#|no\.?|number|:)+\s*"
    r"((?=[A-Za-z0-9\-/.]*\d)[A-Za-z0-9][A-Za-z0-9\-/.]*)"
    r")",
    re.I)

# Job / project numbers: labelled job identifiers.
_JOB_CUE_RE = re.compile(
    r"\b(?:job|project|proj|contract)\s*(?:#|no\.?|number|:)*\s*"
    r"([A-Za-z]{0,4}\d{3,}[A-Za-z0-9\-]*)",
    re.I)


def _clean_token(tok: str) -> str:
    """Normalize an identifier for comparison: uppercase, strip separators."""
    return re.sub(r"[\s\-_./#]", "", str(tok or "")).upper()


def extract_quote_numbers(text: str) -> List[str]:
    out = []
    seen = set()
    for m in _QUOTE_CUE_RE.finditer(text or ""):
        tok = _clean_token(m.group(1))
        # A trustworthy quote number carries at least one letter OR is 5+ digits;
        # a bare 3-4 digit number is too collision-prone to be a quote key.
        if tok and (any(c.isalpha() for c in tok) or len(tok) >= 5) and tok not in seen:
            seen.add(tok)
            out.append(tok)
    return out


def extract_po_numbers(text: str) -> List[str]:
    out = []
    seen = set()
    for m in _PO_CUE_RE.finditer(text or ""):
        tok = _clean_token(m.group(1) or m.group(2) or "")
        # v1.4.0: reject mailing-address captures outright ("P.O. Box 12345" ->
        # "BOX12345") in case the token arrives via the order-number alternative.
        if re.match(r"^BOX\d", tok):
            continue
        # v4.1.41: a PO number must contain a digit -- hard filter against prose
        # fragments ("po for ..." captures). 2-char minimum allows "PO # 12".
        if tok and len(tok) >= 2 and any(c.isdigit() for c in tok) and tok not in seen:
            seen.add(tok)
            out.append(tok)
    return out


def extract_job_numbers(text: str) -> List[str]:
    out = []
    seen = set()
    for m in _JOB_CUE_RE.finditer(text or ""):
        tok = _clean_token(m.group(1))
        if tok and len(tok) >= 4 and tok not in seen:
            seen.add(tok)
            out.append(tok)
    return out


# Bare-number fallback for a job's own text (group names are often just the job
# number, e.g. "260205 7/17/26 monday 7am delivery"). Only used to enrich a
# JOB's known numbers, never to extract from the PO side.
_BARE_LONGNUM_RE = re.compile(r"\b(\d{5,})\b")


def _job_number_pool(job: "OpenJob") -> set:
    pool = {_clean_token(x) for x in job.job_numbers if x}
    pool |= {_clean_token(m) for m in _BARE_LONGNUM_RE.findall(job.text or "")}
    pool |= {_clean_token(x) for x in extract_job_numbers(job.text or "")}
    return {p for p in pool if p}


def normalize_part(p: str) -> str:
    return re.sub(r"[\s\-_./]", "", str(p or "")).upper()


_DESC_STOP = {
    "the", "and", "for", "with", "each", "per", "inch", "in", "ft", "foot", "feet",
    "ea", "pc", "pcs", "of", "a", "an", "to", "x",
}


def _desc_tokens(desc: str) -> set:
    words = re.findall(r"[a-z0-9]+", _lower(desc))
    return {w for w in words if len(w) >= 3 and w not in _DESC_STOP}


# ---------------------------------------------------------------------------
# PO-document detection (is this an award at all?)
# ---------------------------------------------------------------------------

# --- Strong PO anchors: high-confidence evidence the document IS a customer PO.
_PO_ANCHORS = [
    (r"\bpurchase\s+order\b", 5.0, "purchase-order heading"),
    # v4.1.42: the labeled-PO-number anchor is no longer a loose regex here -- it is
    # resolved functionally via _has_delivered_po_number so classification uses the
    # SAME semantics as extraction (digit required, placeholders like TBD rejected,
    # "customer po number" fields excluded). See score_po_document.
    (None, 4.0, "labeled PO number"),
    (r"\bthis\s+(?:is\s+)?(?:our|your)\s+(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b", 4.0, "explicit PO language"),
]

# --- Moderate PO evidence: order-form fields that support, but do not prove, a PO.
_PO_MODERATE = [
    (r"\bqty\s+ord\b", 2.0, "quantity-ordered column"),
    (r"\border\s+qty\b", 2.0, "order-quantity column"),
    (r"\breleas(?:e|ed)\b", 1.5, "release instruction"),
    (r"\border\s+date\b", 1.5, "order date"),
    (r"\bper\s+quote\b", 1.5, "references a quote"),
]

# --- Weak PO evidence: generic commercial-document fields that appear on POs,
# quotes, invoices, and packing lists alike. Never enough on their own.
_PO_WEAK = [
    (r"\bship\s+to\b", 0.5, "ship-to"),
    (r"\bbill\s+to\b", 0.5, "bill-to"),
    (r"\bdeliver\s+to\b", 0.5, "deliver-to"),
    (r"\bbuyer\b", 0.5, "buyer"),
    (r"\bterms\s+and\s+conditions\b", 0.5, "terms and conditions"),
]

# --- Hard non-PO headings: if the document is titled/headed as one of these, it
# is NOT a customer purchase order regardless of what fields it also carries.
# Anchored near the top of a line so a passing mention in body text is ignored.
# v4.1.39a: a "QUOTE #"/"QUOTATION #"/"PROPOSAL #"/"ESTIMATE #"/"PRICE QUOTE"
# heading is a quotation -- the prior "(?!\s*#)" wrongly EXEMPTED "QUOTE #", the
# most common seller-quote heading, letting a customer-facing quote that also
# carried a "CUSTOMER PO NUMBER" field be scored as a PO.
_NONPO_HEADINGS = [
    (r"(?im)^\s*(?:re:\s*)?quotation\b", "quotation"),
    (r"(?im)^\s*(?:re:\s*)?proposal\b", "proposal"),
    (r"(?im)^\s*(?:re:\s*)?estimate\b", "estimate"),
    (r"(?im)^\s*(?:re:\s*)?price\s+quote\b", "quote"),
    (r"(?im)^\s*(?:re:\s*)?(?:sales\s+)?quote\s*(?:#|no\.?|number|:)", "quote"),
    (r"(?im)^\s*(?:re:\s*)?quote\b", "quote"),
    (r"(?im)^\s*(?:re:\s*)?request\s+for\s+quote\b", "RFQ"),
    (r"(?im)^\s*(?:re:\s*)?invoice\b", "invoice"),
    (r"(?im)^\s*(?:re:\s*)?packing\s+(?:list|slip)\b", "packing list"),
    (r"(?im)^\s*(?:re:\s*)?submittal\b", "submittal"),
]

# --- Role-defining veto signatures: a document whose PRIMARY role is clearly one
# of these (title-position term + its defining fields) is vetoed. This is the
# audit's "primary role, not any-word-anywhere" fix: an invoice needs invoice
# identity (Amount Due / Remit To / Invoice #), not merely the word "invoice"
# buried in a PO's terms ("email invoice to accounts payable").
_ROLE_VETOES = [
    ("invoice", [r"\bamount\s+due\b", r"\bremit\s+to\b", r"\binvoice\s*(?:#|no\.?|number|date)\b"]),
    ("packing list", [r"\bpacking\s+(?:list|slip)\b", r"\bshipped\s+qty\b", r"\bcarton\s+count\b"]),
    ("submittal", [r"\bshop\s+drawing\b", r"\bapproval\s+drawing\b", r"\bcut\s*sheet\b", r"\bsubmittal\s+(?:package|number|no)\b"]),
]

# v1.4.0 (audit findings: role/state documents masquerading as POs). A document
# TITLED as one of these is not a new customer purchase order, even when it
# carries a PO number and a matching quote reference: an ORDER ACKNOWLEDGEMENT /
# ORDER CONFIRMATION / SALES ORDER is the seller's document about an order, a
# PURCHASE REQUISITION is an internal pre-PO, and PURCHASE ORDER REQUEST /
# STATUS / INQUIRY and PRELIMINARY / PROPOSED PURCHASE ORDER are states of a PO
# that does not (yet) award anything. Applied POSITIONALLY in score_po_document:
# the role/state title vetoes unless a genuine PURCHASE ORDER title/number
# heading appears EARLIER in the heading zone (a title outranks a later field;
# an acknowledgment's own "Purchase Order No: ..." reference line cannot rescue
# it, while a real PO whose later terms mention "order acknowledgement required
# within 24 hours" is unaffected).
_NONPO_ROLE_TITLE_HEADINGS = [
    # v1.5.0: "(?:purchase\s+)?" -- "PURCHASE ORDER ACKNOWLEDGEMENT" previously
    # slipped past both the ack pattern (which required the line to START with
    # "order") and the PO-title regex, and classified as a confident PO.
    (r"(?im)^\s*(?:re:\s*)?(?:purchase\s+)?order\s+acknowledge?ment\b", "order acknowledgment"),
    (r"(?im)^\s*(?:re:\s*)?p\.?\s*[/]?\s*o\.?\s+acknowledge?ment\b", "order acknowledgment"),
    (r"(?im)^\s*(?:re:\s*)?(?:order|sales\s+order)\s+confirmation\b", "order confirmation"),
    (r"(?im)^\s*(?:re:\s*)?sales\s+order\b", "sales order"),
    (r"(?im)^\s*(?:re:\s*)?purchase\s+requisition\b", "purchase requisition"),
    (r"(?im)^\s*(?:re:\s*)?(?:preliminary|proposed|draft)\s+purchase\s+order\b", "preliminary/draft PO"),
    (r"(?im)^\s*(?:re:\s*)?purchase\s+order\s+(?:request|requisition|status|inquiry|change\s+request|"
     r"cancellation|acknowledge?ment|rejection|revocation|withdrawal|receipt)\b",
     "PO request/status/notice"),
    (r"(?im)^\s*(?:re:\s*)?p\.?\s*[/]?\s*o\.?\s+(?:status|cancellation|rejection|revocation|"
     r"withdrawal|request|requisition|inquiry|acknowledge?ment)"
     r"(?:\s+(?:notice|request|update|report|confirmation))?\b", "PO request/status/notice"),
    (r"(?im)^\s*(?:re:\s*)?(?:p\.?\s*[/]?\s*o\.?|purchase\s+order)\s+receipt\s+confirmation\b",
     "PO receipt confirmation"),
]
# v1.8.0 (eighth audit, finding #1): TWO-TIER TITLES. The spelled-out words are a
# FULL title: they outrank later role headings positionally and bypass the non-PO
# heading vetoes, exactly as before. An abbreviated "PO #12345" line is only a
# REFERENCE-GRADE title: real invoices, packing slips, acknowledgments, sales
# orders and quotations all carry a PO-number field, and OCR frequently emits that
# field before the form's actual title -- so an abbreviated title is DEMOTED
# whenever the heading zone contains an explicit non-PO document-title line, and
# it never overrides those vetoes. A "pure title line" is the role phrase with at
# most a cue and one identifier after it ("ORDER ACKNOWLEDGEMENT", "INVOICE #
# 998") -- instruction text ("ORDER ACKNOWLEDGEMENT REQUIRED WITHIN 24 HOURS")
# does not demote, preserving the v1.7.0 abbreviated-PO fix.
_FULL_PO_TITLE_RE = re.compile(
    r"(?im)^\s*(?:re:\s*)?"
    r"purchase\s+order\s*(?:#|no\.?|number|[:\-]|$|\s+(?=[A-Za-z0-9\-/]*\d))")
_ABBREV_PO_TITLE_RE = re.compile(
    r"(?im)^\s*(?:re:\s*)?"
    r"p\.?\s*[/]?\s*o\.?\s*(?:(?:#|no\.?|number|:)[^\S\n]*(?=[A-Za-z0-9\-/]*\d)|\s+(?=[A-Za-z0-9\-/]*\d))")
_PURE_TITLE_TAIL_RE = re.compile(r"^\s*(?:[#:\-]|no\.?|number)?\s*[A-Za-z0-9\-/#]*\s*$", re.I)


def _is_pure_title_line(zone: str, m) -> bool:
    """True when the matched heading phrase IS the line's title -- the remainder of
    its line is at most a cue and one identifier token."""
    try:
        le = zone.find("\n", m.end())
        if le < 0:
            le = len(zone)
        return bool(_PURE_TITLE_TAIL_RE.match(zone[m.end():le]))
    except Exception:
        return False

# --- Non-award PO language, v4.1.42: split into two tiers.
#
# HARD lifecycle vetoes: the document explicitly says it is not (or is no longer)
# an actionable order -- a disclaimer, draft, sample, pro forma, cancellation, or
# do-not-process instruction. These veto REGARDLESS of headings or PO numbers:
# "PURCHASE ORDER ... This is not a purchase order." is a common disclaimer on
# acknowledgments and quotes, and "Cancelled PO #12345" must never create an award.
# v1.4.0 (fourth external audit): the hard list is split into two tiers.
#
# OBJECT-BOUND tier: the phrase itself names the purchase order it kills
# ("draft PO", "PO #12345 was cancelled", "this is not a purchase order"), so it
# can safely veto anywhere in the document. Auxiliary-verb coverage extended --
# "PO #12345 WAS cancelled" and "PO #12345 HAS BEEN voided" previously fell
# through the (is|has been) group and matched confidently. Superseded added.
_HARD_NOT_PO_LANGUAGE = [
    r"\bnot\s+an?\s+(?:official\s+)?purchase\s+order\b",
    r"\bthis\s+is\s+not\s+an?\s+(?:official\s+)?(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\bdraft\s+(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\bsample\s+(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\bpro\s*[- ]?forma\s+(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\bcancell?ed\s+(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s*#?\s*[A-Za-z0-9\-]*\s+"
    r"(?:is\s+|was\s+|are\s+|were\s+|has\s+been\s+|have\s+been\s+)?cancell?ed\b",
    r"\bvoid(?:ed)?\s+(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s*#?\s*[A-Za-z0-9\-]*\s+"
    r"(?:is\s+|was\s+|has\s+been\s+|have\s+been\s+)?void(?:ed)?\b",
    r"\bsuperseded\s+(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s*#?\s*[A-Za-z0-9\-]*\s+"
    r"(?:is\s+|was\s+|has\s+been\s+|have\s+been\s+)?superseded\b",
    # v1.5.0: terminated-lifecycle verbs and the cancellation-notice noun form --
    # "PO #12345 has been rejected/revoked/withdrawn", "PO CANCELLATION NOTICE".
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s*#?\s*[A-Za-z0-9\-]*\s+"
    r"(?:is\s+|was\s+|has\s+been\s+|have\s+been\s+)?"
    r"(?:rejected|revoked|withdrawn|rescinded|declined)\b",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s+cancellation\b",
    # v1.6.0 (sixth audit): terminated / no-longer-current states. "PO #12345 was
    # denied", "has expired", "is closed", "was deleted" all describe an order that
    # cannot be a live award.
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s*#?\s*[A-Za-z0-9\-]*\s+"
    r"(?:is\s+|was\s+|are\s+|were\s+|has\s+(?:been\s+)?|have\s+(?:been\s+)?)?"
    r"(?:denied|expired|closed|deleted|terminated|nullified)\b",
    r"\b(?:expired|closed|denied|deleted)\s+(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
]

# FREE-FLOATING tier: legitimate PO terms boilerplate uses these very phrases
# about OTHER objects -- "do not process INVOICES without this PO number",
# "SUBSTITUTIONS are not approved without written consent", "PRO FORMA INVOICE
# required before shipment", "FOR REVIEW ONLY DRAWINGS must be submitted". A
# whole-document search rejected those valid POs (the audit's false-negative
# class -- critical because a missed PO is negative-checkpointed by the
# attachment scanner). These now veto only when OBJECT-SCOPED: the line/sentence
# containing the phrase must reference the PO/order/document itself and must not
# name an alternate object, OR the phrase stands alone as a short stamp line
# ("FOR REVIEW ONLY", "NOT APPROVED") with no alternate object.
_FREE_FLOATING_NOT_PO_LANGUAGE = [
    r"\bpro\s*[- ]?forma\b",
    r"\bfor\s+review\s+only\b",
    r"\bfor\s+reference\s+only\b",
    r"\bdo\s+not\s+process\b",
    r"\bnot\s+approved\b",
]
_VETO_PO_OBJECT_RE = re.compile(
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?\b|this\s+order|this\s+document)", re.I)
_VETO_ALT_OBJECT_RE = re.compile(
    r"\b(?:invoices?|invoicing|drawings?|submittals?|substitutions?|samples?|"
    r"shipments?|payments?|deposits?|quotes?|quotations?|returns?|credits?)\b", re.I)


def _free_floating_veto_hits(low: str) -> bool:
    """v1.4.0: object-scoped application of the free-floating tier (see above).
    The scope window is the containing LINE (PO forms are line-structured),
    capped at 300 chars."""
    for p in _FREE_FLOATING_NOT_PO_LANGUAGE:
        for m in re.finditer(p, low):
            ls = low.rfind("\n", 0, m.start()) + 1
            le = low.find("\n", m.end())
            if le < 0:
                le = len(low)
            window = low[max(ls, m.start() - 300):min(le, m.end() + 300)]
            if _VETO_ALT_OBJECT_RE.search(window):
                continue                     # about an invoice/drawing/etc., not the PO
            if _VETO_PO_OBJECT_RE.search(window):
                return True                  # names the PO itself
            if len(window.strip()) <= 40:
                return True                  # standalone stamp line ("FOR REVIEW ONLY")
    return False

# SOFT non-award language: the sender is discussing a FUTURE or requested PO --
# negated issuance, "will be issued", "to follow", a request that someone else
# provide one. v4.1.42: these suppress classification EVEN WHEN a labeled PO number
# is present ("PO #12345 will be issued next week" is not an award), unless the
# document carries a genuine PURCHASE ORDER title heading (a real PO whose terms
# mention future paperwork is unaffected).
_NON_AWARD_PO_LANGUAGE = [
    r"\bpurchase\s+order\s*#?\s*[A-Za-z0-9\-]*\s+has\s+not\s+been\s+issued\b",
    r"\bno\s+purchase\s+order\s+(?:has\s+been\s+)?issued\b",
    r"\bp\.?\s*[/]?\s*o\.?\s*#?\s*[A-Za-z0-9\-]*\s+has\s+not\s+been\s+issued\b",
    r"\bp\.?\s*[/]?\s*o\.?\s+not\s+(?:yet\s+)?issued\b",
    r"\bawait(?:ing)?\s+(?:a\s+|the\s+|your\s+)?purchase\s+order\b",
    r"\bawait(?:ing)?\s+(?:a\s+|the\s+|your\s+)?p\.?\s*[/]?\s*o\.?\b",
    r"\bplease\s+(?:provide|send|issue|forward)\s+(?:a\s+|your\s+|the\s+)?(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s*#?\s*[A-Za-z0-9\-]*\s+to\s+follow\b",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s+(?:is\s+)?required\s+before\b",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s*#?\s*[A-Za-z0-9\-]*\s+will\s+be\s+(?:issued|sent|provided|forthcoming)\b",
    r"\bonce\s+(?:the\s+|a\s+|your\s+)?(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s+is\b",
    r"\bwhen\s+(?:the\s+|a\s+|your\s+)?(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s+is\b",
    # v1.4.0 (audit): object-bound in-limbo states lived here (soft tier, bypassed
    # by a PO title).
    # v1.7.0 (seventh audit, finding #4): MOVED to _PO_STATE_VETO_LANGUAGE below,
    # which applies EVEN WHEN the document carries a full PURCHASE ORDER title --
    # "PURCHASE ORDER 12345 / STATUS: PENDING SIGNATURE" was matching confidently
    # because this soft tier is title-bypassed. Only generic transactional wording
    # ("upon receipt of PO", "once the PO is issued") remains soft.
    # v1.5.0 (audit): follow-up chatter ABOUT a PO is not a new award -- "Can you
    # provide status on PO #12345?", "Please confirm receipt of PO #12345",
    # "Attached is a copy of PO #12345 for your records", "PO acknowledgement".
    # Soft: a document carrying a genuine PURCHASE ORDER title (an attached real
    # PO) bypasses, so transmittal wording never suppresses the PO itself.
    r"\b(?:status|update)\s+(?:on|of|for|regarding)\s+(?:(?:the|our|your|this)\s+)?(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\bconfirm(?:ing)?\s+(?:the\s+)?receipt\s+of\s+(?:(?:the|our|your|this)\s+)?(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\bcopy\s+of\s+(?:(?:the|our|your|this)\s+)?(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\s+acknowledge?ment\b",
    r"\backnowledg(?:e|ing|ement|ment)\s+(?:of\s+)?(?:(?:the|our|your|this)\s+)?(?:purchase\s+order|p\.?\s*[/]?\s*o\.?)\b",
]

# v1.7.0 (seventh audit, finding #4): explicit statements that THIS purchase order
# is not yet an authorization. Applied REGARDLESS of a PO title -- a document may
# be titled PURCHASE ORDER and still say it is not released. Three shapes:
#   (a) object-bound sentences anywhere in the document ("PO #12345 pending
#       signature", "THIS PO IS NOT AUTHORIZED");
#   (b) a STATUS field line in the heading zone ("STATUS: PENDING SIGNATURE",
#       "STATUS: AWAITING FUNDING", "STATUS: ON HOLD");
#   (c) a standalone stamp line in the heading zone ("NOT RELEASED FOR PURCHASE",
#       "PENDING EXECUTION").
# Generic terms about OTHER objects ("submittals pending approval", "payment
# pending approval of invoice") match none of these shapes, so real PO terms
# never trip this.
_PO_STATE_VETO_LANGUAGE = [
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?|this\s+(?:po|order))\s*#?\s*[A-Za-z0-9\-]*\s+(?:is\s+|was\s+)?"
    r"(?:pending\s+(?:approval|release|review|authori[sz]ation)|on\s+hold|"
    r"awaiting\s+(?:approval|release|signature|authori[sz]ation)|under\s+review)\b",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?|this\s+(?:po|order))\s*#?\s*[A-Za-z0-9\-]*\s+(?:is\s+|was\s+)?"
    r"(?:pending|awaiting)\s+(?:signature|signatures|funding|funds|budget|"
    r"release|execution|countersignature)\b",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?|this\s+(?:po|order))\s*#?\s*[A-Za-z0-9\-]*\s+"
    r"(?:is\s+|was\s+|are\s+|were\s+|has\s+|have\s+)?"
    r"not\s+(?:yet\s+)?(?:been\s+)?(?:funded|released|signed|executed|authori[sz]ed|approved|issued|placed)\b",
    # v1.8.2 (ninth audit): a FUTURE PO is not an award, title or no title --
    # "PO #12345 will be issued next week", "PO #12345 to follow", "PO
    # forthcoming". Object-bound (the phrase must directly describe the numbered
    # PO), so ordinary terms about other objects never trip it.
    # v1.8.3 (tenth audit): NARROWED. "will be sent/provided/released" without
    # qualification is ordinary OPERATIONAL language on a real PO -- "this order
    # will be released for shipment / to production", "will be sent via UPS
    # Ground", "will be provided in two releases" -- and vetoing it created
    # missed-PO false negatives (which get negative-checkpointed). "will be
    # issued" / "to follow" / "forthcoming" remain strong future-PO evidence on
    # their own; sent/provided/released now veto only with explicit AUTHORIZATION
    # context ("will be released after approval", "will be sent once signed").
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?|this\s+(?:po|order))\s*#?\s*[A-Za-z0-9\-]*\s+"
    r"(?:will\s+be\s+issued\b|to\s+follow\b|(?:is\s+)?forthcoming\b)",
    r"\b(?:purchase\s+order|p\.?\s*[/]?\s*o\.?|this\s+(?:po|order))\s*#?\s*[A-Za-z0-9\-]*\s+"
    r"will\s+be\s+(?:sent|provided|released)\s+(?:after|upon|once|when|following|pending)\s+"
    r"(?:\w+\s+){0,2}?(?:approv(?:al|ed)|authori[sz](?:ation|ed)|sign(?:ature|ed|-?off)|funding|funded|release)\b",
]
_PO_STATE_STATUS_FIELD_RE = re.compile(
    r"(?im)^\s*(?:p\.?\s*[/]?\s*o\.?\s+)?status\s*[:=\-]\s*(?:pending|awaiting|on\s+hold|hold\b|draft|unapproved|"
    r"not\s+(?:released|authori[sz]ed|signed|funded|approved|executed)|"
    # v1.8.0 (eighth audit, finding #2): TERMINAL states -- "STATUS: CANCELLED" on a
    # document titled PURCHASE ORDER is a cancellation, not an award.
    r"cancell?ed|void(?:ed)?\b|rejected|revoked|withdrawn|closed|expired|"
    r"deleted|denied|terminated|nullified|superseded)")
_PO_STATE_STAMP_RE = re.compile(
    r"(?im)^\s*(?:not\s+released(?:\s+for\s+purchase)?|not\s+(?:yet\s+)?authori[sz]ed|"
    r"pending\s+(?:signature|execution|funding|release|approval|authori[sz]ation)|"
    r"awaiting\s+(?:signature|execution|funding|release|approval|authori[sz]ation)|"
    # v1.8.0: terminal standalone stamps.
    r"cancell?ed|void(?:ed)?|rejected|revoked|withdrawn|superseded|terminated|nullified)\s*[.!]?\s*$")


def _heading_zone(context: "POContext") -> str:
    """Where a document's title/heading lives. v4.1.42: OCR/logo noise routinely
    pushes the real heading past a fixed character offset, so this is now the
    subject plus the first 15 NON-EMPTY lines of the attachment (capped at 1500
    chars) rather than a blind first-400-chars slice. Role and heading tests look
    here; anchors are searched document-wide."""
    lines = []
    for ln in (context.attachment_text or "").splitlines():
        if ln.strip():
            lines.append(ln)
            if len(lines) >= 15:
                break
    top = "\n".join([context.subject or ""] + lines)
    return top[:1500]


# Placeholder "values" that appear after a PO-number label when no real number
# exists yet. These must never establish that a document IS a PO (v4.1.42 --
# "PO NUMBER: TBD" on a quote was scoring as a labeled-PO anchor while the
# extractor correctly rejected it; classification and extraction now agree).
_PO_PLACEHOLDER_VALUES = frozenset({
    "TBD", "TBA", "NA", "NONE", "PENDING", "TOFOLLOW", "FOLLOW", "UNKNOWN",
    "XX", "XXX", "XXXX", "TB", "REQUIRED", "VERBAL", "OPEN",
})


def _has_delivered_po_number(low: str) -> bool:
    """True when the text carries a labeled PO number the sender is PROVIDING
    (e.g. 'PO Number: 12345'), as opposed to merely mentioning purchase orders.
    v4.1.42: the value must CONTAIN A DIGIT and must not be a placeholder
    (TBD/PENDING/N/A/...), matching extract_po_numbers' semantics exactly. A
    'customer po number' label inside a seller's own quote does NOT count."""
    for m in re.finditer(
            r"\bp\.?\s*[/]?\s*o\.?\s*(?:#|no\.?|number)\s*[:#]?\s*"
            r"((?:[A-Za-z]{1,3}\s(?=\d))?[A-Za-z0-9][A-Za-z0-9\-/.]*)", low):
        # Reject when this PO-number field is a "customer po number" line on our quote.
        start = max(0, m.start() - 20)
        if re.search(r"\bcustomer\s+p\.?\s*[/]?\s*o\.?\b", low[start:m.end()]):
            continue
        tok = _clean_token(m.group(1))
        if not tok or not any(c.isdigit() for c in tok):
            continue                      # "TBD", "pending", empty -> not a number
        if tok in _PO_PLACEHOLDER_VALUES:
            continue
        return True
    return False


def score_po_document(context: "POContext") -> Tuple[float, List[str], List[str]]:
    """Return (po_score, cue_hits, veto_reasons).

    Weighted model. A PO needs a strong ANCHOR (a purchase-order title heading, a
    DELIVERED labeled PO number -- digit-bearing, non-placeholder, non-"customer po
    number" -- or explicit this-is-our-PO language); generic ship-to/bill-to/buyer/
    terms fields are only weak support. Quote/proposal/estimate/invoice/RFQ HEADINGS
    are hard non-PO; the quote-family (never RFQ or invoice) can be overridden by a
    delivered PO number in the heading zone. HARD lifecycle language (draft, sample,
    cancelled, "this is not a purchase order", do-not-process) vetoes regardless of
    headings and numbers; SOFT future/request language suppresses classification even
    when a number is present, unless a genuine PO title heading exists.
    """
    combined = "\n".join([
        context.subject or "", context.body or "", context.attachment_text or "",
    ])
    low = combined.lower()
    heading = _heading_zone(context).lower()

    cues: List[str] = []
    vetoes: List[str] = []

    # A genuine PURCHASE ORDER heading overrides quote/role vetoes (a real PO may
    # reference a quote); a "customer po number" FIELD does not count as that.
    # Require the heading to be a TITLE -- not a sentence that merely starts with the
    # words ("Purchase order has not been issued yet." is not a title).
    _hz0 = _heading_zone(context)
    _full_title = bool(_FULL_PO_TITLE_RE.search(_hz0))
    _abbrev_title = bool(_ABBREV_PO_TITLE_RE.search(_hz0))
    if _full_title or _abbrev_title:
        # v1.8.2 (ninth audit): the scan must inspect the first line that actually
        # MATCHES a title regex. The previous loose "starts with PO" test broke on
        # a subject like "PO attached" (not a title) before ever reaching the
        # attachment line "PO #12345 will be issued next week" -- so the negation
        # ("will be") was never seen and the future-PO document matched at 100%.
        for _line in _hz0.splitlines():
            if not (_FULL_PO_TITLE_RE.search(_line) or _ABBREV_PO_TITLE_RE.search(_line)):
                continue
            if re.search(r"(?i)has\s+not|not\s+(?:yet\s+)?been|to\s+follow|not\s+issued|"
                         r"not\s+placed|required\s+before|will\s+be|forthcoming", _line):
                _full_title = _abbrev_title = False
            break
    # v1.8.0 (eighth audit, finding #1): a reference-grade abbreviated title is
    # demoted by any explicit non-PO document-title line in the heading zone --
    # "PO #12345 / INVOICE" is an invoice with a PO field, not a purchase order.
    if _abbrev_title and not _full_title:
        for _dt_pat, _dt_label in list(_NONPO_HEADINGS) + list(_NONPO_ROLE_TITLE_HEADINGS):
            _dt_m = re.search(_dt_pat, _hz0)
            if _dt_m and _is_pure_title_line(_hz0, _dt_m):
                _abbrev_title = False
                break
    _has_po_heading = _full_title or _abbrev_title

    # HARD lifecycle vetoes. v1.4.0: two tiers (see list definitions). The
    # OBJECT-BOUND tier ("draft PO", "PO #12345 was cancelled", "this is not a
    # purchase order") vetoes anywhere in the document -- a heading or a PO number
    # does not rescue it. The FREE-FLOATING tier (do-not-process / not-approved /
    # pro-forma / for-review-only) is scoped to the line that contains it and only
    # vetoes when that line names the PO itself (or is a standalone stamp), so
    # boilerplate about invoices/drawings/substitutions no longer suppresses a
    # valid PO into a permanent negative checkpoint.
    for p in _HARD_NOT_PO_LANGUAGE:
        if re.search(p, low):
            vetoes.append("non-actionable PO (draft/cancelled/disclaimer)")
            break
    else:
        if _free_floating_veto_hits(low):
            vetoes.append("non-actionable PO (draft/cancelled/disclaimer)")

    # v1.7.0 (seventh audit, finding #4): explicit not-yet-authorized states veto
    # REGARDLESS of a PURCHASE ORDER title. The soft tier below is title-bypassed
    # by design (real PO terms mention approvals), but a sentence about THIS PO, a
    # STATUS field, or a standalone stamp describes the document itself -- the
    # title must not override it. "PURCHASE ORDER 12345 / STATUS: PENDING
    # SIGNATURE" is not an active award.
    if not vetoes:
        _hz_state = _heading_zone(context)
        _state_hit = any(re.search(p, low) for p in _PO_STATE_VETO_LANGUAGE)
        if not _state_hit:
            _state_hit = bool(_PO_STATE_STATUS_FIELD_RE.search(_hz_state)
                              or _PO_STATE_STAMP_RE.search(_hz_state))
        if _state_hit:
            vetoes.append("PO not yet authorized (pending signature/funding/release/approval)")

    # v1.4.0: role/state TITLE headings (order acknowledgment, sales order,
    # purchase requisition, PO request/status, preliminary PO). Positional rule:
    # they veto unless a genuine PURCHASE ORDER title/number heading appears
    # EARLIER in the heading zone.
    _hz_raw = _heading_zone(context)
    _po_title_m = _FULL_PO_TITLE_RE.search(_hz_raw)   # v1.8.0: FULL titles only
    for pat, label in _NONPO_ROLE_TITLE_HEADINGS:
        m = re.search(pat, _hz_raw)
        if not m:
            continue
        if _po_title_m and _po_title_m.start() < m.start():
            continue
        # v1.8.0: on an abbreviated-title document, only a PURE role-title line
        # vetoes -- "ORDER ACKNOWLEDGEMENT" (a title) does, "ORDER ACKNOWLEDGEMENT
        # REQUIRED WITHIN 24 HOURS" (terms text) does not.
        if _abbrev_title and not _is_pure_title_line(_hz_raw, m):
            continue
        vetoes.append(label)

    # Hard non-PO heading: the document announces itself as something else. Real PO
    # forms often carry a "QUOTE NO: S…" REFERENCE line while the actual title is a
    # logo image, so the quote-family veto (quote/quotation/proposal/estimate) is
    # overridden by a delivered, non-customer PO number in the HEADING ZONE. v4.1.42:
    # RFQ is deliberately NOT in the override family -- a "REQUEST FOR QUOTE" heading
    # is always an RFQ even when the form carries a PO field -- and neither are
    # invoice/packing/submittal, which legitimately reference the buyer's PO number.
    _quote_family = {"quote", "quotation", "proposal", "estimate"}
    _heading_has_po_number = _has_delivered_po_number(heading)
    for pat, label in _NONPO_HEADINGS:
        _nh_m = re.search(pat, _heading_zone(context))
        if _nh_m:
            # v1.8.0 (eighth audit, finding #1): only a FULL "PURCHASE ORDER" title
            # bypasses these vetoes; a reference-grade "PO #12345" line does not --
            # an invoice/packing slip/quotation with a PO field is not a PO. The
            # quote-family PO-number rescue likewise requires a full title (real
            # POs whose "QUOTE NO: S..." reference trips the quote pattern are
            # titled PURCHASE ORDER; a quote-titled document with a PO field is a
            # quote). On abbreviated-title documents only PURE title lines veto.
            if _full_title:
                continue
            if label in _quote_family and _heading_has_po_number and _full_title:
                continue
            if _abbrev_title and not _is_pure_title_line(_heading_zone(context), _nh_m):
                continue
            vetoes.append(label)

    # Role-defining veto: primary-role identity present in the heading zone.
    for label, sigs in _ROLE_VETOES:
        hits = sum(1 for p in sigs if re.search(p, heading))
        # v1.8.0: identity fields (Amount Due / Remit To / Invoice #) beat a
        # reference-grade abbreviated title; only a FULL title bypasses.
        if hits >= 2 and not _full_title:
            vetoes.append(label)

    # SOFT non-award language (v4.1.42): future/requested-PO wording suppresses BOTH
    # the bare phrase anchor AND the labeled-number anchor unless a genuine PO title
    # heading is present. "PO #12345 will be issued next week" carries a number but
    # is not an award.
    _non_award = any(re.search(p, low) for p in _NON_AWARD_PO_LANGUAGE)
    _suppress_anchors = _non_award and not _has_po_heading

    _delivered_number = _has_delivered_po_number(low)

    # Weighted PO evidence.
    score = 0.0
    for pat, w, name in _PO_ANCHORS:
        if name == "labeled PO number":
            # Functional anchor: same semantics as extraction (v4.1.42).
            if _delivered_number and not _suppress_anchors:
                score += w
                cues.append(name)
            continue
        if name == "purchase-order heading":
            # v1.7.0 (seventh audit, finding #3): an ABBREVIATED authoritative
            # title ("PO 12345", "P.O. #12345", "P/O #12345" at title position,
            # negation-scanned) anchors like the spelled-out words. The words
            # regex alone missed them, so a valid abbreviated-title PO scored
            # below threshold whenever the number lacked a #/no./number cue.
            if re.search(pat, low) or _has_po_heading:
                if not _suppress_anchors:
                    score += w
                    cues.append(name)
            continue
        if re.search(pat, low):
            score += w
            cues.append(name)
    for pat, w, name in _PO_MODERATE:
        if re.search(pat, low):
            score += w
            cues.append(name)
    for pat, w, name in _PO_WEAK:
        if re.search(pat, low):
            score += w
            cues.append(name)

    if _non_award and _suppress_anchors:
        vetoes.append("non-award PO language (negated/future/requested)")

    return score, cues, sorted(set(vetoes))


# ---------------------------------------------------------------------------
# Core
# ---------------------------------------------------------------------------

class POMatchEngine:
    def __init__(self, policy: Optional[POPolicy] = None):
        self.policy = policy or POPolicy()

    # -- field extraction from the PO side -------------------------------
    def _po_signals(self, context: "POContext") -> Dict[str, List[str]]:
        text = "\n".join([
            context.subject or "", context.body or "", context.attachment_text or "",
        ])
        quotes = list(context.stated_quote_numbers) or extract_quote_numbers(text)
        pos = list(context.stated_po_numbers) or extract_po_numbers(text)
        jobs = list(context.stated_job_numbers) or extract_job_numbers(text)
        quotes = [_clean_token(x) for x in quotes if x]
        pos = [_clean_token(x) for x in pos if x]
        jobs = [_clean_token(x) for x in jobs if x]
        parts = [normalize_part(x) for x in context.line_parts if x]
        descs = [str(x) for x in context.line_descriptions if x]
        if not parts and not descs:
            # No pre-parsed lines: pull candidate item codes from the attachment
            # text. Catalog numbers vary widely -- hyphenated (HBL-5362), short
            # alphanumeric (632S), or long numeric -- so accept:
            #   * a hyphen/dot-joined alnum run with a digit (HBL-5362, 40A-SW-82)
            #   * a 4+ char alnum token mixing letters and digits (632S, ABC123)
            #   * a 6+ digit pure-number token
            # then drop any token that is actually a quote/PO/job number we already
            # extracted, so identifiers never masquerade as line items.
            _ident = set(quotes) | set(pos) | set(jobs)
            _raw = context.attachment_text or ""
            _cands = re.findall(r"\b([A-Za-z0-9]+(?:[-.][A-Za-z0-9]+)+)\b", _raw)  # hyphenated
            _cands += re.findall(r"\b([A-Za-z0-9]{4,})\b", _raw)                    # solid tokens
            for m in _cands:
                tok = normalize_part(m)
                if not tok or tok in _ident:
                    continue
                has_d = any(c.isdigit() for c in tok)
                has_a = any(c.isalpha() for c in tok)
                if has_d and has_a and len(tok) >= 4:
                    parts.append(tok)
                elif tok.isdigit() and len(tok) >= 6:
                    parts.append(tok)
        return {
            "quote_numbers": _dedup(quotes),
            "po_numbers": _dedup(pos),
            "job_numbers": _dedup(jobs),
            "parts": _dedup(parts),
            "descriptions": descs,
        }

    # -- score one job against the PO signals ----------------------------
    def _score_job(self, sig: Mapping[str, List[str]], job: "OpenJob",
                   sender_email: str = "") -> Tuple[float, List[str], str]:
        reasons: List[str] = []
        score = 0.0

        po_quotes = set(sig["quote_numbers"])
        job_quotes = {_clean_token(x) for x in job.quote_numbers if x}
        job_quotes |= set(extract_quote_numbers(job.text or ""))
        q_hit = po_quotes & job_quotes
        if q_hit:
            score += self.policy.weight_quote_number
            reasons.append(f"quote number {sorted(q_hit)[0]} matches this job")

        po_nums = set(sig["po_numbers"])
        job_pos = {_clean_token(x) for x in job.po_numbers if x}
        job_pos |= set(extract_po_numbers(job.text or ""))
        p_hit = po_nums & job_pos
        if p_hit:
            score += self.policy.weight_po_number
            reasons.append(f"PO number {sorted(p_hit)[0]} matches this job")

        po_jobs = set(sig["job_numbers"])
        job_jobs = _job_number_pool(job)
        j_hit = po_jobs & job_jobs
        if j_hit:
            score += self.policy.weight_job_number
            reasons.append(f"job number {sorted(j_hit)[0]} matches this job")

        # Material overlap: fraction of PO lines found among the job's requested
        # items (parts first, then description-token overlap).
        po_parts = set(sig["parts"])
        job_parts = {normalize_part(x) for x in job.requested_parts if x}
        part_hits = po_parts & job_parts
        mat_frac = 0.0
        if po_parts:
            mat_frac = len(part_hits) / len(po_parts)
        # Description overlap as a softer signal when parts don't line up.
        if mat_frac < 1.0 and sig["descriptions"] and job.requested_descriptions:
            job_desc_tokens = [_desc_tokens(d) for d in job.requested_descriptions]
            desc_matched = 0
            for d in sig["descriptions"]:
                dt = _desc_tokens(d)
                if not dt:
                    continue
                if any(len(dt & jt) >= max(1, min(len(dt), len(jt)) // 2) for jt in job_desc_tokens):
                    desc_matched += 1
            if sig["descriptions"]:
                mat_frac = max(mat_frac, desc_matched / len(sig["descriptions"]))
        if mat_frac > 0:
            # v4.1.42 strong-identifier gate: material evidence alone must never
            # auto-apply on thin grounds. When NO strong identifier (quote/PO/job
            # number) links this job:
            #   * >= 2 exact part matches  -> full material weight (distinctive)
            #   * exactly 1 exact part     -> reduced weight (single part + customer
            #                                 identity stays below auto-match)
            #   * description overlap only -> half weight, capped (one generic word
            #                                 like "connector" can never auto-match)
            # With a strong identifier present, material keeps full weight as
            # corroboration.
            _strong_id = bool(q_hit or p_hit or j_hit)
            _mat_w = self.policy.weight_material_full
            if not _strong_id:
                if len(part_hits) >= 2:
                    pass                                  # distinctive: full weight
                elif len(part_hits) == 1:
                    _mat_w = self.policy.weight_material_single_part
                else:
                    _mat_w = self.policy.weight_material_desc_only
            score += _mat_w * mat_frac
            if part_hits:
                reasons.append(f"{len(part_hits)} PO line item(s) match the job's quoted parts")
            else:
                reasons.append("PO line descriptions overlap the job's quoted items")
            if not _strong_id and len(part_hits) < 2:
                reasons.append("material evidence is limited -- no quote/PO/job number links this job")

        # Customer identity: ACTUALLY compare the PO sender to the job's known
        # customer address (v4.1.39 -- the prior code claimed a match whenever the
        # job merely HAD a customer email, without comparing). Agreement is mild
        # corroboration; a confirmed conflict is surfaced to the caller which forces
        # Review rather than silently claiming a match.
        relation = _email_identity_relation(sender_email, job.customer_email)
        if relation == "exact":
            score += self.policy.weight_customer_exact
            reasons.append("customer email matches the job")
        elif relation == "domain":
            score += self.policy.weight_customer_domain
            reasons.append("customer company domain matches the job")
        # "conflict" and "unknown" add no score and no reason here; conflict is
        # handled in _evaluate so it can temper the FINAL winning decision.

        # v4.1.39a: mild penalty for a job that already has a PO, so a later PO
        # (change order / partial release) can still match it but it never outranks a
        # fresh open job on equal evidence. Only applied when there is some positive
        # evidence to reduce (a bare prior-PO job with no signal scores 0 either way).
        if getattr(job, "prior_po", False) and score > 0:
            score = max(0.0, score - self.policy.penalty_prior_po)

        return score, reasons, relation

    # -- public API ------------------------------------------------------
    def evaluate(self, context: "POContext", open_jobs: Sequence["OpenJob"]) -> POVerdict:
        try:
            return self._evaluate(context, open_jobs)
        except Exception as exc:  # fail-safe: never raise into the caller
            return POVerdict(
                decision=PODecision.NONE,
                reasons=[f"po-match engine error: {type(exc).__name__}: {exc}"[:200]],
            )

    def _evaluate(self, context: "POContext", open_jobs: Sequence["OpenJob"]) -> POVerdict:
        # 1. Is it a PO document at all?
        po_score, _cues, vetoes = score_po_document(context)
        internal = is_internal_sender(context.sender_email, context.internal_domains) \
            or str(context.sender_type or "").upper() == "E"

        sig = self._po_signals(context)
        v = POVerdict(
            decision=PODecision.NONE,
            is_po=False,
            po_numbers=sig["po_numbers"],
            quote_numbers=sig["quote_numbers"],
            job_numbers=sig["job_numbers"],
        )

        if internal and self.policy.reject_internal_senders:
            v.decision = PODecision.NOT_PO
            v.reasons.append("internal sender cannot originate a customer purchase order")
            return v
        if vetoes:
            v.decision = PODecision.NOT_PO
            v.reasons.append("document reads as " + ", ".join(vetoes) + ", not a purchase order")
            return v
        if po_score < self.policy.min_po_document_score:
            v.decision = PODecision.NOT_PO
            v.reasons.append("insufficient purchase-order evidence in the document")
            return v

        v.is_po = True

        # 2. Which job? Score every open job; best-match-wins. Each score carries the
        # customer-identity relation (exact/domain/conflict/unknown) so a confirmed
        # mismatch on the winner can force Review.
        scored: List[Tuple[float, List[str], "OpenJob", str]] = []
        for job in open_jobs or []:
            s, reasons, relation = self._score_job(sig, job, sender_email=context.sender_email)
            if s > 0:
                scored.append((s, reasons, job, relation))
        scored.sort(key=lambda t: t[0], reverse=True)

        if not scored:
            v.decision = PODecision.REVIEW
            v.reasons.append("recognized as a purchase order, but no open job matched its identifiers")
            return v

        best_score, best_reasons, best_job, best_relation = scored[0]
        runner_score = scored[1][0] if len(scored) > 1 else 0.0
        v.runner_up_label = scored[1][2].label if len(scored) > 1 else ""
        v.runner_up_score = runner_score

        # Confidence is the winning score, capped, tempered by the lead over the
        # runner-up so an ambiguous tie never reads as certain. The tempering floor
        # rises with the lead: once the winner is a full evidence-family ahead of the
        # runner-up (lead >= the required margin) it is not dragged below the match
        # bar, but a near-tie (small lead) is pulled well down into review range.
        lead = best_score - runner_score
        confidence = min(1.0, best_score)
        if runner_score > 0:
            margin = max(1e-6, self.policy.min_lead_over_runner_up)
            temper = 0.5 + 0.5 * min(1.0, lead / margin)
            confidence = min(confidence, temper)

        v.job_id = best_job.job_id
        v.job_label = best_job.label
        v.confidence = confidence
        v.reasons = best_reasons

        strong_enough = (
            best_score >= self.policy.min_match_confidence
            and confidence >= self.policy.min_match_confidence
            and (runner_score <= 0 or lead >= self.policy.min_lead_over_runner_up)
        )
        # v4.1.39: a CONFIRMED customer conflict (the PO sender is a different, known
        # company from the matched job's customer) never auto-applies. The quote
        # number may still legitimately point here (a GC forwarding the owner's PO),
        # so we surface it for confirmation rather than discarding the match -- but we
        # do not silently mark a job won for what looks like another customer.
        v.customer_conflict = (best_relation == "conflict")
        if v.customer_conflict:
            strong_enough = False

        # v1.5.0 (audit): EXISTING-PO FOLLOW-UP GATE. A PO-number match is strong
        # evidence of WHICH job, but when every PO number on the document is
        # ALREADY recorded on that job, the document is about an existing award --
        # a status request, receipt confirmation, duplicate copy, acknowledgment,
        # or a revision -- not automatically a new one. Never auto-apply; surface
        # for review with an explicit reason. A genuinely NEW PO number (or a mix
        # containing one) keeps the normal path. Only the job's RECORDED
        # po_numbers count here (not numbers scraped from its free text), so a
        # first real award is never demoted by an incidental mention.
        _ctx_pos = {t for t in (sig["po_numbers"] or []) if t}
        _job_pos = {_clean_token(x) for x in (best_job.po_numbers or ()) if x}
        _followup = bool(_ctx_pos) and bool(_job_pos) and _ctx_pos <= _job_pos
        _revision = False
        if _followup:
            _low_all = "\n".join([context.subject or "", context.body or "",
                                  context.attachment_text or ""]).lower()
            _revision = bool(re.search(
                r"\brev(?:ision|ised)?\b|\bchange\s+order\b|\bamend(?:ed|ment)\b",
                _low_all))
            strong_enough = False

        v.decision = PODecision.MATCH if strong_enough else PODecision.REVIEW
        if not strong_enough:
            if _followup:
                _pn = sorted(_ctx_pos)[0]
                if _revision:
                    v.reasons.append(
                        f"PO {_pn} is already recorded on this job and the document carries "
                        f"revision/change-order language -- review as a PO revision, not a new award")
                else:
                    v.reasons.append(
                        f"PO {_pn} is already recorded on this job -- reads as a follow-up "
                        f"(status/receipt/copy/acknowledgment), not a new award")
            elif v.customer_conflict:
                v.reasons.append(
                    "the PO sender's company differs from this job's customer -- confirm before applying")
            elif runner_score > 0 and lead < self.policy.min_lead_over_runner_up:
                v.reasons.append(
                    f"another open job ({v.runner_up_label}) is a close second -- needs confirmation")
            else:
                v.reasons.append("match evidence is below the automatic-apply threshold -- needs confirmation")
        return v


def _dedup(seq: Sequence[str]) -> List[str]:
    out = []
    seen = set()
    for x in seq:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out


def evaluate_po_document(context: "POContext", open_jobs: Sequence["OpenJob"],
                         policy: Optional[POPolicy] = None) -> Dict[str, Any]:
    """Stateless convenience wrapper returning a plain dict."""
    return POMatchEngine(policy=policy).evaluate(context, open_jobs).to_dict()
