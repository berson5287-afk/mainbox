"""MaINbox Quote Coverage Intake Engine.

A conservative, dependency-free gate between attachment extraction and the
Quote Coverage requested-item ledger.  The engine classifies the whole document
before accepting any rows, requires positive material evidence per row, detects
priced/commercial documents, blocks internal senders, and quarantines ambiguous
results for human review.

Designed for MaINbox v4.1.30+ but usable as a standalone module.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from enum import Enum
import math
import os
import re
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Set, Tuple

ENGINE_VERSION = "1.0.0"


class IntakeDecision(str, Enum):
    ACCEPT = "accept"
    QUARANTINE = "quarantine"
    REJECT = "reject"


class DocumentRole(str, Enum):
    CUSTOMER_RFQ = "customer_rfq"
    BOM = "bom_or_takeoff"
    VENDOR_QUOTE = "vendor_quote"
    INTERNAL_BID = "internal_bid"
    SALES_ORDER = "sales_order"
    PURCHASE_ORDER = "purchase_order"
    INVOICE = "invoice"
    PACKING_LIST = "packing_list"
    SUBMITTAL = "submittal"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class IntakeContext:
    subject: str = ""
    body: str = ""
    sender_email: str = ""
    sender_name: str = ""
    sender_type: str = ""
    internal_domains: Tuple[str, ...] = ()
    source_file: str = ""
    message_id: str = ""
    conversation_id: str = ""


@dataclass
class RowAssessment:
    source_index: int
    accepted: bool
    score: float
    item: Dict[str, Any]
    reasons: List[str] = field(default_factory=list)
    hard_reject_reason: str = ""
    price_evidence: bool = False
    metadata_evidence: bool = False


@dataclass
class IntakeResult:
    decision: IntakeDecision
    document_role: DocumentRole
    confidence: float
    accepted_items: List[Dict[str, Any]] = field(default_factory=list)
    rejected_rows: List[RowAssessment] = field(default_factory=list)
    accepted_rows: List[RowAssessment] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    review_reason: str = ""
    metrics: Dict[str, Any] = field(default_factory=dict)
    audit: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data["decision"] = self.decision.value
        data["document_role"] = self.document_role.value
        return data


@dataclass(frozen=True)
class IntakePolicy:
    min_accept_confidence: float = 0.78
    min_material_ratio: float = 0.60
    max_metadata_ratio: float = 0.15
    max_price_only_ratio: float = 0.12
    min_row_score: float = 5.0
    max_auto_rows: int = 250
    quarantine_on_mixed_commercial: bool = True
    reject_internal_senders: bool = True


_UOM_ALIASES = {
    "ea": "EA", "each": "EA", "pc": "EA", "pcs": "EA", "piece": "EA", "pieces": "EA",
    "ft": "FT", "feet": "FT", "foot": "FT", "lf": "FT",
    "in": "IN", "inch": "IN", "inches": "IN",
    "m": "M", "meter": "M", "meters": "M",
    "c": "C", "hundred": "C",
    "mft": "MFT", "m\u2032": "MFT",
    "roll": "ROLL", "rolls": "ROLL", "box": "BOX", "boxes": "BOX",
    "bag": "BAG", "bags": "BAG", "set": "SET", "sets": "SET",
    "lot": "LOT", "length": "LENGTH", "lengths": "LENGTH",
    "lb": "LB", "lbs": "LB", "spool": "SPOOL", "reel": "REEL", "reels": "REEL",
}

_MATERIAL_WORDS = {
    "adapter", "bender", "box", "breaker", "bushing", "cable", "clamp", "connector",
    "conduit", "coupling", "disconnect", "elbow", "emt", "enclosure", "fitting", "fuse",
    "gland", "ground", "hub", "imc", "junction", "lug", "mc", "panel", "pipe", "plate",
    "pvc", "raceway", "receptacle", "rigid", "splice", "strap", "strut", "switch",
    "terminal", "transformer", "tray", "wire", "wireway", "whip", "nipple", "locknut",
    "contactor", "relay", "starter", "meter", "fixture", "light", "lamp", "device",
    "tape", "rod", "channel", "hanger", "support", "pullbox", "pull box", "cabinet",
}

_REQUEST_PHRASES = re.compile(
    r"\b(?:please\s+quote|quote\s+request|request\s+for\s+quote|rfq|price\s+and\s+availability|"
    r"pricing\s+and\s+availability|material\s+list|parts\s+list|item\s+list|take[- ]?off|bom|"
    r"quote\s+the\s+attached|price\s+the\s+attached|please\s+price|please\s+advise\s+price)\b",
    re.I,
)

_COMMERCIAL_HEADINGS = {
    "customer number", "customer no", "customer #", "customer po number", "customer po",
    "salesperson", "sales person", "writer", "ship via", "ship date", "terms", "payment terms",
    "freight allowed", "freight", "order qty", "order quantity", "unit price", "extended price",
    "extension", "subtotal", "total", "tax", "quote number", "quotation number", "bid number",
    "sales order", "invoice number", "invoice date", "bill to", "ship to", "sold to",
    "valid through", "valid until", "lead time", "availability", "customer account",
}

_METADATA_LABELS = _COMMERCIAL_HEADINGS | {
    "customer", "po number", "po #", "job number", "job name", "project", "project name",
    "requested by", "prepared by", "contact", "phone", "fax", "email", "date", "page",
    "delivery", "delivery date", "carrier", "warehouse", "branch", "location", "notes",
    "attention", "attn", "reference", "our reference", "your reference", "status", "no",
}

_PRICE_RE = re.compile(
    r"(?:\$\s*\d[\d,]*(?:\.\d+)?|\b\d[\d,]*\.\d{2,4}\s*(?:/\s*(?:ea|each|c|m|mft|ft|lb|roll|box))?\b|"
    r"\b\d[\d,]*(?:\.\d+)?\s*/\s*(?:ea|each|c|m|mft|ft|lb|roll|box)\b)", re.I
)
_PRICE_ONLY_RE = re.compile(
    r"^\s*(?:\$?\s*\d[\d,]*(?:\.\d+)?\s*(?:/\s*(?:ea|each|c|m|mft|ft|lb|roll|box))?"
    r"|\$?\s*\d[\d,]*(?:\.\d+)?\s*[-\u2013]\s*\$?\s*\d[\d,]*(?:\.\d+)?)\s*$",
    re.I,
)
_QTY_ONLY_RE = re.compile(r"^\s*\d[\d,]*(?:\.\d+)?\s*(?:ea|each|pc|pcs|ft|feet|m|c|roll|box|bag|set|lot)?\s*$", re.I)
_DATE_RE = re.compile(r"^\s*(?:\d{1,2}[/-]\d{1,2}[/-](?:\d{2}|\d{4})|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{1,2},?\s+\d{4})\s*$", re.I)
_ADDRESS_RE = re.compile(r"\b(?:[A-Z]{2}\s+\d{5}(?:-\d{4})?|street|st\.|avenue|ave\.|road|rd\.|boulevard|blvd\.|suite|floor|city|state|zip)\b", re.I)
_PAYMENT_TERMS_RE = re.compile(r"^\s*(?:net\s*\d+|cod|prepaid|due\s+on\s+receipt|credit\s+card|cash)\s*(?:days?)?\s*$", re.I)
_PERSONISH_RE = re.compile(r"^[A-Z][A-Z'\-]+(?:\s+[A-Z][A-Z'\-]+){1,3}$")
_NUMBER_RANGE_RE = re.compile(r"^\s*[\d,.]+\s*[-\u2013]\s*[\d,.]+\s*$")
_SPEC_RE = re.compile(
    r"(?:\b\d+(?:/\d+)?\s*(?:\"|in\b|inch\b|ft\b|mm\b)|#\s*\d+\b|\b\d+\s*awg\b|"
    r"\b\d+(?:/\d+)?\s*(?:kv|v|amp|a)\b|\b[123]\s*(?:ph|phase|\u03c6|\u00f8)\b|"
    r"\b[1234]\s*(?:p|pole)\b|\bsch(?:edule)?\s*(?:40|80)\b|\bnema\s*[a-z0-9-]+\b|"
    r"\b\d{1,2}/\d{1,2}\b|\b\d+(?:\.\d+)?\s*(?:cu|al)\b)", re.I
)
_CATALOG_TOKEN_RE = re.compile(r"\b(?=[A-Z0-9._/#\-]{4,35}\b)(?=[A-Z0-9._/#\-]*[A-Z])(?=[A-Z0-9._/#\-]*\d)[A-Z0-9._/#\-]+\b", re.I)
_FAILURE_PROSE_RE = re.compile(
    r"no\s+(?:readable|extractable)?\s*text|could\s+not\s+(?:be\s+)?extract|failed\s+to\s+(?:parse|read|extract)|"
    r"ocr\s+failed|parser\s+(?:error|failed)|processing\s+completed\s+but\s+produced\s+an\s+empty\s+result|#{4,}",
    re.I,
)


def _clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def _domain(email: str) -> str:
    value = str(email or "").strip().lower()
    return value.rsplit("@", 1)[1] if "@" in value else ""


def _normalized_domains(values: Iterable[str]) -> Set[str]:
    result: Set[str] = set()
    for value in values or ():
        for token in re.split(r"[,;\s]+", str(value or "").strip().lower()):
            token = token.lstrip("@").strip(" .")
            if token and "." in token:
                result.add(token)
    return result


def configured_internal_domains(settings: Optional[Mapping[str, Any]] = None, own_domains: Iterable[str] = ()) -> Set[str]:
    values: List[str] = list(own_domains or ())
    settings = settings or {}
    configured = settings.get("internal_domains", ())
    if isinstance(configured, str):
        values.extend(re.split(r"[,;\n]+", configured))
    elif isinstance(configured, (list, tuple, set)):
        values.extend(str(x) for x in configured)
    env = os.environ.get("MAINBOX_INTERNAL_DOMAINS", "")
    if env:
        values.extend(re.split(r"[,;\n]+", env))
    return _normalized_domains(values)


def is_internal_sender(sender_email: str, internal_domains: Iterable[str]) -> bool:
    sender_domain = _domain(sender_email)
    domains = _normalized_domains(internal_domains)
    if not sender_domain or not domains:
        return False
    return any(sender_domain == d or sender_domain.endswith("." + d) for d in domains)


def _numeric_qty(value: Any) -> Optional[float]:
    text = _clean_text(value).replace(",", "")
    if not text:
        return None
    m = re.fullmatch(r"\d+(?:\.\d+)?", text)
    if not m:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    if not math.isfinite(number) or number <= 0 or number > 1_000_000_000:
        return None
    return number


def _canonical_uom(value: Any) -> str:
    text = _clean_text(value).lower().rstrip(".")
    return _UOM_ALIASES.get(text, text.upper() if text else "")


def _strong_part(value: Any) -> bool:
    text = _clean_text(value).upper()
    if not text or len(text) > 45 or " " in text:
        return False
    if _PRICE_ONLY_RE.fullmatch(text) or _QTY_ONLY_RE.fullmatch(text):
        return False
    return bool(_CATALOG_TOKEN_RE.fullmatch(text))


def _item_line(item: Mapping[str, Any]) -> str:
    values = [
        item.get("qty", ""), item.get("unit", ""), item.get("part_number", ""),
        item.get("description", ""), item.get("manufacturer", ""), item.get("notes", ""),
    ]
    return _clean_text(" ".join(str(v or "") for v in values))


def _metadata_label(text: str) -> bool:
    low = re.sub(r"[^a-z0-9#]+", " ", text.lower()).strip()
    if low in _METADATA_LABELS:
        return True
    return any(low == label or low.startswith(label + " ") for label in _COMMERCIAL_HEADINGS)


def _has_material_word(text: str) -> bool:
    low = text.lower()
    return any(re.search(r"\b" + re.escape(word) + r"\b", low) for word in _MATERIAL_WORDS)


def _all_caps_personish(text: str) -> bool:
    if not _PERSONISH_RE.fullmatch(text.strip()):
        return False
    return not _has_material_word(text) and not _CATALOG_TOKEN_RE.search(text)


def _price_evidence_from_item(item: Mapping[str, Any], line: str) -> bool:
    explicit = any(_clean_text(item.get(k, "")) for k in ("price", "unit_price", "extended_price", "total_price"))
    return explicit or bool(_PRICE_RE.search(line))


def assess_candidate_row(item: Mapping[str, Any], index: int, policy: IntakePolicy) -> RowAssessment:
    normalized = dict(item or {})
    qty = _numeric_qty(normalized.get("qty", ""))
    unit = _canonical_uom(normalized.get("unit", ""))
    part = _clean_text(normalized.get("part_number", "")).upper()
    desc = _clean_text(normalized.get("description", ""))
    notes = _clean_text(normalized.get("notes", ""))
    line = _item_line(normalized)
    low = line.lower()
    reasons: List[str] = []
    score = 0.0
    price_evidence = _price_evidence_from_item(normalized, line)
    metadata_evidence = _metadata_label(desc or line)
    hard = ""

    if not line or len(line) < 2:
        hard = "empty row"
    elif _FAILURE_PROSE_RE.search(line):
        hard = "extractor failure prose"
    elif metadata_evidence:
        hard = "document metadata/header"
    elif _PAYMENT_TERMS_RE.fullmatch(desc or line):
        hard = "payment terms"
    elif _DATE_RE.fullmatch(desc or line):
        hard = "date-only row"
    elif _PRICE_ONLY_RE.fullmatch(desc or line) or _NUMBER_RANGE_RE.fullmatch(desc or line):
        hard = "price/number-only row"
    elif _QTY_ONLY_RE.fullmatch(desc or line) and not part and not _has_material_word(desc or line):
        hard = "quantity-only row"
    elif _all_caps_personish(desc or line):
        hard = "person/name row"
    elif _ADDRESS_RE.search(desc or line) and not _has_material_word(desc or line) and not _SPEC_RE.search(desc or line):
        hard = "address/location row"
    elif low in {"yes", "no", "n/a", "na", "none", "new", "open", "closed"}:
        hard = "administrative value"

    if hard:
        return RowAssessment(index, False, -10.0, normalized, reasons, hard, price_evidence, metadata_evidence)

    if qty is not None:
        score += 3.0
        reasons.append("valid quantity")
    if unit and unit in set(_UOM_ALIASES.values()):
        score += 1.0
        reasons.append("recognized UOM")
    if _strong_part(part):
        score += 4.0
        reasons.append("catalog/part number")
    elif part:
        score -= 1.0
        reasons.append("weak part-number field")
    if _has_material_word(desc):
        score += 3.0
        reasons.append("material/product term")
    if _SPEC_RE.search(desc):
        score += 2.0
        reasons.append("electrical/material specification")
    catalog_in_desc = bool(_CATALOG_TOKEN_RE.search(desc))
    if catalog_in_desc:
        score += 2.0
        reasons.append("catalog token in description")
    tokens = re.findall(r"[A-Za-z0-9#./\-]+", desc)
    if 2 <= len(tokens) <= 24:
        score += 1.0
        reasons.append("plausible product description")
    if price_evidence:
        score -= 3.0
        reasons.append("pricing evidence")
    if not qty and not _strong_part(part) and not catalog_in_desc and not _has_material_word(desc):
        score -= 4.0
        reasons.append("no positive item structure")
    if len(desc) > 180:
        score -= 1.0
        reasons.append("description unusually long")

    accepted = score >= policy.min_row_score
    if accepted:
        normalized["unit"] = unit or _clean_text(normalized.get("unit", ""))
        normalized["part_number"] = part
        normalized["description"] = desc
        normalized.setdefault("intake_engine_version", ENGINE_VERSION)
        normalized.setdefault("intake_score", round(score, 2))
    return RowAssessment(index, accepted, score, normalized, reasons, "", price_evidence, metadata_evidence)


class CoverageIntakeEngine:
    """Conservative document and row classifier for requested-item ingestion."""

    def __init__(self, policy: Optional[IntakePolicy] = None):
        self.policy = policy or IntakePolicy()

    def evaluate_customer_request(
        self,
        candidate_items: Sequence[Mapping[str, Any]],
        context: Optional[IntakeContext] = None,
        raw_text: str = "",
    ) -> IntakeResult:
        context = context or IntakeContext()
        items = [dict(x) for x in (candidate_items or []) if isinstance(x, Mapping)]
        combined_text = "\n".join(
            x for x in [context.subject, context.body, raw_text, *[_item_line(i) for i in items]] if x
        )
        low = combined_text.lower()

        internal = is_internal_sender(context.sender_email, context.internal_domains) or context.sender_type.upper() == "E"
        assessments = [assess_candidate_row(item, idx, self.policy) for idx, item in enumerate(items)]
        accepted = [a for a in assessments if a.accepted]
        rejected = [a for a in assessments if not a.accepted]
        total = max(1, len(assessments))
        metadata_count = sum(1 for a in assessments if a.metadata_evidence or a.hard_reject_reason in {
            "document metadata/header", "payment terms", "date-only row", "person/name row", "address/location row", "administrative value"
        })
        price_count = sum(1 for a in assessments if a.price_evidence or a.hard_reject_reason == "price/number-only row")
        price_only_count = sum(1 for a in assessments if a.hard_reject_reason == "price/number-only row")
        material_ratio = len(accepted) / total
        metadata_ratio = metadata_count / total
        price_ratio = price_count / total
        price_only_ratio = price_only_count / total
        failure_text = bool(_FAILURE_PROSE_RE.search(combined_text[:12000]))
        request_language = bool(_REQUEST_PHRASES.search(context.subject + "\n" + context.body))
        commercial_heading_hits = sum(1 for heading in _COMMERCIAL_HEADINGS if re.search(r"\b" + re.escape(heading) + r"\b", low))
        quote_subject = bool(re.search(r"^\s*(?:re:\s*)?(?:bid|quote|quotation|proposal)\b|\b(?:bid|quotation)\s*[#a-z0-9-]*", context.subject, re.I))
        po_hits = len(re.findall(r"\b(?:purchase order|customer po|po number|order qty|ship to|bill to)\b", low, re.I))
        invoice_hits = len(re.findall(r"\b(?:invoice|amount due|remit to|invoice date)\b", low, re.I))
        packing_hits = len(re.findall(r"\b(?:packing list|packing slip|shipped qty|tracking number)\b", low, re.I))
        submittal_hits = len(re.findall(r"\b(?:submittal|shop drawing|approval drawing|cutsheet|cut sheet)\b", low, re.I))

        role = DocumentRole.UNKNOWN
        role_confidence = 0.35
        if internal:
            role = DocumentRole.INTERNAL_BID if (quote_subject or commercial_heading_hits >= 2 or price_count >= 2) else DocumentRole.UNKNOWN
            role_confidence = 0.99
        elif invoice_hits >= 2:
            role, role_confidence = DocumentRole.INVOICE, 0.98
        elif packing_hits >= 1:
            role, role_confidence = DocumentRole.PACKING_LIST, 0.96
        elif po_hits >= 3 and commercial_heading_hits >= 2:
            role, role_confidence = DocumentRole.PURCHASE_ORDER, 0.96
        elif submittal_hits >= 1 and len(accepted) == 0:
            role, role_confidence = DocumentRole.SUBMITTAL, 0.90
        elif (commercial_heading_hits >= 3 and price_count >= 2) or (quote_subject and price_count >= 2):
            role, role_confidence = DocumentRole.VENDOR_QUOTE, min(0.99, 0.80 + 0.03 * commercial_heading_hits + 0.02 * price_count)
        elif request_language and len(accepted) >= 1 and price_ratio <= self.policy.max_price_only_ratio:
            role, role_confidence = DocumentRole.CUSTOMER_RFQ, min(0.98, 0.70 + 0.25 * material_ratio)
        elif len(accepted) >= 2 and material_ratio >= 0.75 and commercial_heading_hits <= 1 and price_ratio <= 0.10:
            # A clean BOM/takeoff may arrive with a vague subject, but accepting a
            # single plausible row without request language is too easy to confuse
            # with a packing slip, submittal, or casual attachment.
            role, role_confidence = DocumentRole.BOM, min(0.94, 0.68 + 0.25 * material_ratio)

        warnings: List[str] = []
        decision = IntakeDecision.QUARANTINE
        review_reason = ""

        if failure_text:
            decision = IntakeDecision.QUARANTINE
            review_reason = "Attachment text appears to contain extractor/OCR failure output."
        elif internal and self.policy.reject_internal_senders:
            decision = IntakeDecision.REJECT
            review_reason = "Internal sender cannot create a customer-request coverage record."
        elif role in {
            DocumentRole.VENDOR_QUOTE, DocumentRole.INTERNAL_BID, DocumentRole.SALES_ORDER,
            DocumentRole.PURCHASE_ORDER, DocumentRole.INVOICE, DocumentRole.PACKING_LIST,
            DocumentRole.SUBMITTAL,
        }:
            decision = IntakeDecision.REJECT
            review_reason = f"Document classified as {role.value.replace('_', ' ')} rather than a customer material request."
        elif not assessments:
            decision = IntakeDecision.QUARANTINE
            review_reason = "No candidate rows were extracted from the attachment."
        elif len(assessments) > self.policy.max_auto_rows:
            decision = IntakeDecision.QUARANTINE
            review_reason = f"Extraction produced {len(assessments)} rows, above the automatic-ingestion limit."
        elif role in {DocumentRole.CUSTOMER_RFQ, DocumentRole.BOM}:
            safe_ratios = (
                material_ratio >= self.policy.min_material_ratio
                and metadata_ratio <= self.policy.max_metadata_ratio
                and price_only_ratio <= self.policy.max_price_only_ratio
            )
            confidence = min(role_confidence, 0.45 + 0.50 * material_ratio - 0.30 * metadata_ratio - 0.25 * price_only_ratio)
            if accepted and safe_ratios and confidence >= self.policy.min_accept_confidence:
                decision = IntakeDecision.ACCEPT
            else:
                decision = IntakeDecision.QUARANTINE
                review_reason = "Material-list quality thresholds were not met safely enough for automatic coverage."
        elif accepted:
            decision = IntakeDecision.QUARANTINE
            review_reason = "Some plausible material rows were found, but the document role is ambiguous."
        else:
            decision = IntakeDecision.REJECT
            review_reason = "No rows contained sufficient positive material evidence."

        confidence = max(0.0, min(1.0, role_confidence))
        if metadata_ratio > self.policy.max_metadata_ratio:
            warnings.append(f"High metadata/header ratio: {metadata_count}/{len(assessments) or 0} rows.")
        if price_ratio > self.policy.max_price_only_ratio:
            warnings.append(f"Pricing evidence appears on {price_count}/{len(assessments) or 0} rows.")
        if rejected:
            warnings.append(f"Rejected {len(rejected)} row(s) lacking positive material evidence.")

        accepted_items = [dict(a.item) for a in accepted] if decision == IntakeDecision.ACCEPT else []
        metrics = {
            "candidate_rows": len(assessments),
            "accepted_rows": len(accepted),
            "rejected_rows": len(rejected),
            "material_ratio": round(material_ratio, 4),
            "metadata_rows": metadata_count,
            "metadata_ratio": round(metadata_ratio, 4),
            "price_rows": price_count,
            "price_ratio": round(price_ratio, 4),
            "price_only_rows": price_only_count,
            "price_only_ratio": round(price_only_ratio, 4),
            "commercial_heading_hits": commercial_heading_hits,
            "request_language": request_language,
            "internal_sender": internal,
            "failure_text": failure_text,
        }
        audit = {
            "engine_version": ENGINE_VERSION,
            "source_file": context.source_file,
            "message_id": context.message_id,
            "sender_email": context.sender_email,
            "subject": context.subject,
            "role_confidence": round(role_confidence, 4),
        }
        return IntakeResult(
            decision=decision,
            document_role=role,
            confidence=confidence,
            accepted_items=accepted_items,
            rejected_rows=rejected,
            accepted_rows=accepted,
            warnings=warnings,
            review_reason=review_reason,
            metrics=metrics,
            audit=audit,
        )

    def analyze_existing_requested_record(self, record: Mapping[str, Any]) -> Dict[str, Any]:
        """Identify likely contaminated requested rows without mutating the record."""
        requested = record.get("requested", []) if isinstance(record, Mapping) else []
        context = IntakeContext(
            subject=_clean_text(record.get("subject", "")),
            sender_email=_clean_text(record.get("customer_email", "")),
            sender_name=_clean_text(record.get("customer", "")),
            source_file="existing-ledger",
        )
        result = self.evaluate_customer_request(requested if isinstance(requested, list) else [], context)
        likely_contaminated = bool(
            result.decision in {IntakeDecision.REJECT, IntakeDecision.QUARANTINE}
            and result.metrics.get("candidate_rows", 0) >= 5
            and (
                result.metrics.get("metadata_ratio", 0) > self.policy.max_metadata_ratio
                or result.metrics.get("price_ratio", 0) > self.policy.max_price_only_ratio
                or result.document_role in {DocumentRole.VENDOR_QUOTE, DocumentRole.INTERNAL_BID}
            )
        )
        return {
            "likely_contaminated": likely_contaminated,
            "result": result.to_dict(),
            "suggested_remove_line_ids": [
                str(a.item.get("line_id", "") or "") for a in result.rejected_rows if a.item.get("line_id")
            ],
        }


def evaluate_customer_request_document(
    candidate_items: Sequence[Mapping[str, Any]],
    *,
    subject: str = "",
    body: str = "",
    sender_email: str = "",
    sender_name: str = "",
    sender_type: str = "",
    internal_domains: Iterable[str] = (),
    source_file: str = "",
    message_id: str = "",
    conversation_id: str = "",
    raw_text: str = "",
    policy: Optional[IntakePolicy] = None,
) -> Dict[str, Any]:
    engine = CoverageIntakeEngine(policy=policy)
    result = engine.evaluate_customer_request(
        candidate_items,
        IntakeContext(
            subject=subject,
            body=body,
            sender_email=sender_email,
            sender_name=sender_name,
            sender_type=sender_type,
            internal_domains=tuple(internal_domains or ()),
            source_file=source_file,
            message_id=message_id,
            conversation_id=conversation_id,
        ),
        raw_text=raw_text,
    )
    return result.to_dict()


__all__ = [
    "ENGINE_VERSION", "CoverageIntakeEngine", "IntakeContext", "IntakeDecision",
    "DocumentRole", "IntakePolicy", "IntakeResult", "RowAssessment",
    "configured_internal_domains", "is_internal_sender", "evaluate_customer_request_document",
]
