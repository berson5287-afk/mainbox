"""MaINbox Auto-Group Engine plug-in.

A deterministic, explainable group-ranking engine designed to sit between
Outlook import and MaINbox's legacy grouping shortcuts.

Public API
----------
    engine = AutoGroupEngine(settings, state_path)
    decision = engine.decide(email, groups_data, emails, quote_coverage, learning_data)
    engine.record_feedback(email, target_group_id, previous_group_id, groups_data)

The module has no Outlook, Tk, or third-party dependencies.
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from datetime import datetime
from difflib import SequenceMatcher
import hashlib
import json
import math
import os
import re
import tempfile
import zlib  # v1.1.0: cheap order-independent membership signature
from typing import Any, Dict, Iterable, List, Mapping, MutableMapping, Optional, Sequence, Set, Tuple

ENGINE_VERSION = "1.2.1"
# v1.2.0 CHANGELOG (integration-audit batch; three deliberate behavior changes, all
# in the precision-improving direction, each individually tested):
#   1. IDENTIFIER CAPTURE STOPS AT THE NEXT LABEL: 'Job 1598 PO 100' now extracts
#      JOB=1598 + PO=100 (was JOB=1598PO100 -- the char class included space and
#      swallowed the following label, manufacturing false job-number conflicts).
#   2. DETERMINISTIC LINEAGE OVERRIDES ORDINARY ID CONFLICTS: when a candidate group
#      matches by MaINbox reference, reply-header Message-ID, or Outlook conversation,
#      a conflicting PO/JOB/RFQ/QUOTE number no longer hard-vetoes it -- a new PO
#      arriving inside the same quote conversation is the normal customer workflow.
#      The conflict becomes a visible 'review job metadata' note plus a moderate
#      penalty. Conflicting customer identity, a DIFFERENT MaINbox reference, and a
#      user rejection remain hard vetoes (deterministic or human contrary evidence).
#   3. RELATIONSHIP IDENTITY INVALIDATES CACHES: email_type / sender_type / sender
#      address are now part of both the per-email feature-cache validity key and the
#      per-group membership crc, so a Vendor->Customer correction recomputes that
#      email's features and rebuilds its group's fingerprint (previously stale).
#   Also: nullable groups/memberships payloads are normalized at every public entry
#   (a structurally valid {'groups': None} can no longer raise), and material
#   description-candidate capping now ranks by token overlap THEN quantity/UOM
#   compatibility, so among near-identical commodity rows the best quantity match
#   survives the cap (closes the v1.1.0 under-match on duplicate descriptions).
# v1.1.0 CHANGELOG (performance -- no scoring-behavior change intended):
#   1. Per-email feature cache: business ids, refs, subjects, materials, and entities are
#      extracted once per email and reused, so a fingerprint rebuild after invalidate()
#      (save_groups fires one per auto-group during bulk refresh) aggregates cached
#      features instead of re-regexing every body. Measured: 10-email refresh storm went
#      from ~10s to well under 1s at 3,000 emails / 150 groups.
#   2. material_set_similarity: token inverted index + exact-part dict + Jaccard gate.
#      Provably equivalent gate: with threshold 0.42 and weights 0.65*jac + 0.35*seq, a
#      pair with too little word overlap can never pass, so SequenceMatcher is skipped
#      for those pairs. Per-row word sets / normalized descriptions are computed once per
#      Material instead of once per pair. Measured: one materials-heavy decide() went
#      from 194s to <0.1s at 200 groups averaging ~594 material rows each.
#   3. Bounded fingerprints: subjects capped (newest first), materials capped after
#      dedupe with parted rows kept preferentially, customer names capped.
#   4. _signature(): order-independent crc32 accumulation replaces sort+join over all
#      memberships (was O(M log M) plus a large string build on every decide()).

PUBLIC_DOMAINS = {
    "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com", "msn.com",
    "yahoo.com", "ymail.com", "aol.com", "icloud.com", "me.com", "mac.com",
    "protonmail.com", "proton.me", "pm.me", "gmx.com", "gmx.net", "mail.com",
    "zoho.com", "fastmail.com", "comcast.net", "verizon.net", "att.net", "cox.net",
    "optonline.net", "sbcglobal.net", "bellsouth.net", "charter.net", "earthlink.net",
}

GENERIC_WORDS = {
    "re", "fw", "fwd", "external", "quote", "quotation", "rfq", "request", "pricing",
    "price", "availability", "estimate", "proposal", "order", "purchase", "po", "job",
    "project", "bid", "attached", "attachment", "please", "thanks", "thank", "your",
    "the", "and", "for", "from", "with", "this", "that", "items", "item", "material",
    "materials", "email", "reply", "new", "additional", "revised", "revision",
}

PRODUCT_WORDS = {
    "breaker", "connector", "coupling", "conduit", "emt", "rigid", "pvc", "wire", "cable",
    "thhn", "thwn", "mc", "panel", "switch", "receptacle", "outlet", "fitting", "bushing",
    "lug", "splice", "tape", "strut", "channel", "box", "cover", "plate", "transformer",
    "disconnect", "fuse", "relay", "contact", "contactor", "enclosure", "meter", "light",
    "fixture", "lamp", "ballast", "driver", "ground", "rod", "clamp", "tray", "raceway",
    "elbow", "nipples", "nipple", "adapter", "adaptor", "hub", "device", "starter",
}

UOM_ALIASES = {
    "EA": "EA", "EACH": "EA", "PC": "EA", "PCS": "EA", "PIECE": "EA", "PIECES": "EA",
    "FT": "FT", "FOOT": "FT", "FEET": "FT", "LF": "FT",
    "M": "M", "METER": "M", "METERS": "M",
    "C": "C", "HUNDRED": "C", "MFT": "MFT", "THOUSAND": "MFT",
    "ROLL": "ROLL", "RLS": "ROLL", "BOX": "BOX", "BX": "BOX", "BAG": "BAG",
    "SET": "SET", "PAIR": "PAIR", "LOT": "LOT",
}

ID_KEYWORDS = {
    "po": r"(?:PO|P\.?O\.?|PURCHASE\s+ORDER)",
    "job": r"(?:JOB|PROJECT|PROJ)",
    "rfq": r"(?:RFQ|REQUEST\s+FOR\s+QUOTE)",
    "quote": r"(?:QUOTE|QUOTATION|QTE)",
    "order": r"(?:SALES\s+ORDER|SO|ORDER)",
    "ref": r"(?:REF|REFERENCE)",
}


@dataclass
class EngineSettings:
    enabled: bool = True
    auto_threshold: float = 0.88
    suggest_threshold: float = 0.62
    min_margin: float = 0.13
    require_two_families: bool = True
    max_candidates: int = 5
    recent_days: int = 180
    use_materials: bool = True
    material_body_limit: int = 7000
    max_group_member_scan: int = 40
    deterministic_lineage_auto: bool = True

    @classmethod
    def from_mapping(cls, values: Optional[Mapping[str, Any]]) -> "EngineSettings":
        values = values or {}
        def f(name: str, default: float, lo: float, hi: float) -> float:
            try:
                return max(lo, min(hi, float(values.get(name, default))))
            except Exception:
                return default
        def i(name: str, default: int, lo: int, hi: int) -> int:
            try:
                return max(lo, min(hi, int(values.get(name, default))))
            except Exception:
                return default
        return cls(
            enabled=bool(values.get("auto_group_engine_enabled", values.get("enabled", True))),
            auto_threshold=f("auto_group_auto_threshold", 0.88, 0.55, 0.99),
            suggest_threshold=f("auto_group_suggest_threshold", 0.62, 0.30, 0.95),
            min_margin=f("auto_group_min_margin", 0.13, 0.03, 0.50),
            require_two_families=bool(values.get("auto_group_require_two_families", True)),
            max_candidates=i("auto_group_max_candidates", 5, 1, 10),
            recent_days=i("auto_group_recent_days", 180, 14, 3650),
            use_materials=bool(values.get("auto_group_use_materials", True)),
            material_body_limit=i("auto_group_material_body_limit", 7000, 1000, 30000),
            max_group_member_scan=i("auto_group_max_member_scan", 40, 5, 250),
            deterministic_lineage_auto=bool(values.get("auto_group_deterministic_lineage_auto", True)),
        )


@dataclass
class Material:
    part: str = ""
    description: str = ""
    qty: Optional[float] = None
    uom: str = ""
    # v1.1.0: per-row features computed once (previously words()/normalize_description()
    # ran inside every pair comparison -- the dominant cost of material scoring).
    dwords: Optional[frozenset] = None
    ndesc: Optional[str] = None


def _prep_material(m: "Material") -> "Material":
    """v1.1.0: fill cached description features exactly once per Material."""
    if m.dwords is None:
        m.ndesc = normalize_description(m.description)
        m.dwords = frozenset(words(m.description))
    return m


@dataclass
class GroupFingerprint:
    group_id: str
    name: str
    archived: bool = False
    created_at: str = ""
    last_activity: float = 0.0
    active_count: int = 0
    member_ids: Set[str] = field(default_factory=set)
    conversation_ids: Set[str] = field(default_factory=set)
    message_ids: Set[str] = field(default_factory=set)
    references: Set[str] = field(default_factory=set)
    subjects: Set[str] = field(default_factory=set)
    subject_words: Set[str] = field(default_factory=set)
    business_ids: Dict[str, Set[str]] = field(default_factory=dict)
    loose_tokens: Set[str] = field(default_factory=set)
    customer_emails: Set[str] = field(default_factory=set)
    customer_domains: Set[str] = field(default_factory=set)
    customer_names: Set[str] = field(default_factory=set)
    vendor_domains: Set[str] = field(default_factory=set)
    materials: List[Material] = field(default_factory=list)
    mainbox_refs: Set[str] = field(default_factory=set)
    # v1.1.0: lazily built material index (part dict, token postings, part-less rows).
    # Built once per fingerprint lifetime and reused across every decide() -- rebuilding
    # it per candidate per email was the dominant warm-path cost after the SM gate.
    material_index: Optional[Tuple[Dict[str, List[int]], Dict[str, List[int]], List[int]]] = None


@dataclass
class CandidateScore:
    group_id: str
    group_name: str
    score: float
    confidence: float
    families: Set[str] = field(default_factory=set)
    reasons: List[str] = field(default_factory=list)
    conflicts: List[str] = field(default_factory=list)
    hard_veto: bool = False
    details: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d["families"] = sorted(self.families)
        return d


@dataclass
class GroupDecision:
    action: str = "none"          # auto | suggest | none
    group_id: str = ""
    group_name: str = ""
    confidence: float = 0.0
    score: float = 0.0
    margin: float = 0.0
    reason: str = ""
    evidence_families: List[str] = field(default_factory=list)
    candidates: List[Dict[str, Any]] = field(default_factory=list)
    engine_version: str = ENGINE_VERSION

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def _text(value: Any) -> str:
    return str(value or "")


def normalize_email(value: Any) -> str:
    s = _text(value).strip().lower()
    m = re.search(r"[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}", s, flags=re.I)
    return m.group(0).lower() if m else (s if "@" in s else "")


def email_domain(value: Any) -> str:
    e = normalize_email(value)
    return e.split("@", 1)[1] if "@" in e else ""


def normalize_subject(value: Any) -> str:
    s = _text(value).strip().lower()
    while True:
        n = re.sub(r"^\s*(?:re|fw|fwd|external)\s*:\s*", "", s, flags=re.I).strip()
        if n == s:
            break
        s = n
    s = re.sub(r"\[[^\]]{1,60}\]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def compact(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", _text(value).lower())


def words(value: Any) -> Set[str]:
    out: Set[str] = set()
    for w in re.findall(r"[a-z0-9][a-z0-9._/&+\-]*", normalize_subject(value)):
        c = compact(w)
        if len(c) >= 3 and c not in GENERIC_WORDS:
            out.add(c)
    return out


def normalize_message_id(value: Any) -> str:
    s = _text(value).strip().lower()
    if not s:
        return ""
    m = re.search(r"<[^<>\s]+>", s)
    return m.group(0) if m else s


def message_ids_from_header(value: Any) -> Set[str]:
    s = _text(value).lower()
    found = {x.strip().lower() for x in re.findall(r"<[^<>\s]+>", s)}
    if not found and s.strip():
        found.add(s.strip())
    return found


def extract_mainbox_refs(value: Any) -> Set[str]:
    return {m.upper() for m in re.findall(r"\bMBX[-_ ]?[A-Z0-9]{5,16}\b", _text(value), flags=re.I)}


def _clean_identifier(value: Any) -> str:
    s = _text(value).upper().strip()
    s = re.split(r"\s{2,}|[,;|]", s, maxsplit=1)[0]
    s = re.sub(r"\b(?:THE|FOR|FROM|PLEASE|ATTACHED|REVISED|REVISION)\b.*$", "", s).strip()
    c = re.sub(r"[^A-Z0-9]", "", s)
    if len(c) < 3 or not any(ch.isdigit() for ch in c):
        return ""
    return c[:32]


_ID_LABEL_SPLIT_RE = re.compile(
    r"\s+(?:P\.?O\.?|JOB|PROJECT|PROJ|RFQ|QUOTE|QUOTATION|ORDER|REF|REFERENCE|SO|WO|W/O)\b",
    re.I)  # v1.2.0: identifier capture terminates at the next label


def extract_business_ids(value: Any) -> Dict[str, Set[str]]:
    raw = _text(value)
    upper = raw.upper()
    out: Dict[str, Set[str]] = {k: set() for k in ID_KEYWORDS}
    for kind, kw in ID_KEYWORDS.items():
        pattern = rf"\b{kw}(?![A-Z])\s*(?:#|NO\.?|NUMBER|:|-)?\s*([A-Z0-9][A-Z0-9./_&\- ]{{1,30}})"
        for m in re.finditer(pattern, upper, flags=re.I):
            raw_ident = m.group(1)
            # v1.2.0: stop at the next recognized label so 'JOB 1598 PO 100' yields
            # JOB=1598 (and the PO keyword's own pass yields PO=100) instead of the
            # swallowed 'JOB=1598PO100' that manufactured false conflicts.
            raw_ident = _ID_LABEL_SPLIT_RE.split(raw_ident, 1)[0]
            ident = _clean_identifier(raw_ident)
            if ident:
                out[kind].add(ident)
    # Strong unanchored codes, useful when a reply drops the JOB/PO label.
    loose: Set[str] = set()
    for pat in (
        r"\b[A-Z]{1,6}[-&][A-Z0-9]{2,}(?:[-/][A-Z0-9]{2,})?\b",
        r"\b\d{2,8}[-/][A-Z0-9]{2,10}\b",
        r"\b[A-Z]{1,5}\d{3,9}[A-Z0-9]{0,5}\b",
        r"\b\d{3,9}[A-Z]{1,5}\b",
    ):
        for m in re.finditer(pat, upper):
            c = compact(m.group(0)).upper()
            # v1.2.1: a loose (unlabeled) code must contain BOTH a letter and a digit.
            # Pure-numeric fragments are almost always street numbers or quantities
            # ("31-03" -> "3103", "20-14" -> "2014") rather than job/quote codes, and
            # they manufacture false id matches between unrelated jobs at nearby
            # addresses. Labeled ids (JOB/PO/...) are unaffected: they parse via
            # _clean_identifier, which requires only a digit.
            if len(c) >= 4 and any(ch.isalpha() for ch in c) and any(ch.isdigit() for ch in c):
                loose.add(c)
    out["loose"] = loose
    return out


def extract_named_entities(subject: Any, body: Any, sender: Any = "") -> Dict[str, Set[str]]:
    text = f"{_text(subject)}\n{_text(body)[:4000]}"
    customer_names: Set[str] = set()
    job_names: Set[str] = set()
    for label, target in ((r"CUSTOMER|CLIENT|ACCOUNT", customer_names), (r"JOB|PROJECT|SITE", job_names)):
        for m in re.finditer(rf"\b(?:{label})\s*(?:NAME)?\s*[:#-]\s*([^\r\n|;]{{3,80}})", text, flags=re.I):
            val = re.sub(r"\s+", " ", m.group(1)).strip(" -:,.\t")
            if val and not re.fullmatch(r"(?:NUMBER|NO|N/?A)", val, flags=re.I):
                target.add(val.lower())
    # Display-name portion of a human sender can be weak customer evidence.
    s = _text(sender).strip()
    if s and "@" not in s and len(s) <= 80 and not re.search(r"no-?reply|system|sales|quotes?", s, re.I):
        customer_names.add(s.lower())
    return {"customer_names": customer_names, "job_names": job_names}


def normalize_uom(value: Any) -> str:
    s = re.sub(r"[^A-Z]", "", _text(value).upper())
    return UOM_ALIASES.get(s, s if len(s) <= 5 else "")


def normalize_part(value: Any) -> str:
    c = compact(value).upper()
    if len(c) < 3:
        return ""
    return c


def normalize_description(value: Any) -> str:
    s = _text(value).lower()
    s = re.sub(r"\$\s*\d[\d,]*(?:\.\d+)?", " ", s)
    s = re.sub(r"\b\d+(?:\.\d+)?\s*/\s*(?:ea|each|c|m)\b", " ", s, flags=re.I)
    s = re.sub(r"[^a-z0-9#&/.'\-]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _float_or_none(value: Any) -> Optional[float]:
    try:
        if value is None or value == "":
            return None
        return float(str(value).replace(",", "").strip())
    except Exception:
        return None


def material_from_mapping(row: Mapping[str, Any]) -> Optional[Material]:
    part = normalize_part(row.get("part_number") or row.get("part") or row.get("catalog") or row.get("sku"))
    desc = normalize_description(row.get("description") or row.get("item") or row.get("display") or row.get("requested_item"))
    qty = _float_or_none(row.get("qty") if "qty" in row else row.get("quantity"))
    uom = normalize_uom(row.get("unit") or row.get("uom"))
    if not part and len(desc) < 3:
        return None
    return Material(part=part, description=desc, qty=qty, uom=uom)


def parse_material_lines(value: Any, limit: int = 7000) -> List[Material]:
    """Conservative body-only extractor for grouping fingerprints.

    It intentionally ignores price-only/header lines and accepts only rows with
    quantity + product evidence or a strong catalog code + product wording.
    """
    out: List[Material] = []
    raw = _text(value)[:limit]
    for line in raw.splitlines():
        line = re.sub(r"\s+", " ", line).strip(" \t|•*-–—")
        if len(line) < 4 or len(line) > 220:
            continue
        low = line.lower()
        if re.search(r"\b(?:customer|salesperson|writer|ship via|terms|ship date|subtotal|total|freight|unit price|extended price|bill to|ship to)\b", low):
            continue
        if re.fullmatch(r"\$?\d[\d,]*(?:\.\d+)?(?:\s*/\s*(?:ea|c|m))?", low, flags=re.I):
            continue
        m = re.match(r"^\s*(\d[\d,]*(?:\.\d+)?)\s*(EA|EACH|PCS?|FT|FEET|LF|M|ROLLS?|BAGS?|SETS?|BOX|BX)?\s*[-:xX]?\s*(.+)$", line, flags=re.I)
        qty: Optional[float] = None
        uom = ""
        desc = line
        if m:
            qty = _float_or_none(m.group(1))
            uom = normalize_uom(m.group(2))
            desc = m.group(3).strip()
        part = ""
        candidates = re.findall(r"\b[A-Z0-9][A-Z0-9._/\-]{2,24}\b", desc, flags=re.I)
        for token in candidates:
            c = normalize_part(token)
            if any(ch.isdigit() for ch in c) and any(ch.isalpha() for ch in c) and c not in {"NET30", "NET60"}:
                part = c
                break
        dwords = words(desc)
        product = bool(dwords & PRODUCT_WORDS)
        spec = bool(re.search(r"(?:\b\d+(?:/\d+)?\s*(?:in|inch|\")|#\s*\d+|\b\d+\s*(?:amp|a|v|volt)|\b[123]\s*(?:ph|phase|pole)|\b\d+/\d+\b)", desc, flags=re.I))
        if qty is not None and (product or part or spec):
            out.append(Material(part=part, description=normalize_description(desc), qty=qty, uom=uom))
        elif part and (product or spec) and len(desc.split()) >= 2:
            out.append(Material(part=part, description=normalize_description(desc), qty=None, uom=uom))
    return out[:250]


def collect_email_materials(email: Mapping[str, Any], body_limit: int = 7000) -> List[Material]:
    out: List[Material] = []
    keys = ("materials", "rfq_items", "extracted_items", "quote_items", "requested_items", "coverage_items", "items")
    for key in keys:
        rows = email.get(key)
        if isinstance(rows, list):
            for row in rows:
                if isinstance(row, Mapping):
                    m = material_from_mapping(row)
                    if m:
                        out.append(m)
    if not out:
        out.extend(parse_material_lines(email.get("body", ""), body_limit))
    return _dedupe_materials(out)


def _dedupe_materials(rows: Iterable[Material]) -> List[Material]:
    out: List[Material] = []
    seen: Set[Tuple[Any, ...]] = set()
    for r in rows:
        key = (r.part, compact(r.description), round(r.qty, 6) if r.qty is not None else None, r.uom)
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def description_similarity(a: str, b: str) -> float:
    aw, bw = words(a), words(b)
    if not aw or not bw:
        return 0.0
    inter = aw & bw
    union = aw | bw
    jac = len(inter) / max(1, len(union))
    seq = SequenceMatcher(None, normalize_description(a), normalize_description(b)).ratio()
    return 0.65 * jac + 0.35 * seq


def _description_similarity_prepped(a: Material, b: Material, sm_budget: Optional[List[int]] = None) -> float:
    """v1.1.0: same formula as description_similarity, but on cached row features and
    with an exact early-out: max achievable is 0.65*jac + 0.35, so when that is below
    the 0.42 acceptance threshold the pair can never match and SequenceMatcher (the
    expensive part) is skipped. Identical accept/reject outcome, ~1000x cheaper on
    non-matching pairs.

    sm_budget is an optional single-element mutable counter shared across one decide()
    call. When exhausted (pathological inboxes where nearly every description shares
    words with every group row), the function returns the Jaccard component only --
    a strict LOWER bound of the true similarity, so scoring degrades toward
    under-matching (precision-safe), never over-matching."""
    aw = a.dwords if a.dwords is not None else frozenset(words(a.description))
    bw = b.dwords if b.dwords is not None else frozenset(words(b.description))
    if not aw or not bw:
        return 0.0
    inter_n = len(aw & bw)
    jac = inter_n / max(1, len(aw | bw))
    if 0.65 * jac + 0.35 < 0.42:  # even a perfect char-level match could not reach 0.42
        return 0.65 * jac
    if sm_budget is not None:
        if sm_budget[0] <= 0:
            return 0.65 * jac  # budget exhausted: conservative lower bound
        sm_budget[0] -= 1
    na = a.ndesc if a.ndesc is not None else normalize_description(a.description)
    nb = b.ndesc if b.ndesc is not None else normalize_description(b.description)
    seq = SequenceMatcher(None, na, nb).ratio()
    return 0.65 * jac + 0.35 * seq


def material_pair_score(a: Material, b: Material, sm_budget: Optional[List[int]] = None) -> Tuple[float, str]:
    if a.part and b.part:
        if a.part != b.part:
            return 0.0, "explicit part conflict"
        base, why = 1.0, "exact part"
    else:
        sim = _description_similarity_prepped(a, b, sm_budget)  # v1.1.0: gated, cached-feature path
        if sim < 0.42:
            return 0.0, ""
        base, why = min(0.90, 0.45 + sim * 0.5), "description/spec"
    if a.uom and b.uom and a.uom != b.uom:
        base *= 0.45
        why += ", UOM differs"
    if a.qty is not None and b.qty is not None and max(a.qty, b.qty) > 0:
        ratio = min(a.qty, b.qty) / max(a.qty, b.qty)
        base *= 0.72 + 0.28 * ratio
        why += f", qty {ratio:.0%}"
    return base, why


def build_material_index(group_rows: Sequence[Material]) -> Tuple[Dict[str, List[int]], Dict[str, List[str]], List[int], Dict[str, List[int]]]:
    """v1.2.0: (part dict, token postings over DISTINCT normalized descriptions,
    part-less row list, ndesc -> row indices). Postings are bounded per token by
    distinct descriptions (30), which is what the v1.1.0 row cap was trying to bound
    -- but duplicate rows of one description are now all reachable through their
    ndesc bucket instead of being truncated in insertion order, so the best
    quantity/UOM match among duplicates can always be selected."""
    by_part: Dict[str, List[int]] = {}
    token_index: Dict[str, List[str]] = {}
    partless_group: List[int] = []
    ndesc_rows: Dict[str, List[int]] = {}
    for j, b in enumerate(group_rows):
        _prep_material(b)
        if b.part:
            by_part.setdefault(b.part, []).append(j)
        else:
            partless_group.append(j)
        nd = b.ndesc or ""
        bucket = ndesc_rows.setdefault(nd, [])
        first_of_desc = not bucket
        bucket.append(j)
        if first_of_desc:
            for t in (b.dwords or ()):
                lst = token_index.setdefault(t, [])
                if len(lst) < 30:
                    lst.append(nd)
    return by_part, token_index, partless_group, ndesc_rows


def material_set_similarity(incoming: Sequence[Material], group_rows: Sequence[Material], sm_budget: Optional[List[int]] = None, prebuilt_index: Optional[Tuple[Dict[str, List[int]], Dict[str, List[str]], List[int], Dict[str, List[int]]]] = None) -> Dict[str, Any]:
    if not incoming or not group_rows:
        return {"score": 0.0, "matches": 0, "exact_parts": 0, "coverage": 0.0, "details": []}
    used: Set[int] = set()
    pairs: List[Tuple[float, int, int, str]] = []
    # v1.1.0: replace the full O(A x B) sweep with two indexes. Outcome-equivalent:
    #   - parted incoming vs parted group rows only ever match on IDENTICAL parts
    #     (different explicit parts score 0.0 in material_pair_score), so a dict on
    #     normalized part covers every scorable parted pair;
    #   - the description path requires word overlap to pass 0.42 (see the gate in
    #     _description_similarity_prepped), so a pair sharing no description token can
    #     never match -- a token inverted index therefore enumerates every pair that
    #     could possibly score > 0, and only those.
    by_part, token_index, partless_group, ndesc_rows = prebuilt_index if prebuilt_index is not None else build_material_index(group_rows)
    partless_set = set(partless_group)
    _MAX_DESC_CANDIDATES = 25  # safety bound for pathological common-token rows
    for i, a in enumerate(incoming):
        _prep_material(a)
        cand: Set[int] = set()
        if a.part:
            # exact-part matches, plus description matches against part-less rows only
            # (parted-vs-different-parted is always 0.0, exactly as before)
            cand.update(by_part.get(a.part, ()))
            restrict_partless = True
        else:
            restrict_partless = False  # part-less incoming may description-match any row, as before
        if a.dwords:
            # v1.2.0: overlap is counted per DISTINCT description; each selected
            # description then contributes its best-quantity-affinity rows, so
            # duplicates of a common commodity line no longer crowd out the row whose
            # quantity actually matches (v1.1.0 truncated postings in insertion order).
            overlap_count: Dict[str, int] = {}
            for t in a.dwords:
                for nd in token_index.get(t, ()):
                    overlap_count[nd] = overlap_count.get(nd, 0) + 1

            def _row_qty_affinity(j2: int) -> float:
                b2 = group_rows[j2]
                if a.qty is None or b2.qty is None or not a.qty or not b2.qty:
                    qa = 0.5
                else:
                    try:
                        qa = min(float(a.qty), float(b2.qty)) / max(float(a.qty), float(b2.qty))
                    except Exception:
                        qa = 0.5
                if a.uom and b2.uom and a.uom == b2.uom:
                    qa += 0.25
                return qa

            def _bucket_rows(nd2: str) -> List[int]:
                rows = ndesc_rows.get(nd2, ())
                if restrict_partless:
                    rows = [j2 for j2 in rows if j2 in partless_set]
                else:
                    rows = list(rows)
                if len(rows) > 8:  # greedy consumes each row once; 8 qty-ranked duplicates suffice
                    rows.sort(key=_row_qty_affinity, reverse=True)
                    rows = rows[:8]
                return rows

            items = overlap_count.items()
            if len(overlap_count) > _MAX_DESC_CANDIDATES:
                def _bucket_key(kv):
                    nd2, n2 = kv
                    rows = ndesc_rows.get(nd2, ())
                    best = max((_row_qty_affinity(j2) for j2 in rows), default=0.0)
                    return (n2, best)
                items = sorted(overlap_count.items(), key=_bucket_key, reverse=True)[:_MAX_DESC_CANDIDATES]
            for nd2, _n in items:
                cand.update(_bucket_rows(nd2))
        for j in cand:
            s, why = material_pair_score(a, group_rows[j], sm_budget)
            if s > 0:
                pairs.append((s, i, j, why))
    pairs.sort(reverse=True)
    chosen_in: Set[int] = set()
    details: List[str] = []
    exact = 0
    total = 0.0
    for s, i, j, why in pairs:
        if i in chosen_in or j in used:
            continue
        chosen_in.add(i); used.add(j); total += s
        if why.startswith("exact part"):
            exact += 1
        details.append(f"{incoming[i].part or incoming[i].description[:35]} ↔ {group_rows[j].part or group_rows[j].description[:35]} ({why})")
    coverage = len(chosen_in) / max(1, len(incoming))
    quality = total / max(1, len(incoming))
    # Exact part coverage is stronger than broad description coverage.
    score = min(1.0, 0.55 * coverage + 0.30 * quality + 0.15 * min(1.0, exact / max(1, min(2, len(incoming)))))
    return {"score": score, "matches": len(chosen_in), "exact_parts": exact, "coverage": coverage, "details": details[:8]}


def _safe_timestamp(value: Any) -> float:
    try:
        return float(value or 0)
    except Exception:
        return 0.0


def _record_key(email: Mapping[str, Any]) -> str:
    imid = normalize_message_id(email.get("internet_message_id"))
    if imid:
        return "imid:" + imid
    eid = _text(email.get("entry_id")).strip()
    if eid:
        return "eid:" + eid
    payload = "|".join([normalize_email(email.get("sender_email") or email.get("sender")), normalize_subject(email.get("subject")), _text(email.get("received_timestamp"))])
    return "fp:" + hashlib.sha256(payload.encode("utf-8", "ignore")).hexdigest()[:24]


def _atomic_json(path: str, data: Mapping[str, Any]) -> None:
    folder = os.path.dirname(os.path.abspath(path))
    os.makedirs(folder, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".autogroup_", suffix=".tmp", dir=folder)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2, sort_keys=True)
            fh.flush(); os.fsync(fh.fileno())
        os.replace(tmp, path)
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass


class AutoGroupEngine:
    def __init__(self, settings: Optional[Mapping[str, Any]] = None, state_path: str = ""):
        self.settings = EngineSettings.from_mapping(settings)
        self.state_path = state_path
        self.state: Dict[str, Any] = {"version": 1, "accepted": {}, "rejected": {}, "aliases": {}, "updated_at": ""}
        self._fingerprints: Dict[str, GroupFingerprint] = {}
        self._cache_signature = ""
        # v1.1.0: per-email extracted features, keyed by entry_id. Email bodies do not
        # change, so ids/refs/subjects/materials extracted from them are reused across
        # fingerprint rebuilds. This is what makes the save_groups() -> invalidate() ->
        # rebuild cycle cheap: a rebuild becomes set aggregation instead of re-running
        # every regex over every body. invalidate() intentionally does NOT clear this
        # cache -- membership moves do not change email content.
        self._email_cache: Dict[str, Dict[str, Any]] = {}
        self._email_cache_max = 20000
        self._group_sigs: Dict[str, str] = {}  # v1.1.0: per-group reuse signatures
        self._load_state()

    def update_settings(self, settings: Optional[Mapping[str, Any]]) -> None:
        self.settings = EngineSettings.from_mapping(settings)

    def invalidate(self) -> None:
        # v1.1.0: only the global signature is cleared. Fingerprints are kept so the
        # next build_fingerprints() can reuse every group whose per-group signature is
        # unchanged (see _group_signature) -- an auto-group during bulk refresh then
        # rebuilds one group instead of all of them.
        self._cache_signature = ""

    def _load_state(self) -> None:
        if not self.state_path:
            return
        try:
            with open(self.state_path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            if isinstance(data, dict):
                self.state.update(data)
        except Exception:
            pass

    def _save_state(self) -> None:
        if not self.state_path:
            return
        self.state["updated_at"] = datetime.now().isoformat()
        try:
            _atomic_json(self.state_path, self.state)
        except Exception:
            pass

    def _signature(self, groups_data: Mapping[str, Any], emails: Sequence[Mapping[str, Any]], quote_coverage: Mapping[str, Any]) -> str:
        groups = (groups_data.get("groups") or []) if isinstance(groups_data, Mapping) else []  # v1.2.0: None-safe
        memberships = (groups_data.get("memberships") or {}) if isinstance(groups_data, Mapping) else {}  # v1.2.0: None-safe
        q_updated = quote_coverage.get("updated_at", "") if isinstance(quote_coverage, Mapping) else ""
        payload = [str(len(groups)), str(len(memberships)), str(len(emails)), _text(q_updated)]
        for g in groups:
            if isinstance(g, Mapping):
                payload.extend([_text(g.get("id")), _text(g.get("name")), _text(g.get("archived")), _text(g.get("updated_at"))])
        # Membership content, not only length: moving one email must invalidate.
        # v1.1.0: order-independent crc32 sum instead of sort+join -- _signature runs on
        # every decide() call, and sorting/concatenating tens of thousands of membership
        # strings was measurable per email during bulk refresh. crc32 is deterministic
        # across processes (unlike hash()), and summing keeps it order-independent while
        # any single membership change still changes the total.
        acc = 0
        for k, v in memberships.items():
            acc = (acc + zlib.crc32((_text(k) + ">" + _text(v)).encode("utf-8", "ignore"))) & 0xFFFFFFFFFFFFFFFF
        payload.append(str(acc))
        return hashlib.sha256("|".join(payload).encode("utf-8", "ignore")).hexdigest()

    def _email_features_cached(self, e: Mapping[str, Any]) -> Dict[str, Any]:
        """v1.1.0: extract-once features for a stored email. Validity key covers the
        content lengths, timestamp and the material body limit so a changed body or a
        changed setting recomputes; anything else is a cache hit."""
        eid = _text(e.get("entry_id"))
        subject = _text(e.get("subject"))
        body = _text(e.get("body"))
        # v1.2.0: relationship identity is part of validity -- a Vendor->Customer
        # correction (or a resolved SMTP change) must recompute this email's features.
        _et = _text(e.get("email_type")).upper()
        _st = _text(e.get("sender_type")).lower()
        _snd = _text(e.get("resolved_sender_smtp") or e.get("sender_email") or e.get("sender")).lower()
        valid = (len(subject), len(body), _safe_timestamp(e.get("received_timestamp")), self.settings.material_body_limit, _et, _st, _snd)
        cached = self._email_cache.get(eid) if eid else None
        if cached is not None and cached.get("_valid") == valid:
            return cached
        sender_val = e.get("resolved_sender_smtp") or e.get("sender_email") or e.get("sender")
        et = _text(e.get("email_type")).upper()
        st = _text(e.get("sender_type")).lower()
        is_customer = et == "C" or "customer" in st
        feats: Dict[str, Any] = {
            "_valid": valid,
            "conv": _text(e.get("conversation_id")).strip().lower(),
            "imid": normalize_message_id(e.get("internet_message_id")),
            "refs": message_ids_from_header(e.get("in_reply_to_id")) | message_ids_from_header(e.get("references_ids")),
            "subj": normalize_subject(subject),
            "subj_words": words(subject),
            "mbx_refs": extract_mainbox_refs(f"{subject}\n{body}"),
            "business_ids": extract_business_ids(f"{subject}\n{body[:1800]}"),
            "sender_email": normalize_email(sender_val),
            "sender_domain": email_domain(sender_val),
            "is_customer": is_customer,
            "is_vendor": et == "V" or "vendor" in st or "supplier" in st,
            "customer_names": (extract_named_entities(subject, body, e.get("sender"))["customer_names"] if is_customer else set()),
            "materials": None,  # v1.1.0: parsed lazily -- most members beyond the per-group scan cap never need it
        }
        if eid:
            if len(self._email_cache) >= self._email_cache_max:
                for old in list(self._email_cache.keys())[: self._email_cache_max // 10]:
                    self._email_cache.pop(old, None)
            self._email_cache[eid] = feats
        return feats

    def _cached_materials(self, e: Mapping[str, Any], feats: Dict[str, Any]) -> List[Material]:
        """v1.1.0: parse materials on first request only, then reuse across rebuilds."""
        if feats.get("materials") is None:
            feats["materials"] = collect_email_materials(e, self.settings.material_body_limit)
        return feats["materials"]

    def _group_signature(self, g: Mapping[str, Any], member_crc: int, member_count: int, q_updated: str) -> str:
        """v1.1.0: per-group change signature. Covers group meta (name/notes feed the
        business-id extraction), the member set INCLUDING each member's subject/body
        lengths (bodies can be loaded lazily after import -- a late-arriving body must
        rebuild this group's fingerprint), and the coverage store timestamp (coverage
        rows feed materials). Any grouping move changes exactly the affected groups'
        signatures, so a rebuild after invalidate() reuses every untouched fingerprint
        wholesale -- keeping its prepped material rows and material_index."""
        blob = "|".join((
            _text(g.get("id")), _text(g.get("name")), _text(g.get("archived")),
            _text(g.get("updated_at")), _text(g.get("created_at")), _text(g.get("notes")),
            str(member_count), str(member_crc), _text(q_updated),
        ))
        return hashlib.sha256(blob.encode("utf-8", "ignore")).hexdigest()[:20]

    def build_fingerprints(self, groups_data: Mapping[str, Any], emails: Sequence[Mapping[str, Any]], quote_coverage: Optional[Mapping[str, Any]] = None) -> Dict[str, GroupFingerprint]:
        quote_coverage = quote_coverage or {}
        sig = self._signature(groups_data, emails, quote_coverage)
        if sig == self._cache_signature and self._fingerprints:
            return self._fingerprints
        memberships = (groups_data.get("memberships") or {}) if isinstance(groups_data, Mapping) else {}
        q_updated = _text(quote_coverage.get("updated_at", "")) if isinstance(quote_coverage, Mapping) else ""
        # v1.1.0: bucket members per group (newest first, exactly the old iteration
        # order) and accumulate a per-group content crc for the reuse signature.
        members_by_gid: Dict[str, List[Mapping[str, Any]]] = {}
        crc_by_gid: Dict[str, int] = {}
        for e in sorted((x for x in emails if isinstance(x, Mapping)), key=lambda x: _safe_timestamp(x.get("received_timestamp")), reverse=True):
            gid = _text(memberships.get(e.get("entry_id", ""))).strip()
            if not gid:
                continue
            members_by_gid.setdefault(gid, []).append(e)
            token = (f"{_text(e.get('entry_id'))}:{len(_text(e.get('subject')))}:{len(_text(e.get('body')))}:"
                     f"{_safe_timestamp(e.get('received_timestamp'))}:{_text(e.get('status'))}:{_text(e.get('urgency'))}:"
                     f"{_text(e.get('email_type'))}:{_text(e.get('sender_type'))}:"
                     f"{_text(e.get('resolved_sender_smtp') or e.get('sender_email') or e.get('sender'))}")  # v1.2.0: relationship identity rebuilds the group
            crc_by_gid[gid] = (crc_by_gid.get(gid, 0) + zlib.crc32(token.encode("utf-8", "ignore"))) & 0xFFFFFFFFFFFFFFFF
        # Coverage recs routed to their group up front (identical resolution as before).
        threads = quote_coverage.get("threads", {}) if isinstance(quote_coverage, Mapping) else {}
        cov_by_gid: Dict[str, List[Mapping[str, Any]]] = {}
        if isinstance(threads, Mapping):
            for key, rec in threads.items():
                if not isinstance(rec, Mapping):
                    continue
                gid = ""
                if _text(key).startswith("grp:"):
                    gid = _text(key)[4:]
                gid = _text(rec.get("group_id") or gid)
                if gid:
                    cov_by_gid.setdefault(gid, []).append(rec)
        old_fps = self._fingerprints
        old_sigs = getattr(self, "_group_sigs", {})
        fps: Dict[str, GroupFingerprint] = {}
        new_sigs: Dict[str, str] = {}
        _MAX_SUBJECTS = 30        # v1.1.0: newest-first cap; SequenceMatcher runs per subject in scoring
        _MAX_CUSTOMER_NAMES = 25  # v1.1.0: bounds the name cross-product in scoring
        _MAX_GROUP_MATERIALS = 400  # v1.1.0: bounds per-candidate material scoring work
        for g in (groups_data.get("groups") or []) if isinstance(groups_data, Mapping) else []:  # v1.2.0: None-safe
            if not isinstance(g, Mapping):
                continue
            gid = _text(g.get("id")).strip()
            if not gid:
                continue
            gsig = self._group_signature(g, crc_by_gid.get(gid, 0), len(members_by_gid.get(gid, ())), q_updated)
            reused = old_fps.get(gid)
            if reused is not None and old_sigs.get(gid) == gsig:
                fps[gid] = reused
                new_sigs[gid] = gsig
                continue
            fp = GroupFingerprint(group_id=gid, name=_text(g.get("name")), archived=bool(g.get("archived")), created_at=_text(g.get("created_at")))
            combined = f"{fp.name}\n{_text(g.get('notes'))}"
            fp.business_ids = extract_business_ids(combined)
            fp.loose_tokens |= fp.business_ids.get("loose", set())
            fp.subjects.add(normalize_subject(fp.name))
            fp.subject_words |= words(fp.name)
            fp.mainbox_refs |= extract_mainbox_refs(combined)
            count = 0
            for e in members_by_gid.get(gid, ()):
                feats = self._email_features_cached(e)  # v1.1.0: extract once, reuse on every rebuild
                fp.member_ids.add(_text(e.get("entry_id")))
                if feats["conv"]: fp.conversation_ids.add(feats["conv"])
                if feats["imid"]: fp.message_ids.add(feats["imid"])
                fp.references |= feats["refs"]
                subj = feats["subj"]
                if subj and len(fp.subjects) < _MAX_SUBJECTS:
                    fp.subjects.add(subj)
                fp.subject_words |= feats["subj_words"]
                fp.mainbox_refs |= feats["mbx_refs"]
                for kind, vals in feats["business_ids"].items():
                    fp.business_ids.setdefault(kind, set()).update(vals)
                fp.loose_tokens |= feats["business_ids"].get("loose", set())
                em = feats["sender_email"]; dom = feats["sender_domain"]
                if feats["is_customer"]:
                    if em: fp.customer_emails.add(em)
                    if dom and dom not in PUBLIC_DOMAINS: fp.customer_domains.add(dom)
                    if len(fp.customer_names) < _MAX_CUSTOMER_NAMES:
                        fp.customer_names |= feats["customer_names"]
                elif feats["is_vendor"] and dom:
                    fp.vendor_domains.add(dom)
                ts = _safe_timestamp(e.get("received_timestamp"))
                fp.last_activity = max(fp.last_activity, ts)
                if _text(e.get("status")) != "Completed" and _text(e.get("urgency")) != "Low":
                    fp.active_count += 1
                count += 1
                if count <= self.settings.max_group_member_scan:
                    fp.materials.extend(self._cached_materials(e, feats))  # v1.1.0: parsed once, cached
            # Quote Coverage is the most authoritative material/customer source.
            for rec in cov_by_gid.get(gid, ()):
                ce = normalize_email(rec.get("customer_email"))
                cd = email_domain(ce)
                if ce: fp.customer_emails.add(ce)
                if cd and cd not in PUBLIC_DOMAINS: fp.customer_domains.add(cd)
                cn = _text(rec.get("customer_name")).strip().lower()
                if cn: fp.customer_names.add(cn)
                for row in list(rec.get("requested", []) or []) + list(rec.get("quoted", []) or []):
                    if isinstance(row, Mapping):
                        m = material_from_mapping(row)
                        if m: fp.materials.append(m)
                conv = _text(rec.get("conversation_id")).strip().lower()
                if conv: fp.conversation_ids.add(conv)
                fp.mainbox_refs |= extract_mainbox_refs(json.dumps(rec, default=str)[:12000])
            fp.materials = _dedupe_materials(fp.materials)
            if len(fp.materials) > _MAX_GROUP_MATERIALS:
                # Keep explicit-part rows (the strongest, cheapest evidence) first, then
                # the earliest-collected description rows (member loop is newest-first,
                # so these come from the most recent activity).
                parted = [m for m in fp.materials if m.part]
                partless = [m for m in fp.materials if not m.part]
                fp.materials = (parted + partless)[:_MAX_GROUP_MATERIALS]
            # NOTE: rows are prepped lazily inside material_set_similarity; because the
            # Material objects persist in the cache/fingerprint, prep still runs once.
            fps[gid] = fp
            new_sigs[gid] = gsig
        self._fingerprints = fps
        self._group_sigs = new_sigs
        self._cache_signature = sig
        return fps

    def _score_candidate(self, email: Mapping[str, Any], fp: GroupFingerprint, email_features: Mapping[str, Any], sm_budget: Optional[List[int]] = None) -> CandidateScore:
        c = CandidateScore(group_id=fp.group_id, group_name=fp.name, score=0.0, confidence=0.0)
        subject = email_features["subject"]
        conv = email_features["conversation_id"]
        imid_refs: Set[str] = email_features["reply_refs"]
        refs: Set[str] = email_features["mainbox_refs"]
        ids: Dict[str, Set[str]] = email_features["business_ids"]
        sender_email = email_features["sender_email"]
        sender_domain = email_features["sender_domain"]
        is_customer = email_features["is_customer"]
        is_vendor = email_features["is_vendor"]

        # Deterministic linkage.
        if refs and fp.mainbox_refs and refs & fp.mainbox_refs:
            c.score += 1000; c.families.add("lineage"); c.reasons.append("Exact MaINbox reference matched")
            c.details["mainbox_refs"] = sorted(refs & fp.mainbox_refs)
        if imid_refs and fp.message_ids and imid_refs & fp.message_ids:
            c.score += 970; c.families.add("lineage"); c.reasons.append("Reply header points to a message in this group")
        if conv and conv in fp.conversation_ids:
            c.score += 930; c.families.add("lineage"); c.reasons.append("Outlook conversation matched")

        # v1.2.0: a MaINbox reference belonging to a DIFFERENT group is deterministic
        # contrary evidence (our own durable ref) and stays a hard veto even with a
        # conversation match -- forwarded RFQs can share a thread across jobs.
        if refs and fp.mainbox_refs and not (refs & fp.mainbox_refs):
            c.hard_veto = True
            c.conflicts.append("MaINbox reference points to a different group")

        # Hard business identities. Different explicit IDs of the same kind veto the
        # candidate -- UNLESS deterministic lineage (MaINbox ref / reply header /
        # Outlook conversation) already ties the email to this group. v1.2.0: a new
        # PO number arriving inside the same quote conversation is the normal customer
        # workflow; the email stays in its conversation's group and the conflict
        # becomes a visible review note plus a moderate penalty instead of leaving
        # the email ungrouped. (See the identifier-label fix in extract_business_ids:
        # most of these conflicts were manufactured by swallowed labels anyway.)
        _has_lineage = "lineage" in c.families
        id_hits: List[str] = []
        _matched_id_kinds: Set[str] = set()
        _conflict_under_lineage = False
        for kind in ("po", "job", "rfq", "quote", "order", "ref"):
            ev = set(ids.get(kind, set()))
            gv = set(fp.business_ids.get(kind, set()))
            if ev and gv:
                shared = ev & gv
                if shared:
                    _matched_id_kinds.add(kind)
                    weight = 760 if kind in {"po", "job", "rfq"} else 620
                    c.score += weight + min(60, 15 * len(shared)); c.families.add("business_id")
                    c.details["business_id_labeled"] = True  # v1.2.1: a LABELED id (JOB/PO/RFQ/QUOTE/ORDER/REF) is real corroboration
                    id_hits.extend(f"{kind.upper()} {x}" for x in sorted(shared))
                elif _has_lineage:
                    # v1.2.0 policy split: PO/QUOTE/ORDER/REF/RFQ are per-transaction
                    # numbers -- a new one inside the quote conversation is the normal
                    # customer workflow, so lineage keeps the email grouped with a
                    # review note. A conflicting JOB number is different: customers
                    # reply to an OLD thread to start a NEW job (thread reuse), so a
                    # conversation must not auto-carry the email across an explicit
                    # job change -- it becomes a SUGGESTION for human confirmation.
                    # Exception: an exact MaINbox reference is our own per-job durable
                    # ref and outranks a typed job number.
                    _mbx_matched = bool(c.details.get("mainbox_refs"))
                    _conflict_under_lineage = True
                    if kind == "job" and not _mbx_matched:
                        c.score -= 400
                        c.details["job_conflict_review"] = True
                    else:
                        c.score -= 150
                    _note = (f"New/conflicting {kind.upper()} detected in this conversation "
                             f"({', '.join(sorted(ev))} vs {', '.join(sorted(gv))}) -- review job metadata")
                    c.conflicts.append(_note)
                    c.reasons.append(_note)
                    c.details["id_conflict_review"] = True
                else:
                    c.hard_veto = True
                    c.conflicts.append(f"Explicit {kind.upper()} conflicts ({', '.join(sorted(ev))} vs {', '.join(sorted(gv))})")
        # v1.2.0 corroboration rule: lineage overrides an id conflict to full AUTO
        # only when the email ALSO agrees with this group on an identifier of a
        # DIFFERENT kind (the auditor's case: Job matched while a new PO arrived in
        # the thread) or the MaINbox reference matched. A conversation carrying a
        # conflicting id with NO agreeing identifier is the thread-reuse pattern --
        # keep the lineage suggestion but require a human confirmation.
        if _conflict_under_lineage and not (_matched_id_kinds or c.details.get("mainbox_refs")):
            c.details["id_conflict_review_suggest"] = True
        loose_hits = set(ids.get("loose", set())) & fp.loose_tokens
        if loose_hits and not c.hard_veto:
            # v1.2.1: loose (unlabeled) tokens still add confidence, but they routinely
            # capture manufacturer part numbers and street-address fragments -- the same
            # evidence the materials family represents -- so a business_id built ONLY
            # from loose tokens does NOT satisfy the two-family independence gate (see
            # decide()). business_id_labeled is left unset here on purpose.
            c.score += min(330, 180 + 35 * len(loose_hits)); c.families.add("business_id")
            id_hits.extend(sorted(loose_hits))
        if id_hits:
            c.reasons.append("Business identifiers matched: " + ", ".join(id_hits[:5]))
            c.details["business_hits"] = id_hits

        # Customer identity. Vendor domain is deliberately not job evidence.
        if is_customer:
            if sender_email and sender_email in fp.customer_emails:
                c.score += 230; c.families.add("customer"); c.reasons.append("Exact customer contact matched")
            elif sender_domain and sender_domain not in PUBLIC_DOMAINS and sender_domain in fp.customer_domains:
                c.score += 175; c.families.add("customer"); c.reasons.append("Customer company domain matched")
            elif sender_domain and fp.customer_domains and sender_domain not in PUBLIC_DOMAINS:
                # Only a hard conflict when both sides are confirmed customer identities and no lineage.
                if "lineage" not in c.families:
                    c.conflicts.append("Confirmed customer domain differs")
                    c.score -= 220
        elif is_vendor and sender_domain and sender_domain in fp.vendor_domains:
            c.score += 18  # role/history hint only, never an evidence family
            c.reasons.append("Vendor has previously participated in this group (weak hint)")

        ent_names = email_features["customer_names"]
        if ent_names and fp.customer_names:
            best_name = max((SequenceMatcher(None, a, b).ratio() for a in ent_names for b in fp.customer_names), default=0.0)
            if best_name >= 0.88:
                c.score += 150; c.families.add("customer"); c.reasons.append("Customer name matched")

        # Material multiset similarity.
        mats = email_features["materials"]
        if self.settings.use_materials and mats and fp.materials and not c.hard_veto:
            if fp.material_index is None:
                fp.material_index = build_material_index(fp.materials)  # v1.1.0: built once per fingerprint
            ms = material_set_similarity(mats, fp.materials, sm_budget, fp.material_index)  # v1.1.0: shared per-decide SequenceMatcher budget
            c.details["materials"] = ms
            if ms["score"] >= 0.35:
                pts = 410 * ms["score"]
                # One common line is not enough by itself; cap unless exact/distinctive coverage is stronger.
                if ms["matches"] == 1 and ms["exact_parts"] < 1:
                    pts = min(pts, 95)
                c.score += pts; c.families.add("materials")
                c.reasons.append(f"Materials matched {ms['matches']}/{len(mats)} lines ({ms['coverage']:.0%} coverage, {ms['exact_parts']} exact parts)")

        # Subject/job wording is supporting evidence only.
        sw = email_features["subject_words"]
        inter = sw & fp.subject_words
        if sw and fp.subject_words:
            jac = len(inter) / max(1, len(sw | fp.subject_words))
            seq = max((SequenceMatcher(None, subject, s).ratio() for s in fp.subjects if s), default=0.0)
            subj_sim = 0.65 * jac + 0.35 * seq
            if subj_sim >= 0.28:
                c.score += min(120, 130 * subj_sim); c.families.add("subject")
                c.reasons.append(f"Subject/job wording similarity {subj_sim:.0%}")
                c.details["subject_similarity"] = subj_sim

        # Existing MaINbox learning: only when the learned rule points to this group and business evidence matches.
        learned = email_features.get("learned_group_ids", set())
        if fp.group_id in learned and ("business_id" in c.families or "lineage" in c.families):
            c.score += 130; c.families.add("learning"); c.reasons.append("Prior user-approved grouping pattern matched")

        # Plug-in feedback. A prior correction away from this group is a strong veto for the same signature.
        feedback_key = email_features["feedback_key"]
        rejected = self.state.get("rejected", {}).get(feedback_key, {}) if isinstance(self.state.get("rejected"), dict) else {}
        if isinstance(rejected, Mapping) and fp.group_id in rejected:
            c.hard_veto = True; c.conflicts.append("User previously rejected this grouping pattern")
        accepted = self.state.get("accepted", {}).get(feedback_key, {}) if isinstance(self.state.get("accepted"), dict) else {}
        if isinstance(accepted, Mapping) and fp.group_id in accepted and not c.hard_veto:
            c.score += min(180, 90 + 15 * int(accepted.get(fp.group_id, 1) or 1)); c.families.add("learning")
            c.reasons.append("User previously confirmed this grouping pattern")

        # Lifecycle/recent activity is a tie-breaker, never a core family.
        if fp.active_count:
            c.score += min(25, fp.active_count * 3)
        if fp.archived and "lineage" not in c.families:
            c.score -= 130; c.conflicts.append("Group is archived")

        if c.hard_veto:
            c.score = min(c.score, 0.0)
        # Smooth confidence; score 700≈0.80, 900≈0.92. Deterministic links clamp high.
        if "lineage" in c.families and c.score >= 900:
            c.confidence = 0.995
        else:
            c.confidence = max(0.0, min(0.99, 1.0 / (1.0 + math.exp(-(c.score - 520.0) / 145.0))))
        return c

    def _feedback_signature(self, email: Mapping[str, Any]) -> str:
        sender = normalize_email(email.get("resolved_sender_smtp") or email.get("sender_email") or email.get("sender"))
        ids = extract_business_ids(f"{email.get('subject','')}\n{_text(email.get('body'))[:1200]}")
        id_blob = ",".join(sorted(set().union(*(v for k, v in ids.items() if k != "loose"))))
        subj_words = ",".join(sorted(words(email.get("subject"))))
        raw = f"{sender}|{id_blob}|{subj_words}"
        return hashlib.sha256(raw.encode("utf-8", "ignore")).hexdigest()[:24]

    def decide(self, email: Mapping[str, Any], groups_data: Mapping[str, Any], emails: Sequence[Mapping[str, Any]], quote_coverage: Optional[Mapping[str, Any]] = None, learning_data: Optional[Mapping[str, Any]] = None) -> Dict[str, Any]:
        if not self.settings.enabled or not isinstance(email, Mapping):
            return GroupDecision(reason="Auto-group engine disabled").to_dict()
        fps = self.build_fingerprints(groups_data, emails, quote_coverage or {})
        if not fps:
            return GroupDecision(reason="No active groups available").to_dict()
        entry_id = _text(email.get("entry_id"))
        memberships = groups_data.get("memberships", {}) if isinstance(groups_data, Mapping) else {}
        if entry_id and memberships.get(entry_id):
            return GroupDecision(reason="Email is already grouped").to_dict()
        sender_val = email.get("resolved_sender_smtp") or email.get("sender_email") or email.get("sender")
        et = _text(email.get("email_type")).upper()
        st = _text(email.get("sender_type")).lower()
        ids = extract_business_ids(f"{email.get('subject','')}\n{_text(email.get('body'))[:3000]}")
        learned_ids: Set[str] = set()
        if isinstance(learning_data, Mapping):
            bucket = learning_data.get("group_rules", {})
            if isinstance(bucket, Mapping):
                for rule in bucket.values():
                    if isinstance(rule, Mapping):
                        gid = _text(rule.get("group_id"))
                        rule_tokens = set(rule.get("business_tokens") or [])
                        if gid and (not rule_tokens or rule_tokens & set(ids.get("loose", set()))):
                            learned_ids.add(gid)
        ents = extract_named_entities(email.get("subject"), email.get("body"), email.get("sender"))
        features = {
            "subject": normalize_subject(email.get("subject")),
            "subject_words": words(email.get("subject")),
            "conversation_id": _text(email.get("conversation_id")).strip().lower(),
            "reply_refs": message_ids_from_header(email.get("in_reply_to_id")) | message_ids_from_header(email.get("references_ids")),
            "mainbox_refs": extract_mainbox_refs(f"{email.get('subject','')}\n{email.get('body','')}"),
            "business_ids": ids,
            "sender_email": normalize_email(sender_val),
            "sender_domain": email_domain(sender_val),
            "is_customer": et == "C" or "customer" in st,
            "is_vendor": et == "V" or "vendor" in st or "supplier" in st,
            "customer_names": ents["customer_names"],
            "job_names": ents["job_names"],
            "materials": collect_email_materials(email, self.settings.material_body_limit),
            "learned_group_ids": learned_ids,
            "feedback_key": self._feedback_signature(email),
        }
        # v1.1.0: one SequenceMatcher budget per decide() bounds worst-case latency on
        # inboxes where common commodity descriptions overlap across many groups. Exact
        # scoring is unaffected until the budget is spent; beyond it, description
        # similarity falls back to its Jaccard lower bound (see _description_similarity_prepped).
        _sm_budget = [4000]
        scored = [self._score_candidate(email, fp, features, _sm_budget) for fp in fps.values() if not fp.archived or features["conversation_id"] in fp.conversation_ids]
        scored.sort(key=lambda x: (x.score, x.confidence, len(x.families)), reverse=True)
        valid = [x for x in scored if not x.hard_veto and x.score > 0]
        if not valid:
            candidates = [x.to_dict() for x in scored[: self.settings.max_candidates]]
            return GroupDecision(reason="No candidate passed conflict and confidence checks", candidates=candidates).to_dict()
        best = valid[0]
        second_conf = valid[1].confidence if len(valid) > 1 else 0.0
        margin = best.confidence - second_conf
        deterministic = ("lineage" in best.families and best.confidence >= 0.99
                         and not best.details.get("job_conflict_review")
                         and not best.details.get("id_conflict_review_suggest"))  # v1.2.0: thread-reuse guards
        # v1.2.1: a business_id family built ONLY from loose (unlabeled) tokens is not
        # independent corroboration. Loose tokens routinely capture manufacturer part
        # numbers (Hubbell "HBL5369C") and street-address fragments (the Astoria address
        # "31-03" -> "3103") -- the SAME evidence the materials family already carries.
        # Counting it as a second family let two different GCs at one building, quoting
        # the same commodity devices, AUTO-group. It still contributes confidence; it
        # just cannot be the second independent family. A LABELED id match
        # (business_id_labeled) is real corroboration and still counts.
        _independent = set(best.families) - {"subject", "learning"}
        if "business_id" in _independent and not best.details.get("business_id_labeled"):
            _independent.discard("business_id")
        enough_families = len(_independent) >= 2 or deterministic
        action = "none"
        if deterministic and self.settings.deterministic_lineage_auto:
            action = "auto"
        elif best.confidence >= self.settings.auto_threshold and margin >= self.settings.min_margin and (enough_families or not self.settings.require_two_families):
            # Materials alone must be unusually strong and contain at least two exact parts.
            if best.families == {"materials"}:
                mdet = best.details.get("materials", {})
                if mdet.get("exact_parts", 0) >= 2 and mdet.get("coverage", 0) >= 0.75:
                    action = "auto"
                else:
                    action = "suggest"
            else:
                action = "auto"
        elif best.confidence >= self.settings.suggest_threshold:
            action = "suggest"
        reason_parts = list(best.reasons)
        if best.conflicts:
            reason_parts.append("Warnings: " + "; ".join(best.conflicts))
        if action != "auto" and margin < self.settings.min_margin and len(valid) > 1:
            reason_parts.append(f"Winner margin is only {margin:.0%}")
        if action == "auto" and self.settings.require_two_families and not enough_families:
            action = "suggest"
        if action == "auto" and (best.details.get("job_conflict_review") or best.details.get("id_conflict_review_suggest")):
            action = "suggest"  # v1.2.0: an uncorroborated id conflict in the thread always gets a human look
        return GroupDecision(
            action=action,
            group_id=best.group_id,
            group_name=best.group_name,
            confidence=best.confidence,
            score=best.score,
            margin=margin,
            reason=" | ".join(reason_parts) or "Best candidate",
            evidence_families=sorted(best.families),
            candidates=[x.to_dict() for x in valid[: self.settings.max_candidates]],
        ).to_dict()

    def record_feedback(self, email: Mapping[str, Any], target_group_id: str, previous_group_id: str = "", groups_data: Optional[Mapping[str, Any]] = None, reason: str = "manual move") -> None:
        if not isinstance(email, Mapping) or not target_group_id:
            return
        key = self._feedback_signature(email)
        accepted = self.state.setdefault("accepted", {}).setdefault(key, {})
        accepted[target_group_id] = int(accepted.get(target_group_id, 0) or 0) + 1
        if previous_group_id and previous_group_id != target_group_id:
            rejected = self.state.setdefault("rejected", {}).setdefault(key, {})
            rejected[previous_group_id] = {"at": datetime.now().isoformat(), "reason": reason}
        # Keep state bounded.
        for bucket_name in ("accepted", "rejected"):
            bucket = self.state.get(bucket_name, {})
            if isinstance(bucket, dict) and len(bucket) > 3000:
                for old_key in list(bucket.keys())[: len(bucket) - 3000]:
                    bucket.pop(old_key, None)
        self._save_state()


def decide_group(email: Mapping[str, Any], groups_data: Mapping[str, Any], emails: Sequence[Mapping[str, Any]], quote_coverage: Optional[Mapping[str, Any]] = None, learning_data: Optional[Mapping[str, Any]] = None, settings: Optional[Mapping[str, Any]] = None, state_path: str = "") -> Dict[str, Any]:
    """Stateless convenience wrapper."""
    return AutoGroupEngine(settings=settings, state_path=state_path).decide(email, groups_data, emails, quote_coverage, learning_data)
