"""Material-line classifier driven by material_rules.json (v1.2, post-verification + patch7).

    from material_classify import classify
    classify("1/2\" EMT compression coupling") -> ("fittings_emt", 0.85, ["kw:fittings_emt:...", ...])

Resolution: (0) the line is normalised - unicode quotes/fractions, an "availability for ..." / "please quote"
prefix is stripped, quantities glued to text ("STRAP100ea") are split, and vendor abbreviations (MILB, HUBW,
DPLX, RCP, GTTR ...) are expanded; (1) '^'-anchored keywords are catalog-prefix OVERRIDES and win outright;
(2) categories are tried in descending priority, a category whose `negative` regex matches the whole line is
vetoed; (3) brand aliases found in the line add a hint - they boost an agreeing keyword match, break near-ties
in favour of the brand's categories, and supply the category when no keyword fired (single-line brands score
0.4, multi-line brands 0.35, +0.2 when a part-number-like token is present); (4) with neither, a coarse group
guess is returned as "group:<group>" with low score.

Run as a script: fills material_rules.json['vendor_map'] from vendor_evidence.json (applying
rules['vendor_overrides']: customer removal, domain merges, manual vendors, per-vendor category fixes) and
prints the hard-example self-check.
"""
import json, re, os, sys, collections
HERE = os.path.dirname(os.path.abspath(__file__))
RULES_PATH = os.path.join(HERE, "material_rules.json")
_R = json.load(open(RULES_PATH, encoding="utf-8"))

NORM = str.maketrans({"“": '"', "”": '"', "″": '"', "‘": "'", "’": "'", "′": "'",
                      "½": "1/2", "¾": "3/4", "¼": "1/4", "⅜": "3/8", "⅝": "5/8", "–": "-", "—": "-", " ": " ", "\t": " "})
CATS = sorted(_R["categories"], key=lambda c: -c["priority"])
CAT_BY_ID = {c["id"]: c for c in CATS}
GROUP_OF = {c["id"]: c["group"] for c in CATS}
_OVR, _KW, _NEG = [], {}, {}
for c in CATS:
    _KW[c["id"]] = []
    for k in c["keywords"]:
        rx = re.compile(k, re.I)
        (_OVR.append((rx, c["id"])) if k.startswith("^") else _KW[c["id"]].append(rx))
    _NEG[c["id"]] = [re.compile(n, re.I) for n in c["negative"]]
BRANDS = _R["brands"]
ABBR = _R.get("abbreviations", {})
_ABBR_RX = re.compile(r"(?<![A-Za-z0-9])(" + "|".join(sorted(map(re.escape, ABBR), key=len, reverse=True)) + r")(?![A-Za-z0-9])") if ABBR else None
# aliases that collide with ordinary product words ("global", "ideal", "bell" ...) get half weight
WEAK = {"global", "bell", "ge", "gb", "keystone", "halo", "cooper", "essex", "champion", "genesis", "orbit", "ideal", "republic", "mason",
        "service wire", "current lighting", "juno", "carol", "scotch", "satco", "chief", "morris", "andrew", "rib", "sti", "pip", "woods", "o-z", "fre"}
def _is_weak(alias): return alias in WEAK
def _brand_pattern(a):
    core = re.escape(a).replace(r"\ ", r"\s*").replace(r"\-", r"[\s-]?")
    if len(re.sub(r"[^a-z0-9]", "", a)) <= 3:      # short aliases (3m, ge, gb, t&b): no '-' / '/' neighbours -> not inside part numbers
        return re.compile(r"(?<![a-z0-9/#.-])" + core + r"(?![a-z0-9/-])", re.I)
    return re.compile(r"(?<![a-z0-9])" + core + r"(?![a-z0-9])", re.I)
_BRAND_RX = [(_brand_pattern(a), a) for a in sorted(BRANDS, key=len, reverse=True)]
_PN_RX = re.compile(r"(?<![a-z])(?=[a-z0-9#-]*\d)(?=[a-z0-9#-]*[a-z])[a-z0-9#-]{4,}|\b\d{4,}[a-z]?\b", re.I)   # part-number-like token
_PREFIX = re.compile(r"^\s*(?:(?:price\s*(?:and|&)\s*)?availability\s*(?:for|on|of)?\s*(?:a|an|the)?\s*|(?:please\s*)?(?:quote|p&a|pricing)\s*(?:on|for|me)?\s*(?:the)?\s*)", re.I)
_GLUED_QTY = re.compile(r"(?<=[A-Za-z\"])(\d{1,5})\s*(ea|pcs?|pk|ft|rl|bx|cs)\b", re.I)
GROUP_GUESS = [
 ("wire_cable", re.compile(r"\b(?:wire|cable|cbl|awg|conductors?|cu\b|copper|alum(?:inum)?|feet|ft\b|reel|spool)\b|\d{1,2}/\d\b|\d+'\s*(?:of\b)?", re.I)),
 ("raceway", re.compile(r"\b(?:conduit|pipe|emt|rigid|imc|pvc|raceway|strut|channel|tray|10'|20')\b", re.I)),
 ("fittings", re.compile(r"\b(?:fitting|conn(?:ector)?s?|couplings?|cplg|adapters?|bushings?|locknuts?|nipples?|hubs?)\b", re.I)),
 ("boxes_enclosures", re.compile(r"\b(?:box(?:es)?|enclosure|cabinet|trough|gutter|wireway|cover|plaster\s*ring|mud\s*ring)\b", re.I)),
 ("gear", re.compile(r"\b(?:breaker|panel|panelboard|loadcenter|disconnect|switchgear|transformer|fuse|amp|kva|contactor|starter|meter)\b", re.I)),
 ("lighting", re.compile(r"\b(?:light|lighting|fixture|lamp|led|lumen|bulb|troffer|highbay|dimmer|sensor|photocell)\b", re.I)),
 ("devices", re.compile(r"\b(?:receptacle|outlet|switch|plate|cover\s*plate|gfci|decora|toggle|plug|dimmer)\b", re.I)),
 ("connectors_grounding", re.compile(r"\b(?:lug|terminal|splice|crimp|ground(?:ing)?|gnd|bond(?:ing)?|wire\s*nut|tie)s?\b", re.I)),
 ("hardware_misc", re.compile(r"\b(?:anchor|screw|bolt|nut|washer|strap|clip|clamp|hanger|tape|tool|drill|bit|blade|glove|marker|label)s?\b", re.I)),
]


def normalize(text):
    s = str(text).translate(NORM).strip()
    s = _PREFIX.sub("", s, count=1)
    s = _GLUED_QTY.sub(r" \1 \2", s)
    if _ABBR_RX: s = _ABBR_RX.sub(lambda m: ABBR[m.group(1)], s)
    return s.strip()


def _brand_hits(s):
    hits = [a for rx, a in _BRAND_RX if rx.search(s)]
    return [a for a in hits if not any(a != b and a in b for b in hits)]   # longest non-overlapping ("cooper lighting" beats "cooper")


def classify(text):
    """Return (category_id | 'group:<g>' | None, score 0..1, reasons list)."""
    if not text:
        return None, 0.0, ["empty"]
    s = normalize(text)
    if not s:
        return None, 0.0, ["empty"]
    reasons = []
    brands = _brand_hits(s)
    brand_cats, brand_w = collections.OrderedDict(), 0.0
    for a in brands:
        b = BRANDS[a]
        w = 0.5 if _is_weak(a) else 1.0
        brand_w = max(brand_w, w)
        reasons.append(f"brand:{a}->{b['manufacturer']}({w})")
        for c in b["categories"]:
            brand_cats[c] = max(brand_cats.get(c, 0), w)
    # 1) overrides
    for rx, cid in _OVR:
        if rx.search(s):
            reasons.append(f"override:{cid}:{rx.pattern[:40]}")
            return cid, 1.0, reasons
    # 2) priority scan; collect all matching categories
    matches = []
    for c in CATS:
        cid = c["id"]
        veto = next((n for n in _NEG[cid] if n.search(s)), None)
        if veto:
            reasons.append(f"veto:{cid}")
            continue
        hit = next((rx for rx in _KW[cid] if rx.search(s)), None)
        if hit:
            matches.append((cid, c["priority"], hit.pattern))
    if matches:
        top = matches[0]
        chosen = top
        if brand_cats and top[0] not in brand_cats:
            for m in matches[1:]:
                if m[0] in brand_cats and brand_cats[m[0]] >= 1.0 and top[1] - m[1] <= 6:
                    chosen = m
                    reasons.append(f"brand-tiebreak:{top[0]}->{m[0]}")
                    break
        cid = chosen[0]
        score = 0.7
        if len(matches) == 1: score += 0.1
        if cid in brand_cats: score += 0.2 * brand_cats[cid]
        elif brand_cats: score -= 0.1
        reasons.append(f"kw:{cid}:{chosen[2][:60]}")
        if len(matches) > 1:
            reasons.append("also:" + ",".join(m[0] for m in matches if m is not chosen))
        return cid, round(min(score, 1.0), 2), reasons
    # 3) brand only (a category vetoed by its own negatives is not a valid brand fallback either)
    brand_cats = collections.OrderedDict((c, w) for c, w in brand_cats.items() if not any(n.search(s) for n in _NEG.get(c, [])))
    if brand_cats:
        cid = next(iter(brand_cats))
        a0 = brands[0]
        base = 0.4 if len(BRANDS[a0]["categories"]) <= 2 else 0.35
        score = base * brand_cats[cid]
        if _PN_RX.search(re.sub(re.escape(a0), "", s, flags=re.I)):
            score += 0.2; reasons.append("brand+partno")
        reasons.append(f"brand-only:{cid}")
        return cid, round(min(score, 1.0), 2), reasons
    # 4) group guess
    for g, rx in GROUP_GUESS:
        if rx.search(s):
            reasons.append(f"group-guess:{g}:{rx.pattern[:40]}")
            return "group:" + g, 0.2, reasons
    return None, 0.0, reasons + ["no-match"]


# noise filter for raw RFQ lines (addresses, chatter, price fragments, signatures)
DROP = re.compile(r"(^\d+[\w\-]*\s+(?:[A-Za-z0-9.']+\s+){0,4}(?:street|st\.?|ave\.?|avenue|blvd\.?|road|rd\.?|drive|dr\.?|lane|ln\.?|way|place|pl\.?|suite|floor|parkway|pkwy|highway|hwy|turnpike|tpke|court|ct\.?|boulevard)\b|tel:|^\d{4}-\d{2}-\d{2}|\d{3}[-.]\d{3}[-.]\d{4}|\bp\.?o\.?\s*box\b|\b[A-Z]{2}\s+\d{5}\b|rights reserved|\bquotes?\b|\bpricing\b|\bprices?\b|\bportal\b|\bexpire|\bsubject to\b|\bholiday\b|\bdeliver(?:y|ies)\b|\bfreight\b|\bterms\b|\b(?:we|you|your|this|that|can|will|is|are|have|need|know|let|see|per|our|it|was|not|please|thanks|thank|hi|hello|regards|best)\b|^[\d.,\s$/-]+$|^[\d.]+\s*in\s+\d+$|^\.?\d+in\s+\d+$"
                  r"|https?://|<[^>]+@[^>]+>|\.(?:jpe?g|png|gif|pdf|xlsx?|docx?)\b|\btariff\b|\bcost\s*adjustments?\b|\border\s*qty\b|\bupc\s*description\b|\bpay\s*by\b|\bdiscount\b|\binvoice\b|\bremit\b|\bunsubscribe\b|\bconfidential\b|\bsincerely\b|\bcell:|\bfax:|\bphone:|\bmobile:|\bwww\.|\.com\b|\bbid\s*invite\b|\bjob\s*:|\bproject\s*:|\bre:|\bfw:|\bfwd:|\battached\b|\bbelow\b|\bequal\s*to\b|\bsame\s*specs\b|\bstock\b\s*\?|\bin\s*stock\b|\bmust\s*be\b|\bor\s*equal\b|^\s*(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+\d{1,2}|\bemployee\b|\bowned\b|\bramland\b|^o\s*:\s*x\d+|\bx\d{3,4}\b\s*$)", re.I)


def is_noise(line):
    return not line or len(line.strip()) < 3 or bool(DROP.search(line))


def _dedupe_key(line):
    return re.sub(r"[^a-z0-9]+", " ", normalize(line).lower()).strip()


def build_vendor_map(vendor_evidence_path):
    V = json.load(open(vendor_evidence_path, encoding="utf-8"))
    ov = _R.get("vendor_overrides", {})
    customers = set(ov.get("customers_not_vendors", [])); merge = ov.get("merge", {}); cat_ov = ov.get("category_overrides", {})
    # gather evidence per (merged) domain
    ev = collections.OrderedDict()
    for conf, key in (("strong", "vendors"), ("weak", "weak_vendors")):
        for v in V.get(key, []):
            d = v["domain"]
            if d in customers: continue
            tgt = merge.get(d, d)
            e = ev.setdefault(tgt, {"name": v["name"], "rfq_count": 0, "confidence": conf, "lines": [], "merged_from": []})
            if tgt != d: e["merged_from"].append(d)
            e["rfq_count"] += v["rfq_count"]; e["lines"] += v.get("item_lines", [])
            if conf == "strong": e["confidence"] = "strong"
    vmap, samples = {}, []
    for dom, e in ev.items():
        counts = collections.Counter(); brand_support = set(); seen = set()
        dkey = re.sub(r"[^a-z0-9]", "", (dom.split("@")[-1].split(".")[0] + " " + e["name"]).lower())
        domain_brand = []
        for a, b in BRANDS.items():
            k = re.sub(r"[^a-z0-9]", "", a)
            if (len(k) >= 4 or k == dkey.split(" ")[0]) and k in dkey:
                domain_brand.append(a); brand_support.update(b["categories"])
        for line in e["lines"]:
            if is_noise(line): continue
            k = _dedupe_key(line)
            if not k or k in seen: continue          # "availability for X" duplicates of X
            seen.add(k)
            cid, score, reasons = classify(line)
            if cid and not cid.startswith("group:") and score >= 0.4:
                counts[cid] += 1
                if len(samples) < 400: samples.append((line, cid, score))
            for r in reasons:
                if r.startswith("brand:"):
                    a = r[6:].split("->")[0]
                    if not _is_weak(a): brand_support.update(BRANDS[a]["categories"])
        # anti-stray rule: a category needs >=2 lines, or 1 line backed by a brand hit
        cats = {c: n for c, n in counts.items() if n >= 2 or c in brand_support}
        if not cats and domain_brand:
            a = max(domain_brand, key=len)
            cats = {BRANDS[a]["categories"][0]: counts.get(BRANDS[a]["categories"][0], 0)}
        o = cat_ov.get(dom, {})
        for c in o.get("remove", []): cats.pop(c, None)
        for c in o.get("add", []): cats.setdefault(c, counts.get(c, 0))
        conf = e["confidence"]
        if cats and all(n == 0 for n in cats.values()): conf = "manual" if o.get("add") else "brand-inferred"
        elif o: conf = conf + "+manual"
        vmap[dom] = {"name": e["name"], "categories": dict(sorted(cats.items(), key=lambda x: -x[1])), "rfq_count": e["rfq_count"], "confidence": conf}
        if e["merged_from"]: vmap[dom]["merged_from"] = e["merged_from"]
    for dom, m in ov.get("manual_vendors", {}).items():
        if dom not in vmap:
            vmap[dom] = {"name": m["name"], "categories": {c: 0 for c in m["categories"]}, "rfq_count": 0, "confidence": "manual", "evidence": m.get("evidence", "")}
    return vmap, samples


def self_check(proposal_path, verbose=True):
    P = json.load(open(proposal_path, encoding="utf-8"))
    ok, fails = 0, []
    for h in P["hard_examples"]:
        cid, score, reasons = classify(h["text"])
        if cid == h["category"]: ok += 1
        else: fails.append((h["text"], h["category"], cid, reasons[-2:]))
    n = len(P["hard_examples"])
    if verbose:
        print(f"hard-examples: {ok}/{n} = {ok/n:.1%}")
        for f in fails: print("  MISS:", f)
    return ok, n, fails


if __name__ == "__main__":
    tax = HERE
    ok, n, fails = self_check(os.path.join(tax, "taxonomy_proposal.json"))
    vmap, samples = build_vendor_map(os.path.join(tax, "vendor_evidence.json"))
    _R["vendor_map"] = vmap
    json.dump(_R, open(RULES_PATH, "w", encoding="utf-8"), indent=1, ensure_ascii=False)
    mapped = sum(1 for v in vmap.values() if v["categories"])
    print(f"categories: {len(CATS)}  brands: {len(BRANDS)}  vendors: {len(vmap)} (with >=1 category: {mapped})")
    for d, v in vmap.items():
        print(f"  {v['confidence']:15} {d:28} rfq={v['rfq_count']:3} {v['categories']}")
    print("samples:")
    import random; random.seed(3)
    for line, cid, score in random.sample(samples, min(10, len(samples))):
        print(f"  {line[:70]!r} -> {cid} ({score})")
