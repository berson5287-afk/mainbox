"""Validates the coverage per-item status priority + outstanding count added to
MaINbox's compute_coverage_for_record. The item matchers (coverage_item_keys /
coverage_items_match) are existing, trusted MaINbox code and are stubbed here so
the test isolates the NEW decision tree. Mirrors the real branch verbatim."""

PASS, FAIL, results = "  ok  ", " FAIL ", []
def check(name, cond):
    results.append(bool(cond)); print(f"[{PASS if cond else FAIL}] {name}")

# --- the exact decision tree from compute_coverage_for_record ---
def decide(manual, matched_quote, sourcing_ack, auto_wait=False):
    covered = False
    if manual in ("quoted", "waiting", "partial", "no_bid", "outstanding"):
        status = manual
        if manual == "quoted":
            covered = True
    elif matched_quote:
        status, covered = "quoted", True
    elif sourcing_ack or auto_wait:
        status = "waiting"
    else:
        status = "outstanding"
    outstanding = status not in ("quoted", "no_bid")
    return status, covered, outstanding

# auto progression (no manual override)
check("bare request -> outstanding (counts)", decide("", False, False) == ("outstanding", False, True))
check("sourcing ack -> waiting (still counts)", decide("", False, True) == ("waiting", False, True))
check("per-item auto_wait -> waiting (still counts)", decide("", False, False, auto_wait=True) == ("waiting", False, True))
check("matched quote -> quoted (not counted)", decide("", True, False) == ("quoted", True, False))
check("matched quote beats ack", decide("", True, True)[0] == "quoted")
check("matched quote beats per-item auto_wait", decide("", True, False, auto_wait=True)[0] == "quoted")

# manual override always wins
check("manual no_bid wins over ack (not counted)", decide("no_bid", False, True) == ("no_bid", False, False))
check("manual no_bid wins over auto_wait", decide("no_bid", False, False, auto_wait=True) == ("no_bid", False, False))
check("manual no_bid wins over matched quote", decide("no_bid", True, True) == ("no_bid", False, False))
check("manual outstanding beats auto-waiting", decide("outstanding", False, True, auto_wait=True) == ("outstanding", False, True))
check("manual quoted -> covered, not counted", decide("quoted", False, False) == ("quoted", True, False))
check("manual waiting counts as outstanding", decide("waiting", False, False) == ("waiting", False, True))
check("manual partial counts as outstanding", decide("partial", True, True) == ("partial", False, True))

# --- outgoing-request -> Waiting LINK decision (subject ref OR strong item match) ---
def linked(rec_subj, req_subj, matched_count, requested_count):
    subj_link = bool(rec_subj and len(rec_subj) >= 5 and req_subj and rec_subj in req_subj)
    strong_item_link = matched_count >= 2 or (matched_count and matched_count == requested_count)
    return subj_link or strong_item_link

check("subject reference links (274870 quote in request subj)",
      linked("274870 quote", "price and availability request - 274870 quote", 0, 4))
check("two+ item matches link (no subject)", linked("", "vendor rfq", 2, 4))
check("all items match links even if single-item", linked("", "vendor rfq", 1, 1))
check("one generic match on multi-item record does NOT link", not linked("", "vendor rfq", 1, 4))
check("short/generic record subject does NOT subject-link", not linked("rfq", "re: rfq stuff", 0, 3))
check("unrelated subject + no item match does NOT link", not linked("acme job 7", "different vendor rfq", 0, 5))

# coverage_set_item_status sync behavior (manual_done kept in step with 'quoted'/clear)
def set_status(rec, key, status):
    if status not in ("quoted","waiting","partial","no_bid","outstanding"): return
    isd = rec.setdefault("item_status", {})
    isd[key] = status
    dk = rec.setdefault("manual_done_keys", [])
    if status == "quoted":
        if key not in dk: dk.append(key)
    else:
        rec["manual_done_keys"] = [k for k in dk if k != key]

rec = {}
set_status(rec, "k", "quoted")
check("set quoted -> item_status + manual_done", rec["item_status"]["k"]=="quoted" and "k" in rec["manual_done_keys"])
set_status(rec, "k", "no_bid")
check("re-mark no_bid clears manual_done", rec["item_status"]["k"]=="no_bid" and "k" not in rec["manual_done_keys"])
set_status(rec, "k", "outstanding")
check("mark outstanding clears manual_done", rec["item_status"]["k"]=="outstanding" and "k" not in rec["manual_done_keys"])
set_status(rec, "k", "bogus")
check("bogus status ignored", rec["item_status"]["k"]=="outstanding")

# --- NEW v3.9.4: bare catalog-number match (632S as part on one side, desc on the other) ---
import re as _re
def cat_tokens(item):
    if isinstance(item, str): item = {"description": item}
    txt = (str(item.get("part_number","") or "") + " " + str(item.get("description","") or "")).lower()
    txt = _re.sub(r"[^a-z0-9]+", " ", txt)
    return {t for t in txt.split() if len(t) >= 3 and any(c.isdigit() for c in t) and any(c.isalpha() for c in t)}
def catalog_match(a, b):       # the rule added to coverage_items_match
    return bool(cat_tokens(a) & cat_tokens(b))

check("632S part-vs-desc now matches",
      catalog_match({"part_number":"632S","description":""}, {"part_number":"","description":"25 632S"}))
check("632S desc-vs-desc matches",
      catalog_match({"description":"632s"}, {"description":"qty 632s units"}))
check("different catalog numbers do NOT match", not catalog_match({"part_number":"632S"}, {"part_number":"451T"}))
check("plain qty/size words yield no catalog token (no false match)",
      not catalog_match({"description":"100 ft 3/4 emt"}, {"description":"100 ft 1 pvc"}))
check("EMT-only line has no catalog token (rule stays targeted)", cat_tokens({"description":"3/4 emt"}) == set())

# --- v3.9.4: link a QUOTE not threaded to the RFQ -- subject reference OR strong item match ---
REQUEST_PHRASES = ["please provide", "price and availability", "please quote", "request for quote", "rfq"]
def strip_trailer(text):       # mirrors _strip_reply_boilerplate_for_intent (flattens, for GATING)
    lines = []
    for line in str(text or "").replace("\r", "\n").split("\n"):
        l = line.strip().lower()
        if l.startswith("--"): break
        if any(m in l for m in ("-----original message-----","from:","sent:","to:","subject:","on "," wrote:")) and len(lines) >= 1:
            break
        lines.append(line)
    return " ".join(lines)
def reply_above_quote(text):   # mirrors _coverage_reply_above_quote (KEEPS newlines, for extraction)
    raw = str(text or "").replace("\r\n", "\n").replace("\r", "\n")
    cut = len(raw)
    for pat in (r"(?m)^\s*>?\s*On\s.{0,250}?\bwrote:", r"(?im)^\s*From:\s"):
        m = _re.search(pat, raw)
        if m and m.start() < cut: cut = m.start()
    head = raw[:cut].strip()
    return head or raw.strip()
def extract_items(text):       # crude stand-in for extract_materials_from_quote_text
    items = []
    for line in str(text or "").split("\n"):
        l = line.strip().lstrip("-").strip()
        if _re.match(r"^\d+\s*x", l, _re.IGNORECASE):
            desc = _re.sub(r"^\d+\s*x\s*-?\s*", "", l, flags=_re.IGNORECASE)
            desc = _re.sub(r"\$\S*", "", desc).strip(" -")
            items.append({"description": desc})
    return items
def cim(a, b):                 # stub item matcher: shared description token (len > 3)
    at = {t for t in str(a.get("description","")).lower().split() if len(t) > 3}
    bt = {t for t in str(b.get("description","")).lower().split() if len(t) > 3}
    return bool(at & bt)
def norm_subj(s):
    s = str(s or "").lower()
    s = _re.sub(r"^\s*(re|fw|fwd)\s*:\s*", "", s)
    s = _re.sub(r"[^a-z0-9 ]+", " ", s); s = _re.sub(r"\s+", " ", s).strip()
    return s
def quote_link(subject, body, threads):   # mirrors _coverage_record_for_quote
    gate = strip_trailer(body)
    if any(p in gate.lower() for p in REQUEST_PHRASES): return {}     # request -> Waiting handles it
    if not _re.search(r"\$\s*\d|\b\d+\.\d{2}\b", gate): return {}     # no price -> not a quote
    recs = [r for r in threads if not r.get("dismissed") and r.get("requested")]
    if not recs: return {}
    msg = norm_subj(subject)                                          # (1) subject reference
    if msg:
        best, blen = {}, 0
        for r in recs:
            rs = norm_subj(r.get("subject",""))
            if rs and len(rs) >= 5 and rs in msg and len(rs) > blen: best, blen = r, len(rs)
        if best: return best
    items = extract_items(reply_above_quote(body)) or extract_items(body)   # (2) strong item match
    if not items: return {}
    scored = []
    for r in recs:
        hits = sum(1 for qi in items if any(cim(ri, qi) for ri in r.get("requested", [])))
        if hits >= 2: scored.append((hits, r))
    if not scored: return {}
    scored.sort(key=lambda t: t[0], reverse=True)
    if len(scored) > 1 and scored[0][0] == scored[1][0]: return {}    # ambiguous -> don't guess
    return scored[0][1]

THREADS = [{"subject":"274870 Quote","requested":[{"k":1}],"dismissed":False}]
QUOTE_BODY  = "Here is your pricing:\n100 ft 3/4 EMT - $85.00\n632S - $12.50"
REQ_BODY    = "Please provide price and availability for:\n100 ft 3/4 EMT\n632S"
NOPRICE_BODY= "Here are the items:\n100 ft 3/4 EMT\n632S"
QSUBJ = "RE: Price and Availability Request - 274870 Quote"
STEVE_QUOTE_BODY = ('- 100 ft - 3/4" EMT   $100/c\n- 25x - 632S - 632S  $25/c\n'
                    '- 100 ft - 3/4 EMT  $100/c\n- 100 ft - 3/4 Sealtight  $150/c\n\n'
                    'On Sat, Jun 6, 2026 at 7:43 PM Steve Berson <steveb@americanpoweresc.com> wrote:\n'
                    'Hi, Please provide me with price and availability for the below. Thanks! '
                    '- 100 ft - 3/4" EMT - 25x - 632S - 632S - 100 ft - 3/4 EMT - 100 ft - 3/4 Sealtight')

check("quote w/ price + subject reference links to record",
      quote_link(QSUBJ, QUOTE_BODY, THREADS).get("subject") == "274870 Quote")
check("REAL quote over quoted-back request still links (trailer ignored)",
      quote_link(QSUBJ, STEVE_QUOTE_BODY, THREADS).get("subject") == "274870 Quote")
check("request body never links as quoted (stays Waiting)", quote_link(QSUBJ, REQ_BODY, THREADS) == {})
check("quote with no recognizable price does NOT link", quote_link(QSUBJ, NOPRICE_BODY, THREADS) == {})
check("quote whose subject references nothing does NOT link",
      quote_link("RE: unrelated note", QUOTE_BODY, THREADS) == {})
check("dismissed record is not linked",
      quote_link(QSUBJ, QUOTE_BODY, [{"subject":"274870 Quote","requested":[{"k":1}],"dismissed":True}]) == {})

# --- v3.9.4: PARTIAL quote (Steve's real email) -- item-match link + only-priced extraction ---
STEVE_PARTIAL_QUOTE = ('Only have a few of these items please see below\n\n'
    '- 200x - 3/4" emt pipe- $100/c\n'
    '- 75x - 4" octagon boxes 2 1/8" deep 1/2"ko $/50/c\n'
    '- 1x - step bit until 1"- $24ea\n\nThank you\n\n'
    'On Sat, Jun 6, 2026 at 8:59 PM Steve Berson <steveb@americanpoweresc.com> wrote:\n'
    'Hi, Please provide me with price and availability for the below. Thanks! '
    '- 200x - 3/4" emt pipe - 75x - 4" octagon boxes 2 1/8" deep 1/2"ko - 25x - 3/4" emt compression connector '
    '- 100x - 1/2" push in low voltage connector Arlington 4400 - 100x - 8/32"x2" machine screw combo head '
    '- 100x - 8/32" nuts - 100x - #8x3/4" washers - 1x - 3" emt elbow 90DEG '
    '- 10x - 11/16"square extension box - 10x - 4" square extension box - 1x - step bit until 1"')
REQ12 = [{"description": d} for d in [
    "3/4 emt pipe","4 octagon boxes 2 1/8 deep 1/2 ko","3/4 emt compression connector",
    "push in low voltage connector arlington 4400","8/32 machine screw combo head","8/32 nuts",
    "8 washers hole for screw","3 emt elbow 90deg","11/16 square extension box",
    "4 square extension box","step bit until 1"]]
REC12 = {"subject":"rfq","requested":REQ12,"dismissed":False}   # short subject -> no subject-ref

check("reply-above-quote keeps only the 3 priced lines (not the 12 quoted under it)",
      len(extract_items(reply_above_quote(STEVE_PARTIAL_QUOTE))) == 3)
check("partial quote links via item match when subject is too generic to reference",
      quote_link("RE: rfq", STEVE_PARTIAL_QUOTE, [REC12]) is REC12)
ONE_ITEM = '- 1x - step bit until 1"- $24ea\n\nThank you\n\nOn Sat, Jun 6 wrote:\nHi please provide pricing'
check("single matched item is too weak to item-match link", quote_link("RE: rfq", ONE_ITEM, [REC12]) == {})
RECA = {"subject":"job a","requested":[{"description":"3/4 emt pipe"},{"description":"octagon boxes"}],"dismissed":False}
RECB = {"subject":"job b","requested":[{"description":"3/4 emt pipe"},{"description":"octagon boxes"}],"dismissed":False}
TWO_ITEM = ('- 200x - 3/4" emt pipe- $100/c\n- 75x - octagon boxes- $50/c\n\nThank you\n\n'
            'On Sat wrote:\nplease provide')
check("ambiguous item match (two records tie) does NOT mark either quoted",
      quote_link("RE: misc", TWO_ITEM, [RECA, RECB]) == {})

# --- v3.9.4: a priced "Thank you" reply is a quote, so the Sent loop must still run the
# coverage hook on it instead of skipping it as a closed acknowledgement. ---
def closed_ack_still_quotes(body):     # mirrors the price gate added to the closed-ack branch
    return bool(_re.search(r"\$\s*\d|\b\d+\.\d{2}\b", str(body or "")))
check("priced thank-you reply still triggers the coverage hook (not skipped as an ack)",
      closed_ack_still_quotes(STEVE_PARTIAL_QUOTE))
check("a bare thank-you with no pricing is not treated as a quote",
      not closed_ack_still_quotes("Thanks, got it! Appreciate it."))

# --- v3.9.4: vendor pricing arrives INBOUND (a vendor replies to our RFQ with a price). The
# inbox scan feeds those replies to the same linker; a referencing subject or item match links
# them and flips matched items to Quoted. Body shape is identical to an outbound quote. ---
VENDOR_REPLY = ('Pricing as requested, all in stock:\n'
                '- 200x - 3/4" emt pipe - $100/c\n- 75x - 4" octagon boxes - $50/c\n\nThanks\n\n'
                'On Sat, Jun 6, 2026 at 9:10 PM Steve Berson <steveb@americanpoweresc.com> wrote:\n'
                'Hi, Please provide me with price and availability for the below. Thanks! '
                '- 200x - 3/4" emt pipe - 75x - 4" octagon boxes 2 1/8" deep 1/2"ko - 1x - step bit until 1"')
check("inbound vendor pricing reply links by subject reference",
      quote_link("RE: Price and Availability Request - test quote", VENDOR_REPLY,
                 [{"subject":"test quote","requested":REQ12,"dismissed":False}]).get("subject") == "test quote")

# --- Re-quote: drops the old quote + sticky manual, marks the item Waiting (auto) so a fresh quote re-flips it ---
def requote(rec, keys):   # mirrors coverage_requote_items
    keyset = {k for k in keys if k}
    targets = [r for r in rec.get("requested", []) if r.get("pk") in keyset]
    st = rec.setdefault("item_status", {})
    for k in list(keyset):
        st.pop(k, None)
    rec["manual_done_keys"] = [k for k in rec.get("manual_done_keys", []) if k not in keyset]
    rec["quoted"] = [q for q in rec.get("quoted", []) if not any(cim(t, q) for t in targets)]
    aw = rec.setdefault("auto_waiting_keys", [])
    for k in keyset:
        if k not in aw:
            aw.append(k)
    return rec

RQ = {"requested":[{"pk":"p1","description":"3/4 emt pipe"},{"pk":"p2","description":"octagon boxes"}],
      "quoted":[{"description":"3/4 emt pipe"}], "item_status":{"p1":"quoted"},
      "manual_done_keys":["p1"], "auto_waiting_keys":[]}
requote(RQ, ["p1"])
check("re-quote drops the matched quote for the item", RQ["quoted"] == [])
check("re-quote clears the sticky manual status", "p1" not in RQ["item_status"])
check("re-quote marks the item Waiting (auto)", "p1" in RQ["auto_waiting_keys"])

RQ2 = {"requested":[{"pk":"p1","description":"3/4 emt pipe"},{"pk":"p2","description":"octagon boxes"}],
       "quoted":[{"description":"octagon boxes"}], "item_status":{"p2":"quoted"},
       "manual_done_keys":["p2"], "auto_waiting_keys":["p1"]}
requote(RQ2, ["p2"])
check("re-quote keeps other items' Waiting and adds the re-quoted item",
      "p1" in RQ2["auto_waiting_keys"] and "p2" in RQ2["auto_waiting_keys"] and RQ2["quoted"] == [])

print("-"*52)
ok, total = sum(results), len(results)
print(f"{ok}/{total} checks passed")
raise SystemExit(0 if ok==total else 1)
