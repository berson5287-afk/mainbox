#!/usr/bin/env python3
r"""Repair the v4.1.37-and-earlier sent-RFQ contamination in Quote Coverage.

WHAT HAPPENED
-------------
When you replied to a customer and the reply was recorded as "Waiting on Vendor"
(because the customer's row was mis-categorized "Vendor / Supplier"), MaINbox
treated your outgoing message as a vendor RFQ and read its ATTACHMENT -- which was
your own customer's PO/quote PDF -- through the raw parser. The PDF's headers,
addresses, phone numbers, quote number, and price lines became dozens of phantom
"requested" rows in that job's coverage ledger (the 57-row Michael case).

Every phantom row carries source="rfq_extraction". Genuine lines you typed or
pasted carry source="manual" (or were captured by the real Quote RFQ extractor,
also "rfq_extraction" -- see the SAFETY note below). This script removes ONLY the
contamination while preserving everything real.

HOW IT DECIDES WHAT TO REMOVE
-----------------------------
A requested row is treated as CONTAMINATION and removed only when ALL of:
  * source == "rfq_extraction", AND
  * the row has no manual status, is not marked done, and is not in check/quoted
    state (i.e. you never reviewed or acted on it), AND
  * the ledger it lives in shows the contamination signature: it contains rows
    whose description text is clearly PO/quote boilerplate (address lines, phone
    numbers, "PER QUOTE#", "SUBTOTAL", "TOTAL", "PAGE", "BUYER", city/state/zip,
    "AMERICAN POWER ELECTRICAL", etc.) OR far more rfq_extraction rows than a
    human RFQ would contain.

A ledger with only clean, plausible material rows is LEFT ALONE even if they are
rfq_extraction -- this script never touches a healthy job.

SAFETY
------
* DRY-RUN by default: prints exactly what it WOULD remove, changes nothing.
  Re-run with --apply to write.
* Backs up the coverage JSON (timestamped) before writing.
* Preserves the SQLite mirror: writes the repaired ledger back to both the JSON
  and mainbox_data.db (app_state key "quote_coverage") so a restart cannot
  resurrect the old rows from the DB.
* Aborts if the DB is locked (MaINbox is open) -- close MaINbox first.
* Never deletes a whole ledger; only individual phantom rows. If removing the
  phantoms would empty a ledger, the ledger is kept (empty requested list) so its
  group/tracker/history are undisturbed.

USAGE
-----
    python repair_sent_rfq_contamination.py                 # dry run, all jobs
    python repair_sent_rfq_contamination.py --apply         # write the repair
    python repair_sent_rfq_contamination.py --job grp:g123  # limit to one job
    python repair_sent_rfq_contamination.py --json PATH      # explicit JSON path

Run with MaINbox CLOSED.
"""

import argparse
import json
import os
import re
import shutil
import sqlite3
import sys
from datetime import datetime


def default_paths():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    data_dir = os.path.join(base, "MaINbox")
    return (
        os.path.join(data_dir, "mainbox_quote_coverage.json"),
        os.path.join(data_dir, "mainbox_data.db"),
    )


# --- contamination signatures -------------------------------------------------

_BOILERPLATE_RE = re.compile(
    r"per\s+quote\s*#|sub\s*total|subtotal|grand\s+total|^\s*total\s*[:$]|page\s+\d|"
    r"\bbuyer\b|\bfrmn\b|\bremit\b|\bfax\b|purchase\s+order|"
    r"american\s+power\s+electrical|\bdisc\s*%|due\s+days|freight|"
    r"\bext\b\s*$|price\s*/?\s*um|qty\s+ord|"
    # bare PO/quote column-label fragments that leak as standalone rows -- these
    # carry no material identity (no part, qty, unit, or product noun):
    r"^\s*(?:quote|po|p\.o\.?|order|invoice|expiration|effective|ship|bill|attn|"
    r"terms|conditions|reference|contact|phone|email|date|number|quantity|"
    r"description|item\s+code|line|tax)\s*(?:#|no\.?|number|date|to|code)?\s*$|"
    # v4.1.39a: administrative FIELD rows that carry a label but no material identity.
    # These leaked from sent-quote/PO headers and were being kept because they contain
    # letters. Each is anchored so it never matches a genuine product description.
    r"^\s*customer\s+(?:number|no\.?|p\.?\s*o\.?(?:\s+number|\s+no\.?)?)\s*[:#]?\s*[\w-]*\s*$|"
    r"^\s*(?:sales\s*person|salesperson|writer|clerk|entered\s+by|taken\s+by|"
    r"our\s+order|your\s+order|order\s+no\.?|order\s+number|order\s+date|"
    r"ship\s+via|ship\s+date|shipped\s+via|via|f\.?o\.?b\.?|fob|"
    r"ship\s+qty|order\s+qty|b\/?o\s+qty|back\s*order\s+qty|qty\s+shipped|qty\s+ordered)"
    r"\s*[:#]?\s*[\w\/-]*\s*$|"
    r"^\s*net\s+\d+\s*(?:days?)?\s*$|"                       # payment terms: "Net 30 Days"
    r"^\s*terms\b.*$|"                                       # "TERMS NET30", "Terms: 2% 10 Net 30"
    r"^\s*\d+\s*%\s*(?:net\s+\d+)?\s*$|"                     # discount terms: "2% Net 10"
    r"^\s*(?:cod|c\.o\.d\.?|prepaid|collect|ppd|net\s+eom)\s*$|"
    # v4.1.42 (external audit): admin VALUE rows that slipped past label-only tests.
    r"^\s*\d+\s*(?:ea|pcs?|ft|feet|each|box|cs|case|roll|lot)\s*\.?\s*$|"   # "4ea", "200ea" qty-only
    r"^\s*ref(?:erence)?\s*[:#]?\s*[A-Za-z0-9\- ]+\s*$|"                    # "REFERENCE ABC123"
    r"^\s*tax\s*\d*\s*$|"                                                   # "TAX1"
    r"^\s*p\.?o\.?\s*[:#]?\s*[A-Za-z0-9\-]+\s*$|"                           # "PO12345"
    r"^\s*quote\s*[:#]?\s*[A-Za-z0-9\-]*\s*$",                              # "QUOTE1000999"
    re.I)

# Freight/routing carrier rows like "BX BRONX TRUCK", "UPS GROUND", "WILL CALL" --
# short all-caps routing tokens with no product noun. Kept separate so it can be
# tuned without touching the main boilerplate expression.
_ROUTING_RE = re.compile(
    r"^\s*(?:[A-Z]{1,3}\s+)?(?:bronx|brooklyn|queens|manhattan|nyc|truck|ups|fedex|"
    r"freight|will\s*call|pick\s*up|pickup|our\s+truck|customer\s+truck|ltl|common\s+carrier|"
    r"best\s+way|route|delivery)\b[A-Za-z \t]*$",
    re.I)

# Person-name rows (First Last, all letters, 2-3 tokens, no digits/product noun).
# Anchored to avoid matching descriptions; only used as a contaminant signal when a
# row has no part number, no quantity and no unit -- see _looks_boilerplate.
_PERSON_NAME_RE = re.compile(r"^\s*[A-Z][a-zA-Z]+(?:\s+[A-Z][a-zA-Z.]+){1,2}\s*$")

_MATERIAL_NOUN_RE = re.compile(
    r"\b(?:elbow|adapter|adaptor|coupling|conduit|emt|pvc|rigid|wire|cable|"
    r"connector|receptacle|breaker|panel|box|strut|fitting|lug|bushing|"
    r"nipple|clamp|strap|screw|bolt|nut|washer|gasket|valve|pipe|tube|"
    r"fixture|lamp|ballast|transformer|switch|outlet|cover|plate|bracket|"
    r"hanger|anchor|sleeve|reducer|tee|coupler|union|cap|plug|flange|"
    r"enclosure|raceway|whip|pull|terminal|fuse|relay|contactor|starter|"
    # v4.1.42 (external audit): rows like "Electrical Tape", "Myers Hub",
    # "Kindorf Channel", "Ground Rod", "Hubbell Device", "Crouse Hinds Condulet",
    # "Noalox Compound", "Greenlee Punch", "Burndy Die" were removed because these
    # everyday supply-house nouns were missing.
    r"tape|hub|channel|rod|device|condulet|compound|punch|die|crimp|tool|"
    r"ground|grounding|bond|bonding|liquidtight|sealtite|sealtight|greenfield|"
    r"romex|thhn|xhhw|thwn|mc|nm|so\s+cord|staple|tie|grommet|bit|blade|"
    r"disconnect|meter|socket|weatherhead|mast|riser|bell|ring|mud\s+ring|"
    r"extension|offset|locknut|chase|knockout|ko|seal|duct|tray|ladder|"
    r"photocell|timeclock|timer|dimmer|sensor|occupancy|ballast|driver|led|"
    r"strip|wrap|loom|heat\s*shrink|splice|tap|reel|spool)s?\b",
    re.I)

# v4.1.42: recognized electrical manufacturers -- a brand name on a row is material
# identity even when the generic noun is missing or unusual ("Crouse Hinds LB50").
_BRAND_RE = re.compile(
    r"\b(?:hubbell|square\s*d|crouse|hinds|myers|kindorf|noalox|greenlee|burndy|"
    r"caddy|ilsco|ideal|klein|panduit|thomas\s*(?:&|and)\s*betts|t\s*&\s*b|"
    r"appleton|killark|raco|steel\s*city|wiremold|carlon|cantex|southwire|cerro|"
    r"encore|leviton|lutron|pass\s*(?:&|and)\s*seymour|p\s*&\s*s|eaton|siemens|"
    r"schneider|cooper|arlington|bridgeport|topaz|halex|madison|orbit|intermatic|"
    r"tork|bussmann|littelfuse|mersen|ferraz|3m|gardner\s*bender|dottie|minerallac|"
    r"o-?z\s*gedney|red\s*dot|bell\s*outdoor|hoffman|milbank|midwest|ge|"
    r"general\s*electric|westinghouse|allen\s*bradley|acme|hammond|sola|federal\s*pacific|"
    r"zinsco|ite|challenger|murray|cutler\s*hammer|hubbel)\b",
    re.I)

_ADDRESS_RE = re.compile(
    r"\b\d{1,6}\s+[-A-Za-z0-9 .']+\s+(street|st\.?|road|rd\.?|avenue|ave\.?|drive|dr\.?|"
    r"lane|ln\.?|blvd|boulevard|parkway|pkwy|court|ct\.?|circle|cir\.?|place|way|north|south|east|west)\b",
    re.I)

_CITY_STATE_ZIP_RE = re.compile(
    r"[A-Za-z .'-]+,\s*[A-Z]{2}\s+\d{5}(?:-\d{4})?", re.I)

_PHONE_RE = re.compile(r"\b\d{3}[-.)\s]+\d{3}[-.\s]+\d{4}\b")

_MONEY_ONLY_RE = re.compile(r"^\s*\$?\s*[\d,]+\.\d{2}\s*(?:/\s*\w+)?\s*$")


def _row_text(row):
    parts = [str(row.get("part_number", "") or ""), str(row.get("description", "") or "")]
    return " ".join(p for p in parts if p).strip()


def _looks_boilerplate(row):
    """A single requested row that is clearly PO/quote boilerplate, not a material."""
    txt = _row_text(row)
    if not txt:
        return True  # an empty requested row is never legitimate
    if _BOILERPLATE_RE.search(txt):
        return True
    if _ADDRESS_RE.search(txt):
        return True
    if _CITY_STATE_ZIP_RE.search(txt):
        return True
    if _PHONE_RE.search(txt):
        return True
    if _ROUTING_RE.search(txt):
        return True
    desc = str(row.get("description", "") or "")
    if _MONEY_ONLY_RE.match(desc):
        return True
    # v4.1.39a: a row with no part number, no quantity and no unit that is just a
    # person's name (First Last) is a header artifact ("Steve Berson", "Attn: ...").
    _has_qty = bool(str(row.get("qty", "") or row.get("quantity", "") or "").strip())
    _has_unit = bool(str(row.get("unit", "") or row.get("uom", "") or "").strip())
    _has_part = bool(str(row.get("part_number", "") or "").strip())
    if not (_has_qty or _has_unit or _has_part) and _PERSON_NAME_RE.match(txt) \
            and not _MATERIAL_NOUN_RE.search(txt):
        return True
    return False


def _row_is_reviewed(rec, row):
    """True when the user has interacted with this row (manual status / done / check
    / an attached quote) -- such a row is NEVER auto-removed."""
    key_fields = [str(row.get("line_id", "") or ""), str(row.get("item_key", "") or "")]
    item_status = rec.get("item_status", {}) or {}
    done_keys = set(rec.get("manual_done_keys", []) or [])
    check_keys = set(rec.get("check_keys", []) or [])
    for k in key_fields:
        if not k:
            continue
        if k in item_status or k in done_keys or k in check_keys:
            return True
    # A row that already has a matched quote is reviewed by definition.
    if row.get("quoted") or row.get("allocated"):
        return True
    return False


def analyze_ledger(tkey, rec):
    """Return (contaminated_rows, kept_rows, is_contaminated_ledger)."""
    requested = rec.get("requested", []) or []
    if not requested:
        return [], [], False

    rfq_rows = [r for r in requested if str(r.get("source", "")) == "rfq_extraction"]
    boiler_rows = [r for r in requested if _looks_boilerplate(r)]

    # Contamination signature: the ledger contains PO/quote boilerplate rows, OR an
    # implausibly large rfq_extraction set (a human RFQ rarely exceeds ~25 lines).
    is_contaminated = bool(boiler_rows) or len(rfq_rows) > 25
    if not is_contaminated:
        return [], list(requested), False

    contaminated = []
    kept = []
    for r in requested:
        src = str(r.get("source", ""))
        if src != "rfq_extraction":
            kept.append(r)                      # manual/other rows always kept
            continue
        if _row_is_reviewed(rec, r):
            kept.append(r)                      # user acted on it -- keep
            continue
        if _looks_boilerplate(r):
            contaminated.append(r)              # phantom boilerplate -- remove
            continue
        # An rfq_extraction row that is NOT boilerplate in a contaminated ledger:
        # keep it only if it genuinely reads like a material line. v4.1.39a: "has
        # letters and a plausible length" was too weak -- it kept admin fields like
        # "CUSTOMER PO NUMBER" and "SALESPERSON". Require real material identity: a
        # part number, OR a recognized material noun, OR a quantity+unit structure,
        # OR (v4.1.41) a part-like catalog token in the text (letters+digits mixed,
        # 4+ chars: HBL5362, QO120, 812M) -- the extractor does not always fill the
        # part_number field, and "Hubbell HBL5362" must not be discarded.
        txt = _row_text(r)
        has_part = bool(str(r.get("part_number", "") or "").strip())
        has_noun = bool(_MATERIAL_NOUN_RE.search(txt))
        has_qty = bool(str(r.get("qty", "") or r.get("quantity", "") or "").strip())
        has_unit = bool(str(r.get("unit", "") or r.get("uom", "") or "").strip())
        # A leading "<qty> - <desc>" or "<desc> <qty> EA" pattern also signals a line.
        has_qty_pattern = bool(re.search(r"^\s*\d+\s*[-x]\s*\w", txt) or
                               re.search(r"\b\d+\s*(?:ea|pcs?|ft|feet|each|box|cs|case|roll|lot)\b", txt, re.I))
        has_partlike = bool(re.search(
            r"\b(?=[A-Za-z0-9\-]{4,}\b)(?=[A-Za-z0-9\-]*[A-Za-z])(?=[A-Za-z0-9\-]*\d)[A-Za-z0-9\-]+\b", txt))
        has_brand = bool(_BRAND_RE.search(txt))
        if has_part or has_noun or has_brand or (has_qty and has_unit) or has_qty_pattern or has_partlike:
            kept.append(r)
        else:
            contaminated.append(r)
    return contaminated, kept, True


def load_json(path):
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def _db_precheck(db_path):
    """Verify BEFORE any write that the DB exists, opens, and has the app_state
    schema MaINbox actually uses (state_key, json_value, updated_at). Returns
    (ok, reason). This runs before the JSON is touched so we never leave JSON and
    DB out of sync: if the DB cannot be written, we abort before changing anything."""
    if not os.path.exists(db_path):
        return True, "no-db"          # JSON-only environment is allowed
    try:
        conn = sqlite3.connect(db_path, timeout=2.0)
    except sqlite3.OperationalError as exc:
        return False, f"cannot open DB ({exc}) -- is MaINbox open?"
    try:
        conn.execute("PRAGMA busy_timeout=2000")
        cols = {row[1] for row in conn.execute("PRAGMA table_info(app_state)")}
        if not cols:
            return False, "app_state table not found in DB"
        needed = {"state_key", "json_value", "updated_at"}
        if not needed.issubset(cols):
            return False, (f"app_state has unexpected columns {sorted(cols)}; "
                           f"expected {sorted(needed)} -- refusing to write")
        # Confirm we can actually acquire a write lock right now (MaINbox closed).
        conn.execute("BEGIN IMMEDIATE")
        conn.execute("ROLLBACK")
        return True, "ok"
    except sqlite3.OperationalError as exc:
        return False, f"DB is locked ({exc}) -- close MaINbox and re-run"
    finally:
        conn.close()


def _read_db_state(db_path, state_key):
    """Return the current json_value string for a state_key in app_state, or None if
    the DB/row is absent. Used to snapshot before writing so a failed JSON write can
    be rolled back in the DB."""
    if not os.path.exists(db_path):
        return None
    try:
        conn = sqlite3.connect(db_path, timeout=2.0)
    except sqlite3.OperationalError:
        return None
    try:
        row = conn.execute(
            "SELECT json_value FROM app_state WHERE state_key = ?", (state_key,)).fetchone()
        return row[0] if row else None
    except sqlite3.OperationalError:
        return None
    finally:
        conn.close()


def _restore_db_state(db_path, state_key, prior_json_value):
    """Restore a previously-snapshotted json_value for a state_key. Returns True on
    success. Used to roll the DB back if the JSON write fails after the DB update."""
    if not os.path.exists(db_path):
        return False
    try:
        conn = sqlite3.connect(db_path, timeout=5.0)
    except sqlite3.OperationalError:
        return False
    try:
        conn.execute(
            "INSERT OR REPLACE INTO app_state(state_key, json_value, updated_at) "
            "VALUES (?, ?, ?)", (state_key, prior_json_value, datetime.now().isoformat()))
        conn.commit()
        return True
    except sqlite3.OperationalError:
        return False
    finally:
        conn.close()


def write_db_mirror(db_path, coverage_obj):
    """Write the repaired ledger into the SQLite mirror using MaINbox's OWN schema
    and idiom: app_state(state_key, json_value, updated_at) via INSERT OR REPLACE.
    Assumes _db_precheck already passed (called from main before the JSON write)."""
    if not os.path.exists(db_path):
        print(f"  (no DB mirror at {db_path}; JSON-only write)")
        return True
    try:
        conn = sqlite3.connect(db_path, timeout=5.0)
    except sqlite3.OperationalError as exc:
        print(f"  ! Could not open DB ({exc}). No DB change made.")
        return False
    try:
        conn.execute("PRAGMA busy_timeout=5000")
        payload = json.dumps(coverage_obj, ensure_ascii=False)
        now = datetime.now().isoformat()
        conn.execute(
            "INSERT OR REPLACE INTO app_state(state_key, json_value, updated_at) "
            "VALUES (?, ?, ?)",
            ("quote_coverage", payload, now))
        conn.commit()
        print("  DB mirror updated (app_state key 'quote_coverage').")
        return True
    except sqlite3.OperationalError as exc:
        # Distinguish a genuine lock from any other operational error so the
        # message is accurate (the old code reported every error as a lock).
        msg = str(exc)
        if "locked" in msg.lower() or "busy" in msg.lower():
            print(f"  ! DB is locked ({exc}). Close MaINbox and re-run. No DB change made.")
        else:
            print(f"  ! DB write failed ({exc}). No DB change made.")
        return False
    finally:
        conn.close()


def main():
    ap = argparse.ArgumentParser(description="Repair sent-RFQ coverage contamination.")
    ap.add_argument("--apply", action="store_true", help="write the repair (default is dry-run)")
    ap.add_argument("--job", default="", help="limit to one ledger key, e.g. grp:g123")
    ap.add_argument("--json", default="", help="explicit path to mainbox_quote_coverage.json")
    ap.add_argument("--db", default="", help="explicit path to mainbox_data.db")
    args = ap.parse_args()

    json_path, db_path = default_paths()
    if args.json:
        json_path = args.json
    if args.db:
        db_path = args.db

    if not os.path.exists(json_path):
        print(f"Coverage file not found: {json_path}")
        print("Pass --json PATH if your data lives elsewhere.")
        return 2

    print(f"Coverage JSON: {json_path}")
    print(f"DB mirror:     {db_path}")
    print(f"Mode:          {'APPLY (will write)' if args.apply else 'DRY RUN (no changes)'}")
    print("-" * 70)

    data = load_json(json_path)
    threads = data.get("threads", {}) or {}

    total_removed = 0
    affected = 0
    for tkey, rec in threads.items():
        if args.job and tkey != args.job:
            continue
        if not isinstance(rec, dict):
            continue
        contaminated, kept, is_bad = analyze_ledger(tkey, rec)
        if not is_bad or not contaminated:
            continue
        affected += 1
        label = str(rec.get("subject", "") or rec.get("customer", "") or tkey)[:70]
        print(f"\nLedger: {tkey}")
        print(f"  {label}")
        print(f"  requested rows: {len(rec.get('requested', []) or [])} "
              f"-> keeping {len(kept)}, removing {len(contaminated)} phantom")
        for r in contaminated[:40]:
            print(f"    REMOVE: {(_row_text(r) or '(empty row)')[:80]}")
        if len(contaminated) > 40:
            print(f"    ... and {len(contaminated) - 40} more")
        total_removed += len(contaminated)
        if args.apply:
            rec["requested"] = kept
            rec["updated_at"] = datetime.now().isoformat()
            rec["contamination_repaired_at"] = datetime.now().isoformat()

    print("\n" + "=" * 70)
    print(f"Ledgers affected: {affected}   Phantom rows: {total_removed}")

    if not args.apply:
        print("\nDRY RUN complete -- nothing was changed.")
        print("Re-run with --apply to write the repair.")
        return 0

    if total_removed == 0:
        print("Nothing to write.")
        return 0

    # Pre-flight: confirm the DB mirror can actually be written BEFORE we touch the
    # JSON. This is the fix for the out-of-sync failure mode -- if the DB is locked,
    # missing the expected schema, or otherwise unwritable, we abort now and leave
    # both the JSON and the DB exactly as they were.
    db_ok, db_reason = _db_precheck(db_path)
    if not db_ok:
        print(f"\n! Cannot safely write the DB mirror: {db_reason}")
        print("  Nothing was changed. Fix the above and re-run --apply.")
        print("  (Most commonly: close MaINbox so the database is not locked.)")
        return 1
    if db_reason == "no-db":
        print("\nNote: no DB mirror present; this will be a JSON-only repair.")

    # Back up, then write DB mirror, then the JSON. Writing the DB first means that
    # if anything unexpected still fails at the DB step, the JSON is untouched and
    # the two remain consistent (the pre-check already proved the DB is writable).
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = f"{json_path}.contamination_backup_{stamp}"
    shutil.copy2(json_path, backup)
    print(f"\nBackup written: {backup}")

    data["updated_at"] = datetime.now().isoformat()

    # v4.1.39a: snapshot the DB's current quote_coverage value so that if the JSON
    # write fails AFTER the DB is updated, we can roll the DB back and keep the two
    # copies consistent (previously a failed JSON write left the DB newer).
    _db_prior = _read_db_state(db_path, "quote_coverage")

    if not write_db_mirror(db_path, data):
        print("\n! DB mirror write failed after the pre-check passed (unexpected).")
        print("  The JSON was NOT modified; nothing is out of sync. Re-run --apply.")
        return 1

    tmp = json_path + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
        os.replace(tmp, json_path)
    except OSError as exc:
        # JSON write failed after the DB was updated -- roll the DB back so it does
        # not become newer than the JSON.
        print(f"\n! Writing the JSON failed ({exc}).")
        if _db_prior is not None and _restore_db_state(db_path, "quote_coverage", _db_prior):
            print("  The DB mirror was rolled back to its previous value; nothing is out of sync.")
        else:
            print("  WARNING: could not roll back the DB. Restore the JSON backup and")
            print(f"  re-run, or reconcile manually. Backup: {backup}")
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        return 1
    print(f"Coverage JSON updated: {json_path}")

    print("\nRepair complete. Open MaINbox and refresh Quote Coverage to confirm.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
