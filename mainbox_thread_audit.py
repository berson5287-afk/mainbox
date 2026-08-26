"""mainbox_thread_audit.py -- COM-on-the-Tk-main-thread audit for MaINbox.

Read-only. Imports nothing from the app and never touches Outlook: it parses the
newest MaINbox_v*.pyw, builds a call graph of OutlookWorkflowMonitor, marks which
functions actually touch Outlook COM, marks which are reachable from a Tk main-thread
entry point WITHOUT crossing run_outlook_worker, and prints the intersection ranked by
how dangerous the COM is.

    python mainbox_thread_audit.py                 # newest MaINbox_v*.pyw
    python mainbox_thread_audit.py <file.pyw>      # a specific version

Run it every delivery. The point is to catch work drifting back onto the main thread
the moment it happens, instead of noticing it later as a freeze.

Known limits -- read a finding, do not just count them:
  * a static call graph over Python is never a proof; dispatch through getattr, a
    lambda, or a callback table is invisible to it
  * comments and string literals are blanked before scanning, because the app
    documents its own COM history in prose and a naive scan reads that as code
  * a dual-path function (reachable from BOTH a worker and the main thread) is
    reported and tagged [DUAL-PATH]. That is deliberate: it is not a pass. Two real
    ones are _collect_sales_uncategorized_rows (by design) and import_from_folder,
    which a worker uses for Update-Since while full_update still walks it on main
  * "SINGLE-ITEM COM" is usually the app's deliberate open-in-Outlook pattern, not
    a defect; the FOLDER WALKS section is where the real freezes live
"""
import ast, io, os, re, sys
from collections import defaultdict

def _newest_app():
    """Highest vX_Y_Z next to this script -- never a hardcoded version, which goes
    stale the moment a delivery ships."""
    here = os.path.dirname(os.path.abspath(__file__))
    best, best_key = None, ()
    for name in os.listdir(here):
        m = re.match(r"MaINbox_v(\d+)_(\d+)_(\d+)_.*\.pyw$", name)
        if m:
            key = tuple(int(g) for g in m.groups())
            if key > best_key:
                best, best_key = os.path.join(here, name), key
    return best


APP = sys.argv[1] if len(sys.argv) > 1 else _newest_app()
if not APP or not os.path.exists(APP):
    raise SystemExit("no MaINbox_v*.pyw found next to this script")
src = io.open(APP, encoding="utf-8").read()
lines = src.split("\n")
tree = ast.parse(src)

# ---------------------------------------------------------------- COM signals
# Tiered by what actually hurts. A folder WALK on the main thread freezes the UI for
# as long as Outlook takes; a single GetItemFromID+Display is bounded and is the
# app's documented "open in Outlook" pattern.
WALK = ("Restrict(", ".GetFirst()", ".GetNext()", ".Items", "GetDefaultFolder",
        ".Folders", "GetSharedDefaultFolder", "PropertyAccessor", "ResolveAll()")
ITEM = ("GetItemFromID", ".Display()", ".Send()", ".Save()", ".Reply()",
        ".ReplyAll()", ".Forward()", "CreateItem(", ".Attachments", ".Recipients",
        ".Categories", ".UnRead", ".Move(", ".Delete()")
ENTRY = ("fresh_outlook(", "Dispatch(\"Outlook", "GetNamespace(")

FUNCS = {}          # qualname -> node
SEGS = {}           # qualname -> source
PARENT = {}         # qualname -> enclosing qualname or None
LINE = {}


class Collect(ast.NodeVisitor):
    def __init__(self):
        self.stack = []

    def visit_ClassDef(self, n):
        self.stack.append(n.name)
        self.generic_visit(n)
        self.stack.pop()

    def _fn(self, n):
        qn = ".".join(self.stack + [n.name])
        i = 2
        base = qn
        while qn in FUNCS:
            qn = "%s#%d" % (base, i)
            i += 1
        FUNCS[qn] = n
        SEGS[qn] = ast.get_source_segment(src, n) or ""
        LINE[qn] = n.lineno
        PARENT[qn] = ".".join(self.stack) if self.stack else None
        self.stack.append(n.name)
        self.generic_visit(n)
        self.stack.pop()

    visit_FunctionDef = _fn
    visit_AsyncFunctionDef = _fn


Collect().visit(tree)

# method short-name -> qualnames (for self.X() resolution)
by_short = defaultdict(list)
for qn in FUNCS:
    by_short[qn.split(".")[-1].split("#")[0]].append(qn)

# ------------------------------------------------------- direct COM per function
import tokenize as _tok
from io import StringIO as _SIO


def strip_noncode(s):
    """Blank comments and string literals IN PLACE, preserving every other character
    and position. Without this the audit reads the app's own prose -- e.g.
    _sent_tracker_check_tick documents that it 'connects its OWN thread-local
    fresh_outlook()', and a naive scan flags that comment as a COM call. Blanking in
    place (rather than dropping tokens) matters because re-joining tokens puts spaces
    around dots and destroys the `.Items` / `.Categories` attribute syntax the COM
    patterns match on."""
    rows = s.split("\n")
    try:
        spans = []
        for t in _tok.generate_tokens(_SIO(s).readline):
            if t.type in (_tok.COMMENT, _tok.STRING):
                spans.append((t.start, t.end))
    except Exception:
        return "\n".join(l.split("#")[0] for l in rows)
    for (r1, c1), (r2, c2) in spans:
        for r in range(r1 - 1, min(r2, len(rows))):
            line = rows[r]
            a = c1 if r == r1 - 1 else 0
            b = c2 if r == r2 - 1 else len(line)
            rows[r] = line[:a] + " " * max(0, min(b, len(line)) - a) + line[min(b, len(line)):]
    return "\n".join(rows)


def own_source(qn):
    """Source of qn minus its nested function bodies, so a parent is not blamed
    for COM that only happens inside a nested worker closure. Comments and strings
    are stripped so only real code is scanned."""
    node = FUNCS[qn]
    inner = []
    for child in ast.walk(node):
        if child is node:
            continue
        if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
            seg = ast.get_source_segment(src, child)
            if seg:
                inner.append(seg)
    s = SEGS[qn]
    for seg in inner:
        s = s.replace(seg, "")
    return strip_noncode(s)


direct = {}
for qn in FUNCS:
    s = own_source(qn)
    hits = {"walk": [t for t in WALK if t in s],
            "item": [t for t in ITEM if t in s],
            "entry": [t for t in ENTRY if t in s]}
    if any(hits.values()):
        direct[qn] = hits

# ------------------------------------------------------------------ call graph
calls = defaultdict(set)
for qn in FUNCS:
    s = own_source(qn)
    for m in re.finditer(r"self\.([A-Za-z_][A-Za-z0-9_]*)\s*\(", s):
        for t in by_short.get(m.group(1), []):
            calls[qn].add(t)
    # nested closures execute inside their parent unless handed to the worker
    for child in FUNCS:
        if PARENT.get(child) and PARENT[child].split("#")[0] == qn.split("#")[0]:
            pass

# ------------------------------------------------- worker boundary + main roots
worker_roots = set()
for qn in FUNCS:
    s = SEGS[qn]
    for m in re.finditer(r"run_outlook_worker\(\s*([A-Za-z_][A-Za-z0-9_]*)", s):
        name = m.group(1)
        for cand in FUNCS:
            if PARENT.get(cand) == qn.split("#")[0] or cand.endswith("." + name):
                if cand.split(".")[-1].split("#")[0] == name:
                    worker_roots.add(cand)
for qn in FUNCS:
    doc = ast.get_docstring(FUNCS[qn]) or ""
    if "WORKER-THREAD ONLY" in doc.upper() or "WORKER THREAD ONLY" in doc.upper():
        worker_roots.add(qn)
    # Only the run_outlook_worker signature convention (work(outlook, namespace))
    # counts as an automatic worker root. `folder`/`msg` params were tried and are
    # WRONG: import_from_folder(folder, ...) takes a folder yet is called straight
    # from full_update / update_since_date, which are main-thread (they build Tk
    # loading windows). Excusing on those params hid the largest walk in the app.
    args = [a.arg for a in FUNCS[qn].args.args]
    if "namespace" in args:
        worker_roots.add(qn)

# propagate worker-ness downward
worker = set(worker_roots)
changed = True
while changed:
    changed = False
    for qn in list(worker):
        for t in calls.get(qn, ()):
            if t not in worker:
                worker.add(t)
                changed = True

# main-thread entry points: Tk wiring anywhere in the file
main_roots = set()
for m in re.finditer(r"command\s*=\s*(?:lambda[^:]*:\s*)?self\.([A-Za-z_][A-Za-z0-9_]*)", src):
    main_roots.update(by_short.get(m.group(1), []))
for m in re.finditer(r"\.bind\([^,]+,\s*(?:lambda[^:]*:\s*)?self\.([A-Za-z_][A-Za-z0-9_]*)", src):
    main_roots.update(by_short.get(m.group(1), []))
for m in re.finditer(r"after\(\s*[^,]+,\s*self\.([A-Za-z_][A-Za-z0-9_]*)", src):
    main_roots.update(by_short.get(m.group(1), []))
for m in re.finditer(r"after_idle\(\s*self\.([A-Za-z_][A-Za-z0-9_]*)", src):
    main_roots.update(by_short.get(m.group(1), []))
# LOCAL CLOSURES wired to Tk. Missing these hid the biggest finding in the app:
# "Full Update" is command=run_full_update, a nested closure that calls
# self.full_update() -- so the whole synchronous deep-walk chain looked unreachable
# from the main thread when it is in fact the definition of a main-thread walk.
for m in re.finditer(r"command\s*=\s*([A-Za-z_][A-Za-z0-9_]*)\s*[,)]", src):
    main_roots.update(by_short.get(m.group(1), []))
for m in re.finditer(r"\.bind\([^,]+,\s*([A-Za-z_][A-Za-z0-9_]*)\s*[,)]", src):
    main_roots.update(by_short.get(m.group(1), []))
main_roots = {q for q in main_roots if q not in worker_roots}
# a nested closure runs on whatever thread its parent runs on
for qn in list(main_roots):
    for child in FUNCS:
        if PARENT.get(child) == qn.split("#")[0]:
            main_roots.add(child)

# propagate main-ness downward, but STOP at the worker boundary
main = set(main_roots)
changed = True
while changed:
    changed = False
    for qn in list(main):
        for t in calls.get(qn, ()):
            if t in worker_roots:
                continue
            if t not in main:
                main.add(t)
                changed = True

# ------------------------------------------------------------------- reporting
flagged = []
for qn, hits in direct.items():
    if qn not in main:
        continue
    # DUAL-PATH IS NOT A PASS. Excusing anything also reachable from a worker was a
    # real blind spot: once update_since_date moved to a worker (v4.2.96),
    # import_from_folder became worker-reachable and vanished from this report --
    # while full_update still walks it on the main thread. A function that can run on
    # either thread is reported, and labelled, precisely because that is the shape
    # main-thread work drifts back into.
    dual = qn in worker or qn in worker_roots
    sev = "WALK" if hits["walk"] else ("ITEM" if hits["item"] else "ENTRY")
    flagged.append((sev, qn, hits, dual))

order = {"WALK": 0, "ITEM": 1, "ENTRY": 2}
flagged.sort(key=lambda x: (order[x[0]], x[1]))

print("=" * 78)
print("MAIN-THREAD COM AUDIT  --  %s" % os.path.basename(APP))
print("=" * 78)
print("functions parsed        : %d" % len(FUNCS))
print("touch COM directly      : %d" % len(direct))
print("worker-side (safe)      : %d" % len([q for q in direct if q in worker or q in worker_roots]))
print("main-thread reachable   : %d  (%d of them dual-path)"
      % (len(flagged), sum(1 for f in flagged if f[3])))
print()

for sev in ("WALK", "ITEM", "ENTRY"):
    grp = [f for f in flagged if f[0] == sev]
    if not grp:
        continue
    title = {"WALK": "FOLDER WALKS / BULK COM  (freeze risk -- the real problem)",
             "ITEM": "SINGLE-ITEM COM  (bounded; the documented open/display pattern)",
             "ENTRY": "COM HANDLE ACQUISITION ONLY"}[sev]
    print("-" * 78)
    print("%s  [%d]" % (title, len(grp)))
    print("-" * 78)
    for _, qn, hits, dual in grp:
        toks = sorted(set(hits["walk"] + hits["item"] + hits["entry"]))
        print("  %s:%d%s" % (qn, LINE[qn], "   [DUAL-PATH: also runs on a worker]" if dual else ""))
        print("      %s" % ", ".join(toks[:8]))
    print()

print("%d finding(s). FOLDER WALKS are the ones that freeze the UI." % len(flagged))
sys.exit(1 if any(f[0] == "WALK" for f in flagged) else 0)
