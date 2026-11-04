"""Forward-pass reschedule of an MSPDI file (ASAP, FS links with day lags, SNET constraints)
and compare with the dates stored in the file; check resource overlaps."""
import sys, datetime as dt
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, __file__.rsplit("\\", 1)[0])
from mspdi import load

r, tasks, order, res, assn = load(sys.argv[1])
NS = {"p": "http://schemas.microsoft.com/project"}
# calendar from file
hol, work_ex = set(), set()
for e in r.findall("p:Calendars/p:Calendar/p:Exceptions/p:Exception", NS):
    d = dt.date.fromisoformat(e.findtext("p:TimePeriod/p:FromDate", namespaces=NS)[:10])
    (work_ex if e.findtext("p:DayWorking", namespaces=NS) == "1" else hol).add(d)
def working(d): return (d.weekday() < 5 and d not in hol) or d in work_ex
def nxt(d):
    d += dt.timedelta(days=1)
    while not working(d): d += dt.timedelta(days=1)
    return d
def first_working(d):
    while not working(d): d += dt.timedelta(days=1)
    return d
PSTART = dt.date(2026, 8, 17)
leaf = [u for u in order if tasks[u]["summ"] != "1" and tasks[u]["wbs"]]
sched = {}
# finish represented as (date) end of day; start as date start of day
def compute(u):
    if u in sched: return sched[u]
    t = tasks[u]
    days = int(round(t["dur"] / 8)) if isinstance(t["dur"], (int, float)) else 0
    ms = t["ms"] == "1"
    es = PSTART  # earliest start of day
    for (pu, ty, lag, lf) in t["preds"]:
        if pu not in tasks or tasks[pu]["summ"] == "1":
            continue
        ps, pf, pms = compute(pu)
        L = int(int(lag or 0) / 4800)
        d = pf
        for _ in range(L): d = nxt(d)
        cand = nxt(d)  # successor starts next working day after (finish + lag)
        es = max(es, cand)
    if t["constraint"][0] == "4":
        es = max(es, first_working(dt.date.fromisoformat(t["constraint"][1])))
    es = first_working(es)
    if ms:
        # milestone sits at the end of the day before es (i.e. at the finish of the driving predecessor + lag)
        d = es - dt.timedelta(days=1)
        while not working(d): d -= dt.timedelta(days=1)
        if not t["preds"]: d = PSTART
        sched[u] = (d, d, True)
    else:
        f = es
        for _ in range(days - 1): f = nxt(f)
        sched[u] = (es, f, False)
    return sched[u]
bad = 0
for u in leaf:
    s, f, ms = compute(u)
    t = tasks[u]
    fs, ff = t["start"][:10], t["finish"][:10]
    if (s.isoformat(), f.isoformat()) != (fs, ff) and not (ms and f.isoformat() == ff):
        bad += 1
        print(f"DIFF {t['wbs']:8s} file {fs}..{ff}  computed {s}..{f}  {t['name'][:50]}")
print("date mismatches:", bad)
# overlaps per resource
per = {}
for u in leaf:
    t = tasks[u]
    if t["ms"] == "1": continue
    for p in assn.get(u, []):
        per.setdefault(p, []).append((sched[u][0], sched[u][1], t["wbs"]))
ov = 0
for p, L in per.items():
    L.sort()
    for a, b in zip(L, L[1:]):
        if b[0] <= a[1]:
            ov += 1; print("OVERLAP", p, a, b)
print("overlaps:", ov)
# monthly hours
bounds = [(dt.date(2026, 8, 17), dt.date(2026, 9, 16)), (dt.date(2026, 9, 17), dt.date(2026, 10, 16)), (dt.date(2026, 10, 17), dt.date(2026, 11, 16)), (dt.date(2026, 11, 17), dt.date(2026, 12, 16)), (dt.date(2026, 12, 17), dt.date(2027, 1, 15))]
tot = [0] * 5
for p, L in per.items():
    m = [0] * 5
    for s, f, w in L:
        d = s
        while d <= f:
            if working(d):
                for i, (a, b) in enumerate(bounds):
                    if a <= d <= b: m[i] += 8
            d += dt.timedelta(days=1)
    tot = [x + y for x, y in zip(tot, m)]
    print("MONTHS", p, m, sum(m))
print("MONTHS team", tot, sum(tot))
wd = sum(1 for i in range((dt.date(2027, 1, 15) - PSTART).days + 1) if working(PSTART + dt.timedelta(days=i)))
print("working days in project:", wd)
