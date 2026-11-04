import sys, re, datetime as dt
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, __file__.rsplit("\\", 1)[0])
from mspdi import load
from docx import Document
from dlib import *

XML, DOCX = sys.argv[1], sys.argv[2]
r, tasks, order, res, assn = load(XML)
S, sess = {}, {}
for u in order:
    t = tasks[u]
    if not t["wbs"]: continue
    if t["wbs"].count(".") == 3:
        sess.setdefault(t["wbs"].rsplit(".", 1)[0], []).append(t)
    elif t["summ"] != "1" or t["wbs"] in ("1.7.4", "1.7.5"):
        S[t["wbs"]] = t
RATE = {"Minh": 206067, "Đại": 118566, "Hiếu": 175001, "A": 124261, "Thu": 198818}
FULL = {"Minh": "Trần Vũ Tiến Minh", "Đại": "Nguyễn Phú Đại", "Hiếu": "Trần Đình Hiếu", "A": "Trần Văn A", "Thu": "Lê Quang Thu"}
INV = {v: k for k, v in FULL.items()}
who, hours = {}, {}
for w, t in S.items():
    if w in sess:
        who[w] = assn[sess[w][0]["uid"]][0]; hours[w] = sum(x["work"] for x in sess[w])
    else:
        who[w] = (assn.get(t["uid"], [""])[0]); hours[w] = t["work"]
D = lambda s: dt.date.fromisoformat(s[:10])
num = lambda s: int(s.replace(",", "")) if re.fullmatch(r"[\d,]+", s.strip() or "x") else None
P = []
def bad(m): P.append(m); print("!!", m)

d = Document(DOCX)
body = d.element.body
tbls = body.findall(q("tbl"))
def R(t):
    out = []
    for tr in rows(t):
        out.append([cell_text(c).strip() for c in cells(tr)])
    return out
def by_title(prefix):
    return [t for t in tbls if cell_text(cells(rows(t)[0])[0]).strip().startswith(prefix)]

# WBS dictionary
for t in by_title("WBS DICTIONARY"):
    if "RESOURCE SUMMARY" in cell_text(cells(rows(t)[0])[0]): continue
    for row in R(t):
        if re.fullmatch(r"1\.\d\.\d+", row[0]) and row[2] != "-":
            w = row[0]; h, rate, tot = num(row[3]), num(row[4]), num(row[5])
            if INV.get(row[2]) != who[w]: bad(f"WBSDict {w} resource")
            if h != hours[w]: bad(f"WBSDict {w} hours {h} vs {hours[w]}")
            if rate != RATE[who[w]] or h * rate != tot: bad(f"WBSDict {w} rate/total")
            if row[1] != S[w]["name"]: bad(f"WBSDict {w} name")
# WBS list
for t in by_title("WORK BREAKDOWN STRUCTURE"):
    for line in R(t)[2][0].split("\n"):
        m = re.match(r"(1(?:\.\d+)+)\.(.*)", line)
        if m and m.group(1).count(".") == 2:
            w, n = m.group(1), m.group(2).strip()
            if w not in S: bad(f"WBS {w} missing in schedule")
            elif n != S[w]["name"]: bad(f"WBS {w} '{n}' vs '{S[w]['name']}'")
wbs_ids = {w for w in S}
# Resource requirements
def rng(s, e):
    s, e = D(s), D(e)
    if s == e: return f"{s.strftime('%b')} {s.day}, {s.year}"
    if (s.month, s.year) == (e.month, e.year): return f"{s.strftime('%b')} {s.day}-{e.day}, {s.year}"
    if s.year == e.year: return f"{s.strftime('%b')} {s.day} - {e.strftime('%b')} {e.day}, {s.year}"
    return f"{s.strftime('%b')} {s.day}, {s.year} - {e.strftime('%b')} {e.day}, {e.year}"
seen = set()
for t in by_title("RESOURCE REQUIREMENTS"):
    for row in R(t):
        if re.fullmatch(r"1\.\d\.\d+", row[0]):
            w = row[0]; seen.add(w)
            m = re.search(r"x (\d+) h\s*\n?\((\d+) days?\)", row[2])
            if int(m.group(1)) != hours[w] or int(m.group(2)) * 8 != hours[w]: bad(f"ResReq {w} qty")
            if w in sess:
                for x in sess[w]:
                    s0 = D(x["start"]); txt_ = f"{s0.strftime('%b')} {s0.day}"
                    if txt_ not in row[3]: bad(f"ResReq {w} missing session date {txt_}")
            else:
                exp = "Available " + rng(S[w]["start"], S[w]["finish"])
                if not row[3].startswith(exp): bad(f"ResReq {w} '{row[3][:45]}' expected '{exp}'")
            if S[w]["name"] not in row[4]: bad(f"ResReq {w} activity text")
            if INV.get(row[1].split("(")[-1].rstrip(")").strip()) != who[w]: bad(f"ResReq {w} person")
missing = [w for w in S if S[w]["ms"] != "1" and w not in seen and w not in ("1.7.4", "1.7.5") or (w in ("1.7.4", "1.7.5") and w not in seen)]
missing = [w for w in missing if S[w]["ms"] != "1"]
if missing: bad(f"ResReq missing {missing}")
# Activity cost estimates
PCT = {"1.1": .15, "1.2": .20, "1.3": .40, "1.4": .40, "1.5": .35, "1.6": .30, "1.7": .15}
ace_rows = [row for t in by_title("ACTIVITY COST ESTIMATES") for row in R(t)]
tot_l = tot_r = tot_e = 0
for row in ace_rows:
    if re.fullmatch(r"1\.\d\.\d+", row[0]):
        w = row[0]; m = re.match(r"(\S+) (\d+)h", row[1])
        if m.group(1) != who[w] or int(m.group(2)) != hours[w]: bad(f"ACE {w} res/hours")
        lab, phys, rv, est = num(row[2]), num(row[3]) or 0, num(row[4]), num(row[5])
        if lab != hours[w] * RATE[who[w]] or rv != round((lab + phys) * PCT[w[:3]]) or est != lab + phys + rv: bad(f"ACE {w} numbers")
        tot_l += lab; tot_r += rv; tot_e += est
print("ACE totals", tot_l, tot_r, tot_e)
assert tot_l == 406794048 and tot_r == 137904988 and tot_e == 649099036
# Resource summary vs schedule months
bounds = [(dt.date(2026, 8, 17), dt.date(2026, 9, 16)), (dt.date(2026, 9, 17), dt.date(2026, 10, 16)), (dt.date(2026, 10, 17), dt.date(2026, 11, 16)), (dt.date(2026, 11, 17), dt.date(2026, 12, 16)), (dt.date(2026, 12, 17), dt.date(2027, 1, 15))]
HOL = {dt.date(2026, 9, 1), dt.date(2026, 9, 2), dt.date(2027, 1, 1)}; WS = {dt.date(2026, 9, 5)}
wk = lambda x: (x.weekday() < 5 and x not in HOL) or x in WS
mon = {k: [0] * 5 for k in RATE}
leaves = [t for w, t in S.items() if w not in sess and t["ms"] != "1"] + [x for v in sess.values() for x in v]
for t in leaves:
    p = assn[t["uid"]][0]; x = D(t["start"])
    while x <= D(t["finish"]):
        if wk(x):
            for i, (a, b) in enumerate(bounds):
                if a <= x <= b: mon[p][i] += 8
        x += dt.timedelta(days=1)
rs = R(by_title("WBS DICTIONARY: RESOURCE SUMMARY")[0])
for row in rs[3:8]:
    if [num(v) for v in row[2:7]] != mon[INV[row[0]]]: bad(f"ResSummary {row[0]}")
team = [num(v) for v in rs[8][2:7]]
if team != [sum(mon[p][i] for p in mon) for i in range(5)]: bad(f"ResSummary team {team}")
# Requirements Documentation vs RTM
rtm = {}
for t in by_title("REQUIREMENTS TRACEABILITY MATRIX"):
    for row in R(t):
        if re.fullmatch(r"N?FR-\d\d", row[0]): rtm[row[0]] = row
rd = {}
for t in by_title("REQUIREMENTS DOCUMENTATION"):
    for row in R(t):
        if re.fullmatch(r"N?FR-\d\d", row[0]): rd[row[0]] = row
assert len(rd) == 26 and len(rtm) == 26, (len(rd), len(rtm))
for k, a in rd.items():
    b = rtm[k]
    # rd: ID, Requirement, Stakeholder, Category, Priority, Acceptance, Test, Phase
    # rtm: ID, Requirement, Source, Priority, Category, BO, Deliverable, Verification, Validation
    if a[2] != b[2]: bad(f"{k} stakeholder {a[2]} vs RTM source {b[2]}")
    if a[3] != b[4]: bad(f"{k} category {a[3]} vs {b[4]}")
    if a[4] != b[3]: bad(f"{k} priority")
    if a[5].rstrip(".") != b[7].rstrip("."): bad(f"{k} acceptance '{a[5]}' vs RTM verification '{b[7]}'")
    if a[6] != b[8]: bad(f"{k} test '{a[6]}' vs '{b[8]}'")
    for wp in re.findall(r"1\.\d\.\d+", b[6]):
        if wp not in wbs_ids: bad(f"{k} RTM deliverable {wp} not a WBS item")
# leftover phrases
alltext = "\n".join(p_text(p) for p in body.iter(q("p")))
for pat in ["cost.docx", "resource.docx", "quality.docx", "Function List", "SRS Section", "one block", "Sponsor/instructor", "Prime Minister", "second National Day", " 4,400 ", "running stably", "strictly", "significantly", "Quickly", "relevant results", "minor changes", "full CRUD", "complete CRUD", "short lessons", "Page 1 of", "his own code",
            "(SUM)", "Score (C =", "Hours (A)", "(B1:B4)", "below 2%", "under 2%", "or 3 days", "favourite", "without performance degradation",
            "without degradation", "oldest supported", "open major", "if time allows", "reflect stated inputs", "in real time", "24 TB"]:
    if pat in alltext: bad(f"leftover phrase: {pat}")
# Inter-RTM: technical requirements numbered TR-01.. in table order
irtm = by_title("INTER-REQUIREMENTS TRACEABILITY MATRIX")
if irtm:
    trs = [row[4] for row in R(irtm[0]) if len(row) > 4 and re.fullmatch(r"TR-\d\d", row[4])]
    if trs != [f"TR-{i:02d}" for i in range(1, len(trs) + 1)]: bad(f"Inter-RTM TR numbering {trs}")
    print("Inter-RTM:", len(trs), "technical requirements")
# Risk register: each "RISK REGISTER" table (ID..Response) is followed by its lower table (ID, Revised..Comments)
up, low = {}, {}
for t in by_title("RISK REGISTER"):
    lt = t.getnext()
    while lt is not None and lt.tag != q("tbl"): lt = lt.getnext()
    for row in R(t):
        if re.fullmatch(r"R-\d\d", row[0]): up[row[0]] = row
    for row in R(lt):
        if re.fullmatch(r"R-\d\d", row[0]): low[row[0]] = row
if up:
    if list(up) != list(low): bad("RiskReg: upper and lower IDs differ")
    if list(up) != [f"R-{i:02d}" for i in range(1, len(up) + 1)]: bad("RiskReg: IDs not consecutive")
    for k, row in up.items():
        a, b = int(row[3]), [int(x) for x in row[4:8]]
        if not all(1 <= v <= 5 for v in [a] + b) or int(row[8]) != a * max(b): bad(f"RiskReg {k} score")
        if row[2] not in INV: bad(f"RiskReg {k} owner '{row[2]}'")
        lr = low.get(k)
        if not lr: continue
        d_, e = int(lr[1]), [int(x) for x in lr[2:6]]
        if not all(1 <= v <= 5 for v in [d_] + e) or int(lr[6]) != d_ * max(e): bad(f"RiskReg {k} revised score")
        if lr[7] not in INV: bad(f"RiskReg {k} responsible party '{lr[7]}'")
        if lr[9] not in ("Open", "Closed"): bad(f"RiskReg {k} status '{lr[9]}'")
    rng_txt = f"R-01 to R-{len(up):02d}"
    for m in re.findall(r"R-01 to R-\d\d", alltext):
        if m != rng_txt: bad(f"RiskReg: text says '{m}', register has {rng_txt}")
    print("Risk register:", len(up), "risks")
print("PROBLEMS:", len(P))
