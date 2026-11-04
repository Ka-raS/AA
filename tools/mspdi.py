import sys, re
from lxml import etree
NS = {"p": "http://schemas.microsoft.com/project"}

def dur_h(s):
    # PT72H0M0S
    if not s: return 0
    m = re.match(r"PT(\d+)H(\d+)M", s)
    return int(m.group(1)) + int(m.group(2)) / 60 if m else s

def load(path):
    t = etree.parse(path)
    r = t.getroot()
    g = lambda e, tag: (e.findtext("p:" + tag, namespaces=NS) or "")
    tasks = {}
    order = []
    for e in r.findall("p:Tasks/p:Task", NS):
        uid = g(e, "UID")
        preds = []
        for pl in e.findall("p:PredecessorLink", NS):
            preds.append((g(pl, "PredecessorUID"), g(pl, "Type"), g(pl, "LinkLag"), g(pl, "LagFormat")))
        tasks[uid] = dict(uid=uid, id=g(e, "ID"), wbs=g(e, "WBS"), name=g(e, "Name"), dur=dur_h(g(e, "Duration")),
                          start=g(e, "Start")[:16], finish=g(e, "Finish")[:16], ms=g(e, "Milestone"), summ=g(e, "Summary"),
                          preds=preds, notes=g(e, "Notes"), work=dur_h(g(e, "Work")), cost=g(e, "Cost"),
                          constraint=(g(e, "ConstraintType"), g(e, "ConstraintDate")[:10]))
        order.append(uid)
    res = {g(e, "UID"): g(e, "Name") for e in r.findall("p:Resources/p:Resource", NS)}
    assn = {}
    for a in r.findall("p:Assignments/p:Assignment", NS):
        tu, ru = g(a, "TaskUID"), g(a, "ResourceUID")
        if ru in res and res[ru]:
            assn.setdefault(tu, []).append(res[ru])
    return r, tasks, order, res, assn

if __name__ == "__main__":
    path = sys.argv[1]
    r, tasks, order, res, assn = load(path)
    g = lambda tag: r.findtext("p:" + tag, namespaces=NS)
    print("Name:", g("Name"), "| Title:", g("Title"), "| Start:", g("StartDate"), "| Finish:", g("FinishDate"), "| Created:", g("CreationDate"), "| LastSaved:", g("LastSaved"))
    print("Resources:", res)
    for uid in order:
        t = tasks[uid]
        pr = []
        for (pu, ty, lag, lf) in t["preds"]:
            pw = tasks[pu]["wbs"] if pu in tasks else "?" + pu
            s = pw
            if ty and ty != "1": s += {"0": "FF", "2": "SF", "3": "SS"}.get(ty, ty)
            if lag and lag != "0": s += f"+{int(lag)/4800:g}d"
            pr.append(s)
        flag = "M" if t["ms"] == "1" else ("S" if t["summ"] == "1" else " ")
        print(f"{t['wbs']:8s}{flag} {t['dur']/8 if isinstance(t['dur'], (int,float)) else t['dur']:>5}d {t['start']} -> {t['finish']} | {t['name'][:80]:80s} | pred={','.join(pr)} | res={','.join(assn.get(uid, []))} | work={t['work']}{' | C=' + str(t['constraint']) if t['constraint'][0] not in ('0','') else ''}")
