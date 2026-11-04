"""Session 5 edits (Nov 4, 2026): consistency fixes, book-style column headers, font sizes, new cover.
Reference only: it stops if the edits are already applied.
usage: python edit5.py <in.docx> <out.docx> <frame_media_dir>
"""
import sys, copy, os
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from docx.shared import Emu
from lxml import etree
from dlib import *

SRC, DST, MEDIA = sys.argv[1], sys.argv[2], sys.argv[3]
d = Document(SRC)
body = d.element.body
if "POSTS AND TELECOMMUNICATIONS" in el_text(body):
    sys.exit("edit5 already applied")
R = Repl(body, "edit5")

def tables():
    return body.findall(q("tbl"))
def first_text(t):
    return cell_text(cells(rows(t)[0])[0]).strip()
def by_first(prefix):
    return [t for t in tables() if first_text(t).startswith(prefix)]
def containing(s):
    return [t for t in tables() if s in el_text(t)]
def one(lst):
    assert len(lst) == 1, len(lst)
    return lst[0]
def runs_with_text(el):
    return [r for r in el.iter(q("r")) if (r.find(q("t")) is not None and el_text(r).strip())]
RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
             "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
             "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
def rpr_get(rpr, tag):
    x = rpr.find(q(tag))
    if x is None:
        x = etree.Element(q(tag))
        later = RPR_ORDER[RPR_ORDER.index(tag) + 1:]
        nxt = next((c for c in rpr if c.tag.split("}")[1] in later), None)
        if nxt is not None: nxt.addprevious(x)
        else: rpr.append(x)
    return x
def cur_sz(r):
    rp = r.find(q("rPr")); s = rp.find(q("sz")) if rp is not None else None
    return s.get(q("val")) if s is not None else None
def set_sz(el, val, keep=("32",), only=None):
    """Set the size of every text run in el (half-points); skip title-bar runs; with only=, change just runs whose size is in it (None = inherited)."""
    for r in runs_with_text(el):
        c = cur_sz(r)
        if c in keep or (only is not None and c not in only):
            continue
        rpr = r.find(q("rPr"))
        if rpr is None:
            rpr = etree.Element(q("rPr")); r.insert(0, rpr)
        for tag in ("sz", "szCs"):
            rpr_get(rpr, tag).set(q("val"), str(val))
def title_row(t):
    for tr in rows(t):
        if "Project Title" in el_text(tr):
            return tr
    return None
def set_par_text(p, text):
    rpr = first_rpr(p)
    for r in p.findall(q("r")):
        p.remove(r)
    p.append(dlib_mk_run(rpr, text))
from dlib import _mk_run as dlib_mk_run
def note_after(el, text, tmpl_par):
    p = copy.deepcopy(tmpl_par)
    set_par_text(p, text)
    el.addnext(p)
    return p
def cell(tr, i):
    return cells(tr)[i]
def set_cell(tr, i, text):
    set_cell_text(cells(tr)[i], text)

# ---------------------------------------------------------------- 1. PM authority threshold: "up to ... and"
R("approve scope and schedule changes below 2% of budget (15,140,000 VND) or 3 days of schedule. Larger changes must be submitted to the Sponsor for approval.",
  "approve a change that costs up to 2% of the budget (15,140,000 VND) and delays the schedule by up to 3 days. Larger changes must be submitted to the Sponsor for approval.")
R("budget (variance above 2%) and schedule (more than 3 days)", "budget (more than 2%, 15,140,000 VND) and schedule (more than 3 days)")
R("Change control: the PM approves changes below 2% of the budget (15,140,000 VND) or 3 days of schedule; larger changes go to the Sponsor.",
  "Change control: the PM approves changes of up to 2% of the budget (15,140,000 VND) and up to 3 days of schedule; larger changes go to the Sponsor.")
R("The PM approves changes under 2% of budget or 3 days of schedule; larger changes go to the Sponsor",
  "The PM approves changes of up to 2% of the budget and up to 3 days of schedule; larger changes go to the Sponsor")
R("within the PM authority (under 2% of budget or 3 days)", "within the PM authority (up to 2% of the budget and up to 3 days)")
R("the PM approves under 2% of budget or 3 days, the Sponsor above.", "the PM approves up to 2% of the budget and up to 3 days, the Sponsor anything larger.")
R("The PM approves changes under 2% of the budget or 3 days of schedule; larger changes are submitted to the Sponsor.",
  "The PM approves changes of up to 2% of the budget and up to 3 days of schedule; larger changes are submitted to the Sponsor.")
R("the PM approves changes under 2% of the budget (15,140,000 VND) or 3 days, the Sponsor approves larger ones.",
  "the PM approves changes of up to 2% of the budget (15,140,000 VND) and up to 3 days, the Sponsor approves larger ones.")

# ---------------------------------------------------------------- 2. defects at the release candidate: S1 or S2
R("resolve 100% of critical bugs before the release candidate", "resolve every critical (S1) and high (S2) defect before the release candidate")
R("no open critical defect at the release candidate (1.5.8)", "no open critical (S1) or high (S2) defect at the release candidate (1.5.8)")
R("No open critical or high defect at the release candidate (1.5.8)", "No open critical (S1) or high (S2) defect at the release candidate (1.5.8)")
R("8. Defects: 100% of critical defects resolved; no open critical or high defect at the release candidate (Charter).",
  "8. Defects: no open critical (S1) or high (S2) defect at the release candidate (Charter).")
R("Crash-free rate >= 99%; API p90 latency < 200 ms confirmed under load; 100% of critical bugs resolved (Charter Other success criterion).",
  "Load test with 15,000 virtual users: error rate < 1% and API p90 latency < 200 ms; no open critical (S1) or high (S2) defect at the release candidate (Charter Other success criterion).")

# ---------------------------------------------------------------- 3. account deletion in FR-02
R("Profile Management: update the profile picture, bio and password.", "Profile Management: update the profile picture, bio and password; delete the account and its personal data.")
R("Profile Management: update profile picture, bio, password", "Profile Management: update profile picture, bio, password; delete account")
R("Profile changes are saved and shown after re-login", "Profile changes are saved and shown after re-login; a deleted account cannot log in and its personal data is erased", 2)
R("account (login, profile), discovery", "account (login, profile, account deletion), discovery")
R("in-app account and data deletion,", "in-app account and data deletion (FR-02),")

# ---------------------------------------------------------------- 5. early checks in 1.3.5
R("FR-04 and FR-09 are code-complete before testing starts.",
  "FR-04 and FR-09 are code-complete before testing starts; also covers the early checks of R-02, R-08 and R-11 with free tools.", 2)
R("Developers write the unit tests of their own code; the QA engineer tests the integrated MVP.",
  "Developers write the unit tests of their own code; the QA engineer tests the integrated MVP and, one week before 1.5.3, runs early checks: "
  "a 1,000-connection WebSocket test on staging, an OWASP ZAP baseline scan and launch-time measurements on the two slowest reference devices (Risk Register R-02, R-08, R-11).")

# ---------------------------------------------------------------- small wording fixes
R("rating/comment/mark as favourite", "rating, comments and favorites")
R("rating, comment and mark as favourite", "rating, comments and favorites")
R("Server support at least 9,000 concurrent users, up to 15,000 users without performance degradation.",
  "The server supports at least 9,000 concurrent users at launch and up to 15,000 with an error rate < 1%.")
R("up to 15,000 without performance degradation.", "up to 15,000 with an error rate < 1%.")            # NFR-01
R("(up to 15,000 without performance degradation)", "(up to 15,000 with an error rate < 1%)")         # Scope Statement
R("1. Concurrent users: at least 9,000 at launch and 15,000 without performance degradation (NFR-01).",
  "1. Concurrent users: at least 9,000 at launch and up to 15,000 with an error rate < 1% (NFR-01).")
R("Server supports >= 9,000 concurrent users, up to 15,000 without degradation", "Server supports >= 9,000 concurrent users, up to 15,000 with error rate < 1%")
R("(15,000 without degradation)", "(15,000 with error rate < 1%)")
R("UI/UX designing standards based on Apple Human Interface Guidelines and Google Material Design 3, with UAT satisfaction score >= 4.2/5.",
  "UI/UX follows Apple Human Interface Guidelines and Google Material Design 3, with a UAT satisfaction score >= 4.2/5.")
R("their unavailability on the scheduled dates (Assumption Log A-01 to A-12).", "their unavailability on the scheduled dates (Assumption Log A-01 to A-12; Risk Register R-01).")
R("requiring additional load testing or architecture changes.", "requiring additional load testing or architecture changes (R-02).")
R("if development slips, Could-have functions are de-scoped first.", "if development slips, Could-have functions are de-scoped first (R-03).")
R("Schedule delays due to mid-project requirement changes.", "Schedule delays due to mid-project requirement changes (R-04).")
R("All summary milestones below are met by their due dates; the absolute schedule variance must not exceed 14 days (2 weeks) from the baseline schedule.",
  "Every summary milestone below is reached no more than 14 days (2 weeks) after its baseline due date.")
R("have been developed and tested to requirements.", "have been developed and pass their RTM verification.")
R("Recommendations reflect stated inputs; effectiveness tracked via likes/skips/plays",
  "Every recommended song matches a genre or artist in the user's last 50 plays or was played by similar listeners; likes, skips and plays are logged", 2)
R("1. k6 load test (1.5.3) on the production servers (PR-01, leased from Oct 13) before they are deployed for go-live in 1.6.4 (staging is too small to be representative): 15,000 virtual users: 5-minute ramp-up, then 15 minutes at full load, with error rate < 1% and p90 latency < 200 ms.",
  "1. k6 load test in 1.5.3 on the production servers (PR-01) before go-live, since staging is smaller: 15,000 virtual users, 5-minute ramp-up, 15 minutes at full load; error rate < 1% and p90 < 200 ms.")
R("12. Project documents: no open major review comment when a document is sent to the Customer or Sponsor.",
  "12. Project documents: every comment marked \"must fix\" in the peer review is closed before a document is sent to the Customer or Sponsor.")
R("the peer-review checklist has no open major comment", "every \"must fix\" comment of the peer review is closed")
R("S4 Low: fixed if time allows, otherwise moved to the warranty backlog.", "S4 Low: fixed in the next fix work package that has hours left, otherwise moved to the warranty backlog.")
R("iPhone 11 64GB (used) - oldest supported iOS device", "iPhone 11 64GB (used) - older, slower iOS device")
R("about 24 TB at 500 VND/GB", "about 22 TB at 550 VND/GB including VAT", 2)
R("phone the Sponsor that day;", "phone the Customer or the Sponsor that day;")
R("Bought in 1.7.3; active within", "Leased in 1.7.3; active within")
R("3. Submit on Jan 8, the first day of 1.6.6;", "3. Submit on Jan 8, the first day of 1.6.6, so that a 3-working-day review ends by Jan 13, one day before the A-13 window closes;")
R("27/08/2026", "August 27, 2026")
R("29/09/2026", "September 29, 2026", 5)
# repeated rate sentence in the 7 WBS Dictionary forms (rates are in the cost table)
R(" Labor rates in VND/hour (2 years of experience, Basis of Estimates): Minh 206,067; Đại 118,566; Hiếu 175,001; A 124,261; Thu 198,818.", "", 7)

# Requirements Documentation / RTM: FR-01 and NFR-12 also get a functional test
for title, col in (("REQUIREMENTS DOCUMENTATION", 6), ("REQUIREMENTS TRACEABILITY MATRIX", 8)):
    for rid in ("FR-01", "NFR-06", "NFR-12"):
        for t in by_first(title):
            try:
                tr = find_row(t, rid)
            except KeyError:
                continue
            assert cell_text(cells(tr)[col]).strip() == "UAT", (title, rid)
            set_cell(tr, col, "Design review + UAT" if rid == "NFR-06" else "Functional test + UAT")

# ---------------------------------------------------------------- review pass: measurable criteria, wording
R("UAT + analytics review", "Functional test + UAT", 2)                         # FR-08: needs a test case before UAT (QMP objective 7)
R("Users join premiere, exchange chat, and stay in sync in real time under concurrent load",
  "Users join and chat; playback stays within 2 s of the premiere clock and chat messages arrive within 1 s with 15,000 concurrent users", 2)
R("k6 Cloud for 1 month (Nov 12-20).", "k6 Cloud subscription for 1 month, tests on Nov 12-20.", 2)
R("Plan-Do-Check-Act each phase: the lessons-learned review (30 minutes, whole team) decides what to change in the checklists, Definition of Done or test cases for the next phase; the change is checked at the next audit.",
  "Plan-Do-Check-Act: each lessons-learned review updates the checklists, Definition of Done and test cases; the next audit checks the change.")
R("Every commit (from Oct 20, when the pipeline is set up in 1.2.7): the CI pipeline runs dart analyze and ESLint, all unit tests and the coverage check; the build fails if any step fails. "
  "Before Oct 20 (backend work starts Oct 7 in 1.3.1 and 1.3.3), developers run the same checks locally before each pull request, and all branches are rerun through CI once it is live.",
  "Every commit (from Oct 20, when 1.2.7 sets up the pipeline): CI runs dart analyze, ESLint, all unit tests and the coverage check, and fails the build if any step fails. "
  "Before Oct 20 (backend work starts Oct 7), developers run the same checks locally and rerun their branches through CI once it is live.")
R("shows where rework is concentrated. A result that misses a metric is handled by the nonconformance procedure below.",
  "shows where rework is concentrated; a missed metric goes through the nonconformance procedure below.")
R("An action that changes scope, schedule or cost goes through change control (1.7.5): the PM approves up to 2% of the budget and up to 3 days, the Sponsor anything larger.",
  "An action that changes scope, schedule or cost goes through change control (1.7.5) within the approval limits of the Charter.")
R("Quality audits: the Project Manager and the QA engineer audit each phase before its review with the Customer or Sponsor, i.e. before 1.2.8 (Oct 9), 1.3.6 (Nov 12), 1.4.6 (Nov 30), 1.5.6 (Dec 14) and 1.6.10 (Jan 6). "
  "Each audit (about 2 hours) checks that the planned process was followed (reviews done, tests run, defects logged) and that the deliverables meet their acceptance criteria in the WBS Dictionary. Findings are logged with an owner and a due date.",
  "Quality audits: before each phase review with the Customer or Sponsor (1.2.8 Oct 9, 1.3.6 Nov 12, 1.4.6 Nov 30, 1.5.6 Dec 14, 1.6.10 Jan 6), the Project Manager and the QA engineer "
  "spend about 2 hours checking that the planned process was followed (reviews done, tests run, defects logged) and that the deliverables meet their WBS Dictionary acceptance criteria; each finding gets an owner and a due date.")
R("Held at the phase ends listed in the Quality Management Approach, using the audit checklist. The audit report lists the checked items, findings and actions; findings are closed within 5 working days and summarized in the next monthly report.",
  "Held on the dates above with the audit checklist; the report lists the checked items, findings and actions; findings close within 5 working days.")
# shorter assumptions (same text in Activity Cost Estimates and Resource Requirements)
R("Built in parallel with the SRS from the requirements collected in 1.1.1; the final trace is checked against the SRS in 1.1.7.",
  "Built in parallel with the SRS from the 1.1.1 requirements; checked against the SRS in 1.1.7.", 2)
R("The review changes existing requirements only, with no new requirement; sign-off on Sep 18 (A-03).",
  "The review changes existing requirements only; sign-off on Sep 18 (A-03).", 2)
R("Written after the RTM (1.1.3), from its acceptance criteria and the Charter's high-level requirements; adopted in the Quality Management Plan (1.7.1).",
  "Written from the RTM acceptance criteria (1.1.3); adopted in the Quality Management Plan (1.7.1).", 2)
R("Lessons learned are collected in a 30-minute review at each phase end; 1.7.6 compiles them into the final project report.",
  "Compiles the 30-minute lessons-learned reviews held at each phase end into the final report.", 2)
for p in body.findall(q("p")):
    if p_text(p) == "Overall Project Risk":
        assert replace_in_p(p, "Overall Project Risk", "Overall Project Risk:") == 1

# Charter approvals: Sponsor name
t = one(containing("Sponsor or Originator Name"))
pm_cell, sp_cell = cells(rows(t)[0])
newp = copy.deepcopy(pm_cell.find(q("p")))
assert replace_in_p(newp, "Project Manager Name", "Sponsor or Originator Name") == 1
assert replace_in_p(newp, "Trần Vũ Tiến Minh", "Nguyễn Trung B") == 1
for p in sp_cell.findall(q("p")):
    sp_cell.remove(p)
sp_cell.append(newp)

# ---------------------------------------------------------------- dates prepared
for t in by_first("WBS DICTIONARY"):
    n = replace_all(title_row(t), "October 8, 2026", "September 7, 2026 (revised October 8, 2026)")
    assert n == 1
rr = by_first("RISK REGISTER")
assert len(rr) == 4
for t in rr:
    assert replace_all(title_row(t), "September 29, 2026", "September 29, 2026 (revised October 8, 2026)") == 1

# ---------------------------------------------------------------- 4. book-style column headers, formulas in notes
ace_note = None
for p in body.findall(q("p")):
    if p_text(p).startswith("Milestones (1.1.8"):
        ace_note = p
assert ace_note is not None
# Bottom-up worksheet headers first ("Labor Hours (A)" contains "Hours (A)")
for old in ("Labor Hours (A)", "Labor Rate (B)", "Total Labor (C = A x B)", "Material (D)", "Supplies (E)", "Equipment (F)", "Travel (G)",
            "Other Direct Costs (H)", "Indirect Costs (I)", "Reserve (J)", "Estimate (K = C + D + E + F + G + H + I + J)"):
    R(old, old.split(" (")[0])
# WBS Dictionary
for old, new in (("Hours (A)", "Hours"), ("Rate (B)", "Rate"), ("Total (C = A x B)", "Total"), ("Units (D)", "Units"), ("Cost (E)", "Cost"),
                 ("Total (F = D x E)", "Total"), ("Total Cost (G = C + F)", "Total Cost"), ("Total (SUM)", "Total")):
    R(old, new, 7)
for t in [x for x in by_first("WBS DICTIONARY") if "RESOURCE SUMMARY" not in first_text(x)]:
    qr = next(tr for tr in rows(t) if cell_text(cells(tr)[0]).startswith("Quality Requirements"))
    nr = copy.deepcopy(qr)
    set_cell_text(cells(nr)[0], "Labor total = hours x rate (VND/hour, Basis of Estimates); material total = units x unit cost; total cost = labor total + material total.")
    for r in nr.iter(q("r")):
        rp = r.find(q("rPr"))
        if rp is not None:
            for b in rp.findall(q("b")) + rp.findall(q("bCs")):
                rp.remove(b)
            rpr_get(rp, "i")
    qr.addprevious(nr)
R("Total Hours (SUM)", "Total Hours")
R("Labor Cost (VND) (= Total Hours x rate)", "Labor Cost (VND)")
R("Cost Baseline Reconciliation: Work-package cost", "Cost Baseline Reconciliation: Labor cost of each member = total hours x the member's rate. Work-package cost")
# Activity Cost Estimates
for old in ("Labor Costs (A = hours x rate)", "Physical Costs (B)", "Reserve (C = (A + B) x phase %)", "Estimate (D = A + B + C)", "Range (O = (A + B) x 0.9 to P = D)"):
    R(old, old.split(" (")[0])
set_par_text(ace_note,
    "Labor costs = hours x member labor rate (Basis of Estimates); reserve = (labor + physical costs) x the phase overrun % (1.1 and 1.7: 15%; 1.2: 20%; "
    "1.3 and 1.4: 40%; 1.5: 35%; 1.6: 30%); estimate = labor + physical costs + reserve; range = from O = (labor + physical costs) x 0.9 to P = estimate. "
    "Milestones (1.1.8, 1.2.11, 1.3.8, 1.4.8, 1.5.8, 1.6.3, 1.6.12, 1.7.2, 1.7.9) are 0-hour events and carry no cost. Resources: Minh = Trần Vũ Tiến Minh (Project Manager), "
    "Đại = Nguyễn Phú Đại (Business Analyst / Full-Stack Developer), Hiếu = Trần Đình Hiếu (Backend Developer / Technical Lead), A = Trần Văn A (QA Engineer / DevOps Engineer), "
    "Thu = Lê Quang Thu (UI/UX Designer / Flutter Developer).")
# Cost Estimating Worksheet
for old, new in (("Cost Estimate (= Cost per Unit x Number of Units)", "Cost Estimate"), ("Multiplier (current / previous)", "Multiplier"),
                 ("Cost Estimate (= Previous Cost x Multiplier)", "Cost Estimate"), ("Optimistic Cost (O = M x 0.9)", "Optimistic Cost"),
                 ("Most Likely Cost (M)", "Most Likely Cost"), ("Pessimistic Cost (P = M x (1 + phase %))", "Pessimistic Cost"),
                 ("Expected Cost Estimate (= (O + 4M + P) / 6)", "Expected Cost Estimate"),
                 ("  (work package values summed per phase; each work package uses its phase range)", "")):
    R(old, new)
cew = one(by_first("COST ESTIMATING WORKSHEET"))
note_after(cew, "Parametric: cost estimate = cost per unit x number of units. Analogous: multiplier = current size / previous size, times the labor-rate ratio when another member "
           "does the work; cost estimate = previous cost x multiplier. Three-point (per phase, summed from its work packages): optimistic O = M x 0.9; most likely M = bottom-up cost; "
           "pessimistic P = M x (1 + phase overrun %); expected = (O + 4M + P) / 6.", ace_note)
# Bottom-up worksheet: headers, k6 and store fees to Supplies
bu = one(by_first("BOTTOM-UP COST ESTIMATING WORKSHEET"))
for rid, sup, oth in (("1.5.3", "17,600,000", "-"), ("1.5 total", "17,600,000", "-"), ("1.6.6", "3,300,000", "-"), ("1.6 total", "3,300,000", "-"),
                      ("WBS total", "20,900,000", "55,500,000"), ("Cost baseline", "20,900,000", "55,500,000"), ("Budget", "20,900,000", "55,500,000")):
    tr = find_row(bu, rid)
    set_cell(tr, 5, sup); set_cell(tr, 8, oth)
bu_note = next(p for p in body.findall(q("p")) if p_text(p).startswith("Budget = cost baseline 720,921,916"))
set_par_text(bu_note, "Total labor = labor hours x labor rate; estimate = total labor + material + supplies + equipment + travel + other direct costs + indirect costs + reserve. "
             "Supplies: load-testing subscription (1.5.3) and store accounts (1.6.6). Equipment: the 4 reference devices (1.5.3). Other direct costs: the monthly cloud service "
             "bill of 1.2.7 (servers, database, storage and CDN; PR-01 to PR-04). Budget = cost baseline 720,921,916 + management reserve 36,046,096 = 756,968,012 VND, "
             "rounded up to 757,000,000 VND.")
# Basis of Estimates
R("Labor rate: ITviec median salary for 1-2 years of experience [R1] + 21.5% employer insurance [R2]; 176h/month",
  "Labor rate = ITviec median salary for 1-2 years of experience [R1] x 1.215 (21.5% employer insurance [R2]) / 176 h per month; labor cost = cost per hour x WBS hours")
for old, new in (("Salary / month (A)", "Salary / month"), ("Employer cost (B = A x 1.215)", "Employer cost / month"), ("Cost / hour (C = B / 176)", "Cost / hour"),
                 ("WBS hours (D)", "WBS hours"), ("Labor cost (E = C x D)", "Labor cost")):
    R(old, new)
# Risk Register
for old, new in (("Probability (A)", "Probability"), ("Score (C = A x MAX(B1:B4))", "Score"), ("Scope (B1)", "Scope"), ("Quality (B2)", "Quality"),
                 ("Schedule (B3)", "Schedule"), ("Cost (B4)", "Cost"), ("Revised Probability (D)", "Revised Probability"),
                 ("Revised Score (F = D x MAX(E1:E4))", "Revised Score"), ("Scope (E1)", "Scope"), ("Quality (E2)", "Quality"),
                 ("Schedule (E3)", "Schedule"), ("Cost (E4)", "Cost")):
    R(old, new, 4)
for old, new in (("Probability (A, D)", "Probability"), ("Scope (B1, E1)", "Scope"), ("Quality (B2, E2)", "Quality"),
                 ("Schedule (B3, E3): delay", "Schedule: delay"), ("Cost (B4, E4): extra cost", "Cost: extra cost"),
                 ("Score C = A x MAX(B1:B4) and revised score F = D x MAX(E1:E4), from 1 to 25.",
                  "Score = Probability x the highest of the four Impact ratings; Revised Score is calculated the same way from the revised ratings; both range from 1 to 25.")):
    R(old, new)
lowers = []
for t in rr:
    lt = t.getnext()
    while lt.tag != q("tbl"):
        lt = lt.getnext()
    lowers.append(lt)
    hdr = rows(lt)[0]
    c = next(c for c in cells(hdr) if cell_text(c).strip() == "Impact")
    set_cell_text(c, "Revised Impact")
for lt in lowers[:3]:
    note_after(lt, "Score = Probability x the highest of the four Impact ratings; Revised Score is calculated the same way from the revised ratings. "
               "Rating scales: Probability and Impact Scale on page 4.", ace_note)

# ---------------------------------------------------------------- 6. Inter-RTM: TRs for NFR-12 and FR-11, renumbered
irtm = one(by_first("INTER-REQUIREMENTS TRACEABILITY MATRIX"))
def tr_row(tid):
    for tr in rows(irtm):
        cs = cells(tr)
        if len(cs) > 4 and cell_text(cs[4]).strip() == tid:
            return tr
    raise KeyError(tid)
tr05, tr08 = tr_row("TR-05"), tr_row("TR-08")
n1 = copy.deepcopy(tr05)
set_cell(n1, 5, "Vietnamese and English GUI strings kept in Flutter localization (ARB) files; the selected language is stored on the device (NFR-12)")
set_cell(n1, 6, "Must-have"); set_cell(n1, 7, "Customer")
tr05.addnext(n1)
n2 = copy.deepcopy(tr08)
set_cell(n2, 5, "Favorites, ratings, comments and listening-history REST endpoints under /api/v1, stored in PostgreSQL (FR-11)")
set_cell(n2, 6, "Should-have"); set_cell(n2, 7, "Project Manager")
tr_row("TR-07").addnext(n2)
k = 0
for tr in rows(irtm):
    cs = cells(tr)
    if len(cs) > 4 and cell_text(cs[4]).strip().startswith("TR-"):
        k += 1
        set_cell(tr, 4, f"TR-{k:02d}")
assert k == 18, k

# ---------------------------------------------------------------- 7. font sizes
# Project Title / Date Prepared rows: 8 pt everywhere
for t in tables():
    tr = title_row(t)
    if tr is not None:
        set_sz(tr, 16)
for p in body.findall(q("p")):
    if p_text(p).startswith("Project Title:"):
        set_sz(p, 16)                                   # Charter
set_sz(rows(one(containing("Project Sponsor: PTIT")))[0], 16)
set_sz(rows(one(containing("Project Sponsor: PTIT")))[1], 16)
# QMP, Team Charter, Resource Requirements, cost sub-tables: body text 7.5 pt
for prefix in ("QUALITY MANAGEMENT PLAN", "TEAM CHARTER", "RESOURCE REQUIREMENTS", "COST ESTIMATING WORKSHEET"):
    t = one(by_first(prefix))
    for tr in rows(t):
        if tr is title_row(t) or el_text(tr).strip() == prefix:
            continue
        set_sz(tr, 15, only=(None, "16", "17"))
for t in tables():
    ft = first_text(t)
    if ft.startswith(("Labor rate =", "Project budget (bottom-up)", "1.2.7 Cloud", "1.5.3 Reference", "1.5.3 Load-testing", "References (accessed")):
        set_sz(rows(t)[0], 15, only=(None, "16", "17"))
# ACE and Bottom-up header rows: at least 7 pt
for prefix in ("ACTIVITY COST ESTIMATES", "BOTTOM-UP COST ESTIMATING WORKSHEET"):
    t = one(by_first(prefix))
    for r in t.iter(q("r")):
        rp = r.find(q("rPr")); s = rp.find(q("sz")) if rp is not None else None
        if s is not None and s.get(q("val")) == "13":
            s.set(q("val"), "14")
            cs = rp.find(q("szCs"))
            if cs is not None: cs.set(q("val"), "14")
for p in body.findall(q("p")):
    if p_text(p).startswith("Indirect costs: none"):
        set_sz(p, 14)

# ---------------------------------------------------------------- cover (layout of frame.docx, English, Arial)
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
rid_frame, _ = d.part.get_or_add_image(os.path.join(MEDIA, "image1.png"))
rid_logo, _ = d.part.get_or_add_image(os.path.join(MEDIA, "image2.png"))
EMU = 914400
def pic_xml(rid, cx, cy, pid, name):
    return (f'<a:graphic xmlns:a="{NS_A}"><a:graphicData uri="{NS_PIC}"><pic:pic xmlns:pic="{NS_PIC}"><pic:nvPicPr><pic:cNvPr id="{pid}" name="{name}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>'
            f'<pic:blipFill><a:blip xmlns:r="{NS_R}" r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>')
fw = int(7.25 * EMU); fh = int(fw * 684 / 502)
anchor = (f'<w:r xmlns:w="{W_NS}"><w:drawing><wp:anchor xmlns:wp="{NS_WP}" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658240" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
          f'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>{int((11 * EMU - fh) / 2)}</wp:posOffset></wp:positionV>'
          f'<wp:extent cx="{fw}" cy="{fh}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="9001" name="Cover frame"/><wp:cNvGraphicFramePr/>'
          + pic_xml(rid_frame, fw, fh, 9001, "frame.png") + '</wp:anchor></w:drawing></w:r>')
lw = int(1.55 * EMU); lh = int(lw * 613 / 483)
inline = (f'<w:r xmlns:w="{W_NS}"><w:drawing><wp:inline xmlns:wp="{NS_WP}" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{lw}" cy="{lh}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
          f'<wp:docPr id="9002" name="PTIT logo"/><wp:cNvGraphicFramePr/>' + pic_xml(rid_logo, lw, lh, 9002, "logo.png") + '</wp:inline></w:drawing></w:r>')
def cpar(text="", size=26, bold=True, before=0, after=0, align="center", extra_runs=()):
    p = etree.Element(q("p"))
    ppr = etree.SubElement(p, q("pPr"))
    sp = etree.SubElement(ppr, q("spacing")); sp.set(q("before"), str(before)); sp.set(q("after"), str(after)); sp.set(q("line"), "276"); sp.set(q("lineRule"), "auto")
    jc = etree.SubElement(ppr, q("jc")); jc.set(q("val"), align)
    for xml in extra_runs:
        p.append(etree.fromstring(xml))
    if text:
        r = etree.SubElement(p, q("r")); rp = etree.SubElement(r, q("rPr"))
        f = etree.SubElement(rp, q("rFonts"))
        for a in ("ascii", "hAnsi", "cs", "eastAsia"): f.set(q(a), "Arial")
        if bold: etree.SubElement(rp, q("b"))
        for tag in ("sz", "szCs"): etree.SubElement(rp, q(tag)).set(q("val"), str(size))
        t = etree.SubElement(r, q("t")); t.text = text; t.set(XML_SPACE, "preserve")
    return p
def cover_table():
    rows_ = [("Class:", "01 - Group 04", ""), ("Students:", "Trần Vũ Tiến Minh", "B23DCCE067"), ("", "Nguyễn Phú Đại", "B23DCVT071"),
             ("", "Trần Đình Hiếu", "B23DCCE034"), ("", "Lê Quang Thu", "B23DCCE091"), ("Instructor:", "Dr. Đỗ Thị Bích Ngọc", "")]
    widths = [1500, 2900, 2100]
    tbl = etree.Element(q("tbl"))
    tp = etree.SubElement(tbl, q("tblPr"))
    etree.SubElement(tp, q("tblW")).attrib.update({q("w"): str(sum(widths)), q("type"): "dxa"})
    etree.SubElement(tp, q("jc")).set(q("val"), "center")
    b = etree.SubElement(tp, q("tblBorders"))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        etree.SubElement(b, q(side)).set(q("val"), "nil")
    etree.SubElement(tp, q("tblLayout")).set(q("type"), "fixed")
    g = etree.SubElement(tbl, q("tblGrid"))
    for w in widths:
        etree.SubElement(g, q("gridCol")).set(q("w"), str(w))
    for row in rows_:
        tr = etree.SubElement(tbl, q("tr"))
        for i, (txt, w) in enumerate(zip(row, widths)):
            tc = etree.SubElement(tr, q("tc"))
            tcp = etree.SubElement(tc, q("tcPr")); etree.SubElement(tcp, q("tcW")).attrib.update({q("w"): str(w), q("type"): "dxa"})
            tc.append(cpar(txt, size=24, bold=True, before=40, after=40, align="left"))
    return tbl
sect_p = next(el for el in body.iterchildren() if el.tag == q("p") and el.find(".//" + q("sectPr")) is not None)
for el in list(body.iterchildren()):
    if el is sect_p:
        break
    body.remove(el)
for r in sect_p.findall(q("r")):
    sect_p.remove(r)
new = [cpar("POSTS AND TELECOMMUNICATIONS INSTITUTE OF TECHNOLOGY", 24, before=520, extra_runs=(anchor,)),
       cpar("FACULTY OF INFORMATION TECHNOLOGY 1", 24, before=60),
       cpar("", before=600, extra_runs=(inline,)),
       cpar("PROJECT MANAGEMENT DOCUMENTS", 34, before=600),
       cpar("Development and Deployment of a Music Streaming Platform", 26, before=120),
       cpar("COURSE: Software Project Management", 24, before=120, after=360),
       cover_table(),
       cpar("Hanoi – 2026", 24, before=1100)]
for el in new:
    sect_p.addprevious(el)

def keep_next(tr):
    for p in tr.iter(q("p")):
        ppr = p.find(q("pPr"))
        if ppr is None:
            ppr = etree.Element(q("pPr")); p.insert(0, ppr)
        if ppr.find(q("keepNext")) is None:
            kn = etree.Element(q("keepNext"))
            ps = ppr.find(q("pStyle"))
            if ps is not None: ps.addnext(kn)
            else: ppr.insert(0, kn)
for t in tables():
    if first_text(t).startswith(("Labor rate =", "Project budget (bottom-up)", "1.2.7 Cloud", "1.5.3 Reference", "1.5.3 Load-testing", "References (accessed")):
        for tr in rows(t)[:2]:
            keep_next(tr)

# an empty paragraph between a form's last table and its section break only adds a blank line: remove it
for p in body.findall(q("p")):
    ppr = p.find(q("pPr"))
    if ppr is None or ppr.find(q("sectPr")) is None:
        continue
    prev = p.getprevious()
    if (prev is not None and prev.tag == q("p") and not p_text(prev).strip() and prev.find(".//" + q("drawing")) is None
            and prev.find(".//" + q("br")) is None and prev.getprevious() is not None and prev.getprevious().tag == q("tbl")):
        body.remove(prev)
# empty section-break paragraphs: 1 pt high, so a form that ends at the page bottom does not push them onto a blank page
for p in body.findall(q("p")):
    ppr = p.find(q("pPr"))
    if ppr is None or ppr.find(q("sectPr")) is None or p_text(p).strip():
        continue
    sp = ppr.find(q("spacing"))
    if sp is None:
        sp = etree.Element(q("spacing"))
        anchor = [ppr.find(q(x)) for x in ("ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "outlineLvl", "rPr", "sectPr")]
        anchor = [a for a in anchor if a is not None]
        anchor[0].addprevious(sp)
    sp.set(q("before"), "0"); sp.set(q("after"), "0"); sp.set(q("line"), "240"); sp.set(q("lineRule"), "auto")
    rpr = ppr.find(q("rPr"))
    if rpr is None:
        rpr = etree.Element(q("rPr")); ppr.find(q("sectPr")).addprevious(rpr)
    for tag in ("sz", "szCs"):
        rpr_get(rpr, tag).set(q("val"), "2")

d.save(DST)
print("ok", DST)
