"""Add the Risk Register (Book of Forms 2.32) as section 7 of 4.docx and update the places that refer to it.
usage: python build_rr.py <src.docx> <dst.docx> <groups, e.g. 4,4,4>"""
import sys, copy
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, r"E:\AA\tools")
from docx import Document
from lxml import etree
from dlib import *

SRC, DST = sys.argv[1], sys.argv[2]
GROUPS = [int(x) for x in sys.argv[3].split(",")]
DATE = "September 29, 2026"
SZ = "14"  # 7 pt, as in the Activity Cost Estimates and Resource Requirements

MINH, DAI, HIEU, A, THU = "Trần Vũ Tiến Minh", "Nguyễn Phú Đại", "Trần Đình Hiếu", "Trần Văn A", "Lê Quang Thu"

# ID, statement, owner, A, (B1 scope, B2 quality, B3 schedule, B4 cost), response,
# D, (E1..E4), responsible party, actions, status, comments
RISKS = [
 ("R-01",
  "If the Customer or the Sponsor is not available on a scheduled review, demo or sign-off date (A-04 to A-12), an approval may come later than planned, delaying the next work packages and milestones 1.7.2, 1.2.11, 1.6.3 and 1.6.12.",
  MINH, 3, (1, 1, 3, 2),
  "Mitigate: book each session one week ahead with a fallback date, send the documents at the start of the review lag and allow online attendance or a recorded demo (PR-12).",
  2, (1, 1, 2, 1), MINH,
  ["1. Send each invitation and document on the due date of A-04 to A-12.",
   "2. Confirm attendance 2 working days before each session.",
   "3. If a sign-off is 1 working day late, phone the Sponsor that day; a delay of more than 3 days is raised as a change request in 1.7.5."],
  "Open",
  "Charter Overall Project Risk. A-01 to A-03 were met. Trigger: no confirmation 2 working days before a session."),
 ("R-02",
  "If the Live Premiere WebSocket service (1.3.3) does not scale linearly with concurrent users, the load test in 1.5.3 (Nov 12-20) may miss the 15,000-user target (error rate < 1%, p90 < 200 ms), leading to architecture changes and rework before the release candidate (1.5.8).",
  HIEU, 4, (1, 4, 4, 3),
  "Mitigate: build the service on Socket.IO with its PostgreSQL adapter, so that a WebSocket server can be added without code changes, and test 1,000 connections on staging in 1.3.5, one week before the full load test.",
  2, (1, 3, 3, 2), DAI,
  ["1. The developer builds and unit-tests the service in 1.3.3 (Oct 7-19); the Technical Lead reviews the pull request.",
   "2. The QA engineer runs a 1,000-connection k6 test on staging in 1.3.5 (Nov 5-11).",
   "3. If 1.5.3 misses the target, add a second WebSocket server (1,975,643 VND/month before VAT) and rerun the 15,000-VU test (about 17,200,000 VND), paid from the 1.5 contingency reserve (27,811,401 VND) with the Sponsor's approval."],
  "Open",
  "Charter Overall Project Risk; Activity Cost Estimates 1.3.3 (main technical risk, +40%). Trigger: error rate >= 1% or p90 >= 200 ms in the 1.3.5 test."),
 ("R-03",
  "If the development work packages of 1.3 and 1.4 need more hours than estimated (up to +40%, Medium-Low confidence in the Activity Cost Estimates), milestone 1.4.8 (Dec 7) may slip, leading to Could-have functions being cut and less time for testing (1.5).",
  MINH, 3, (2, 2, 3, 3),
  "Mitigate: build the Must-have functions first, compare logged and remaining hours every week and, if a package is late, cut the Could-have functions FR-05, FR-10 and FR-13 first (Charter).",
  2, (2, 1, 2, 2), HIEU,
  ["1. Review the logged and remaining hours of each 1.3 / 1.4 package at the Monday team meeting.",
   "2. Pay overruns from the 1.3 (24,184,535 VND) and 1.4 (24,999,699 VND) contingency reserves.",
   "3. A package more than 3 days late goes to the next 1.7.5 review; cutting FR-05, FR-10 or FR-13 is a change request agreed with the Customer."],
  "Open",
  "Charter Overall Project Risk: 2,464 h of 4,320 h, at most 176 h per member per month. Trigger: remaining hours > remaining days x 8 h."),
 ("R-04",
  "If the Customer asks for new or changed requirements after the SRS sign-off (1.1.8, Sep 18), extra design and development work may be needed, leading to schedule delay and extra cost.",
  DAI, 3, (3, 2, 3, 3),
  "Mitigate: every request goes through change control (1.7.5) with an impact analysis; the PM approves changes under 2% of the budget (15,140,000 VND) or 3 days, the Sponsor approves larger ones.",
  2, (2, 1, 2, 2), DAI,
  ["1. Log each request in Jira as a change request.",
   "2. Write its impact on scope, hours, cost and dates within 2 working days.",
   "3. Decide at the next 1.7.5 review; after approval, update the SRS, RTM and schedule.",
   "4. Requests after 1.4.8 (Dec 7) go to the warranty backlog unless they fix a Must-have function."],
  "Open",
  "Charter Overall Project Risk. Trigger: a request at a review, demo or UAT that is not in the SRS."),
 ("R-05",
  "If the App Store or Google Play review takes longer than 3 working days or rejects a build (A-13), the apps may not be published by Jan 15, 2027, leading to a late milestone 1.6.12.",
  THU, 3, (1, 1, 3, 2),
  "Mitigate: check the release builds against the App Store Review Guidelines and Google Play policies, and complete the store listings, privacy details and in-app purchase products before submission.",
  2, (1, 1, 2, 1), THU,
  ["1. Test the release builds through TestFlight and Google Play internal testing.",
   "2. Complete the store checklist (privacy labels, data safety form, account deletion, purchase products) before Jan 8.",
   "3. Submit on Jan 8, the first day of 1.6.6; if a build is rejected, fix and resubmit it within 1 working day and ask Apple for an expedited review."],
  "Open",
  "Assumption Log A-13 (3-day lag before 1.6.12). Trigger: no store decision by Jan 13, 2027."),
 ("R-06",
  "If the cloud and CDN provider activates the purchased accounts more than 2 working days after the purchase (A-14), the setup of the repository, CI/CD, staging and CDN (1.2.7, Oct 13-20) may start late, delaying 1.3.2 and the MVP milestones.",
  A, 2, (1, 1, 2, 1),
  "Mitigate: get the activation time in writing before buying (1.7.3) and start 1.2.7 with the parts that need no cloud account, so that only the staging and CDN setup waits.",
  1, (1, 1, 2, 1), A,
  ["1. Ask Bizfly Cloud and VNG Cloud for the activation time before the purchase (Oct 2-8).",
   "2. If the accounts are not active on Oct 12, start 1.2.7 with the repository, GitHub CI pipeline and Docker images, and set up staging and the CDN last.",
   "3. Report the status at the Oct 14 review (1.7.5)."],
  "Open",
  "Assumption Log A-14 (2-day lag before 1.2.7). Trigger: accounts not active on Oct 12, 2026."),
 ("R-07",
  "If the streaming traffic after the launch exceeds the 5,000 GB/month CDN package (PR-04), overage charges may occur, leading to extra cost on 1.2.7.",
  A, 2, (1, 1, 1, 2),
  "Accept with a contingency reserve: pay the traffic above 5 TB from the 1.2.7 contingency (12,292,906 VND, about 24 TB at 500 VND/GB); keep 128 kbps as the default audio quality to limit traffic.",
  1, (1, 1, 1, 2), A,
  ["1. Set a CDN usage alert at 4,000 GB (80% of the package) each month.",
   "2. Check the CDN traffic report on every day of hypercare (1.6.5, Jan 8-15).",
   "3. Report the traffic and any overage cost in the monthly status report (1.7.4)."],
  "Open",
  "Resource Requirements PR-04 (5 TB = about 87,000 listening hours at 128 kbps). Trigger: the 4,000 GB alert."),
 ("R-08",
  "If the OWASP ZAP scan or the ASVS / MASVS checklist in 1.5.3 (Nov 12-20) finds a critical or high vulnerability, the fix may need more than the 5 days of 1.5.4, leading to a later release candidate (1.5.8) and failing NFR-05.",
  HIEU, 3, (1, 4, 3, 2),
  "Mitigate: check the security items of the code review checklist in every pull request and run an OWASP ZAP baseline scan on staging in 1.3.5, one week before 1.5.3.",
  2, (1, 3, 2, 1), A,
  ["1. Run the ZAP baseline scan on staging in 1.3.5 (Nov 5-11) and log each finding as a defect.",
   "2. Fix S1 findings within 1 working day and S2 findings within 3 working days (QMP).",
   "3. Pay extra fix hours from the 1.5 contingency reserve; a finding still open on Nov 27 goes to a root cause analysis."],
  "Open",
  "QMP quality objective 5; OWASP ASVS 5.0 Level 2 and MASVS. Trigger: a critical or high finding in any scan."),
 ("R-09",
  "If the handling of account and listening data does not meet the Vietnamese Law on Personal Data Protection (effective Jan 1, 2026), the Customer may refuse acceptance (1.6.3) or a store may reject the apps, leading to rework of the account functions.",
  DAI, 2, (3, 3, 3, 2),
  "Avoid: turn the legal duties into checkable items (consent at registration, privacy notice in Vietnamese and English, in-app account and data deletion, listening data used only for FR-08) and check them in the design review (1.2.8).",
  1, (2, 3, 2, 1), DAI,
  ["1. Map each duty of the law to a requirement (FR-01, FR-02, FR-08) and a test case before the design review (Oct 9-12).",
   "2. Raise a missing item as a change request at the Oct 14 review (1.7.5).",
   "3. Check the items in the ASVS / MASVS checklist (1.5.3) and the store checklist (1.6.6)."],
  "Open",
  "QMP Quality Standards. Trigger: a legal duty with no matching requirement or test case."),
 ("R-10",
  "If a team member is absent for more than 3 working days (illness or leave), that member's work packages may stop, because each work package has one responsible person, delaying the dependent milestones.",
  MINH, 3, (2, 1, 3, 2),
  "Mitigate: name a backup person for every work package (Team Charter) and keep the code, documents and Jira tasks updated every day, so that the backup can take over within 1 working day.",
  3, (1, 1, 2, 2), MINH,
  ["1. Publish the backup list, one backup per work package, with the Project Management Plan (Sep 29).",
   "2. Request leave 3 working days ahead (Team Charter).",
   "3. The backup uses free capacity (4,320 - 2,464 = 1,856 h), at most 176 h per member per month."],
  "Open",
  "Team Charter (one task at a time, backup person, leave rules). Trigger: sick leave, or a leave request of more than 3 days."),
 ("R-11",
  "If the Flutter apps miss the device targets on the 4 reference devices (cold start < 2 s, warm start < 1 s, battery < 6% per hour; NFR-03, NFR-04), performance rework may be needed after 1.5.3, leading to a later release candidate (1.5.8).",
  THU, 3, (1, 3, 2, 2),
  "Mitigate: measure the launch time and battery use on the slowest devices (Samsung Galaxy A07, iPhone 11) in 1.3.5, one week before 1.5.3, and fix the gaps before the device tests.",
  2, (1, 2, 2, 1), THU,
  ["1. Measure 10 cold and 10 warm starts of the release builds on the Galaxy A07 and iPhone 11 in 1.3.5 (Nov 5-11).",
   "2. Log each miss as an S2 defect (fixed within 3 working days).",
   "3. Enable Firebase Crashlytics from the first build, to keep the crash-free rate >= 99% in hypercare (1.6.5)."],
  "Open",
  "QMP quality objectives 3, 4 and 9. Trigger: a median launch time above the target in 1.3.5."),
 ("R-12",
  "If the timed restore drill on staging before go-live (1.6.4) takes longer than the 4-hour RTO or loses more than 24 hours of data (NFR-10), go-live may be postponed, leading to a late milestone 1.6.12.",
  A, 2, (1, 3, 2, 1),
  "Mitigate: use the daily automated backup of the managed PostgreSQL package (PR-02) and write and rehearse the restore steps in the operations guide (1.6.9) before the drill.",
  1, (1, 3, 2, 1), A,
  ["1. Turn on the daily backup when the database is set up in 1.2.7 (Oct 13-20).",
   "2. Write the restore steps in 1.6.9 (Dec 23-31) and rehearse them once on staging.",
   "3. Time the drill in 1.6.4 (Jan 5-7); if it takes more than 4 hours, fix the steps and repeat it before go-live."],
  "Open",
  "QMP quality objective 11. Trigger: a rehearsal longer than 3 hours."),
]
assert sum(GROUPS) == len(RISKS)

LEGEND_HDR = ["Level", "Probability (A, D)", "Scope (B1, E1)", "Quality (B2, E2)",
              "Schedule (B3, E3): delay of the affected milestone", "Cost (B4, E4): extra cost, % of the 757,000,000 VND budget"]
LEGEND = [
 ("1", "<= 10%", "No requirement changed", "No defect above S4", "< 1 day", "< 1% (< 7,570,000 VND)"),
 ("2", "11-30%", "A Could-have function (FR-05, FR-10, FR-13) changed, reduced or deferred", "S3 defects", "1-3 days (PM authority)", "1-2% (7,570,000-15,140,000 VND; PM authority)"),
 ("3", "31-50%", "A Should-have function added, changed, reduced or deferred", "An S2 defect, or a QMP quality objective missed by <= 10%", "4-7 days", "2-5% (15,140,001-37,850,000 VND; Sponsor approval)"),
 ("4", "51-70%", "A Must-have function changed with the Customer's agreement", "An S1 defect, or a QMP quality objective missed by > 10%", "8-14 days (Scope Statement limit)", "5-10% (37,850,001-75,700,000 VND)"),
 ("5", "> 70%", "A Must-have function not delivered", "Release candidate (1.5.8) or go-live (1.6.4) blocked", "> 14 days (failed exit criterion, Charter)", "> 10% (> 75,700,000 VND)"),
]
LEGEND_NOTE = ("Score C = A x MAX(B1:B4) and revised score F = D x MAX(E1:E4), from 1 to 25. "
               "15-25: response actions start before the next 1.7.5 review and the risk is reported in the monthly status report (1.7.4). "
               "6-14: checked at every 1.7.5 review. 1-5: accepted and rechecked in the monthly status report. "
               "Status: Open until the trigger date has passed or the risk can no longer occur, then Closed.")

# ---------------------------------------------------------------- helpers
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
def strip_ids(el):
    for e in el.iter():
        for a in (W14 + "paraId", W14 + "textId"):
            if a in e.attrib: del e.attrib[a]
    return el

def sub(parent, tag, **attrs):
    e = etree.SubElement(parent, q(tag))
    for k, v in attrs.items():
        e.set(q(k), str(v))
    return e

def make_p(text, bold=False, jc=None, sz=SZ, keep=False):
    p = etree.Element(q("p"))
    ppr = sub(p, "pPr")
    if keep: sub(ppr, "keepNext")
    if jc: sub(ppr, "jc", val=jc)
    rp = sub(ppr, "rPr"); sub(rp, "color", val="000000"); sub(rp, "sz", val=sz); sub(rp, "szCs", val=sz)
    if text:
        r = sub(p, "r"); rpr = sub(r, "rPr")
        if bold: sub(rpr, "b"); sub(rpr, "bCs")
        sub(rpr, "color", val="000000"); sub(rpr, "sz", val=sz); sub(rpr, "szCs", val=sz)
        t = sub(r, "t"); t.text = text; t.set(XML_SPACE, "preserve")
    return p

def make_tc(width, lines, bold=False, jc=None, span=1, vmerge=None, valign=None, fill=None, narrow=False, keep=False):
    tc = etree.Element(q("tc"))
    pr = sub(tc, "tcPr")
    sub(pr, "tcW", w=width, type="dxa")
    if span > 1: sub(pr, "gridSpan", val=span)
    if vmerge == "restart": sub(pr, "vMerge", val="restart")
    elif vmerge == "cont": sub(pr, "vMerge")
    if fill: sub(pr, "shd", val="clear", color="auto", fill=fill)
    mar = sub(pr, "tcMar"); sub(mar, "top", w=25, type="dxa")
    if narrow: sub(mar, "left", w=15, type="dxa")
    sub(mar, "bottom", w=25, type="dxa")
    if narrow: sub(mar, "right", w=15, type="dxa")
    if valign: sub(pr, "vAlign", val=valign)
    if isinstance(lines, str): lines = [lines]
    for ln in lines:
        tc.append(make_p(ln, bold=bold, jc=jc, keep=keep))
    return tc

def make_tr(tcs, header=False):
    tr = etree.Element(q("tr"))
    trpr = sub(tr, "trPr"); sub(trpr, "cantSplit")
    if header: sub(trpr, "tblHeader")
    sub(trpr, "jc", val="center")
    for c in tcs: tr.append(c)
    return tr

def make_tbl(widths):
    t = etree.Element(q("tbl"))
    t.append(copy.deepcopy(AL.find(q("tblPr"))))
    g = sub(t, "tblGrid")
    for w in widths: sub(g, "gridCol", w=w)
    assert sum(widths) == 9964, sum(widths)
    return t

def title_rows(ncols):
    r0 = strip_ids(copy.deepcopy(rows(AL)[0]))
    c = cells(r0)[0]
    c.find(q("tcPr")).find(q("gridSpan")).set(q("val"), str(ncols))
    set_cell_text(c, "RISK REGISTER")
    r1 = strip_ids(copy.deepcopy(rows(AL)[1]))
    c0, c1 = cells(r1)
    left = 7                      # Project Title over the first 7 grid columns, Date Prepared over the rest
    for c, span, w in ((c0, left, sum(UW[:left])), (c1, ncols - left, sum(UW[left:]))):
        c.find(q("tcPr")).find(q("gridSpan")).set(q("val"), str(span))
        c.find(q("tcPr")).find(q("tcW")).set(q("w"), str(w))
    set_cell_lines(c1, [[("Date Prepared:", True), (" " + DATE, False)]])
    return [r0, r1]

def hdr(w, text, **kw):
    kw.setdefault("narrow", True)
    return make_tc(w, text, bold=True, jc="center", valign="center", **kw)

def num(w, v):
    return make_tc(w, str(v), jc="center", valign="center", narrow=True)

# column widths (twips, total 9964 = Letter portrait text width); numeric columns use 15-twip side margins
UW = [450, 2400, 800, 840, 500, 570, 750, 400, 960, 0]              # upper: ID .. Response
UW[-1] = 9964 - sum(UW[:-1])
LW = [450, 840, 500, 570, 750, 400, 960, 960, 2600, 510, 0]          # lower: ID, Revised .. Comments
LW[-1] = 9964 - sum(LW[:-1])

def upper_table(risks):
    t = make_tbl(UW)
    for r in title_rows(len(UW)): t.append(r)
    t.append(make_tr([hdr(UW[0], "ID", vmerge="restart"), hdr(UW[1], "Risk Statement", vmerge="restart"),
                      hdr(UW[2], "Owner", vmerge="restart"), hdr(UW[3], "Probability (A)", vmerge="restart"),
                      hdr(sum(UW[4:8]), "Impact", span=4), hdr(UW[8], "Score (C = A x MAX(B1:B4))", vmerge="restart"),
                      hdr(UW[9], "Response", vmerge="restart")], header=True))
    t.append(make_tr([hdr(UW[0], "", vmerge="cont"), hdr(UW[1], "", vmerge="cont"), hdr(UW[2], "", vmerge="cont"),
                      hdr(UW[3], "", vmerge="cont"), hdr(UW[4], "Scope (B1)"), hdr(UW[5], "Quality (B2)"),
                      hdr(UW[6], "Schedule (B3)"), hdr(UW[7], "Cost (B4)"), hdr(UW[8], "", vmerge="cont"),
                      hdr(UW[9], "", vmerge="cont")], header=True))
    for (rid, stmt, owner, a, b, resp, d, e, party, actions, status, comm) in risks:
        c = a * max(b)
        t.append(make_tr([make_tc(UW[0], rid), make_tc(UW[1], stmt), make_tc(UW[2], owner), num(UW[3], a)]
                         + [num(UW[4 + i], b[i]) for i in range(4)]
                         + [num(UW[8], c), make_tc(UW[9], resp)]))
    return t

def lower_table(risks):
    t = make_tbl(LW)
    t.append(make_tr([hdr(LW[0], "ID", vmerge="restart"), hdr(LW[1], "Revised Probability (D)", vmerge="restart"),
                      hdr(sum(LW[2:6]), "Impact", span=4), hdr(LW[6], "Revised Score (F = D x MAX(E1:E4))", vmerge="restart"),
                      hdr(LW[7], "Responsible Party", vmerge="restart"), hdr(LW[8], "Actions", vmerge="restart"),
                      hdr(LW[9], "Status", vmerge="restart"), hdr(LW[10], "Comments", vmerge="restart")], header=True))
    t.append(make_tr([hdr(LW[0], "", vmerge="cont"), hdr(LW[1], "", vmerge="cont"),
                      hdr(LW[2], "Scope (E1)"), hdr(LW[3], "Quality (E2)"), hdr(LW[4], "Schedule (E3)"), hdr(LW[5], "Cost (E4)"),
                      hdr(LW[6], "", vmerge="cont"), hdr(LW[7], "", vmerge="cont"), hdr(LW[8], "", vmerge="cont"),
                      hdr(LW[9], "", vmerge="cont"), hdr(LW[10], "", vmerge="cont")], header=True))
    for (rid, stmt, owner, a, b, resp, d, e, party, actions, status, comm) in risks:
        f = d * max(e)
        t.append(make_tr([make_tc(LW[0], rid), num(LW[1], d)] + [num(LW[2 + i], e[i]) for i in range(4)]
                         + [num(LW[6], f), make_tc(LW[7], party), make_tc(LW[8], actions),
                            make_tc(LW[9], status, jc="center", valign="center"), make_tc(LW[10], comm)]))
    return t

GW = [520, 1000, 2000, 2000, 2000]; GW.append(9964 - sum(GW))
def legend_table():
    """Kept on one page: every paragraph except those of the last row has keepNext."""
    t = make_tbl(GW)
    t.append(make_tr([make_tc(9964, "Probability and Impact Scale (1-5) used in this register", bold=True, span=6, keep=True)]))
    t.append(make_tr([hdr(w, h, keep=True) for w, h in zip(GW, LEGEND_HDR)]))
    for row in LEGEND:
        t.append(make_tr([make_tc(GW[0], row[0], jc="center", valign="center", narrow=True, keep=True)]
                         + [make_tc(w, v, keep=True) for w, v in zip(GW[1:], row[1:])]))
    t.append(make_tr([make_tc(9964, LEGEND_NOTE, span=6)]))
    return t

def gap_par(sz="8"):
    p = etree.Element(q("p"))
    ppr = sub(p, "pPr"); rp = sub(ppr, "rPr"); sub(rp, "sz", val=sz); sub(rp, "szCs", val=sz)
    return p

# ---------------------------------------------------------------- build
doc = Document(SRC)
body = doc.element.body
tbls = body.findall(q("tbl"))
def tbl_by_title(s):
    found = [t for t in tbls if cell_text(cells(rows(t)[0])[0]).strip() == s]
    assert len(found) == 1, (s, len(found))
    return found[0]
AL = tbl_by_title("ASSUMPTION LOG")
RBS = tbl_by_title("RESOURCE BREAKDOWN STRUCTURE")
assert not body.findall(".//" + q("tbl")) or all(cell_text(cells(rows(t)[0])[0]).strip() != "RISK REGISTER" for t in tbls), "register already present"
body_sectpr = body.find(q("sectPr"))
# page-break paragraph pattern used between the RTM tables
PB = [p for p in body.findall(q("p")) if p.find(q("pPr")) is not None and p.find(q("pPr")).find(q("pageBreakBefore")) is not None][0]

# drop the empty paragraphs after the RBS and end the RBS section there
nxt = RBS.getnext()
while nxt is not None and nxt.tag == q("p"):
    assert p_text(nxt) == "" and nxt.find(".//" + q("sectPr")) is None and len(nxt) == 0, "unexpected content after RBS"
    nn = nxt.getnext(); body.remove(nxt); nxt = nn
assert nxt is body_sectpr
new = [section_break_par(body_sectpr), heading_par("7. CREATE RISK REGISTER")]
k = 0
for gi, n in enumerate(GROUPS):
    grp = RISKS[k:k + n]; k += n
    if gi > 0:
        new.append(strip_ids(copy.deepcopy(PB)))
    new += [upper_table(grp), gap_par(), lower_table(grp)]
tail = gap_par("2")                      # 1-pt paragraph after the last table, so Word adds no empty page
sub(tail.find(q("pPr")), "spacing", line=20, lineRule="exact")
tail.find(q("pPr")).insert(0, tail.find(q("pPr"))[-1])   # spacing before rPr (schema order)
new += [gap_par(), legend_table(), tail]
anchor = RBS
for e in new:
    anchor.addnext(e); anchor = e

# ---------------------------------------------------------------- references to the register
R = Repl(body, "refs")
R("WBS and WBS Dictionary, Project Schedule, Assumption Log", "WBS and WBS Dictionary, Project Schedule, Assumption Log, Risk Register")
R("communications and procurement plans), monthly status and cost reports, final project report",
  "communications and procurement plans) with the risk register (R-01 to R-12), monthly status and cost reports, final project report")
R("communications and procurement plans) and obtain the Sponsor's approval;",
  "communications and procurement plans) with the risk register (R-01 to R-12) and obtain the Sponsor's approval;")
R("monitor risks and manage change requests;", "monitor the risks in the risk register and manage change requests;")
R("and a final review on Jan 13. Lessons learned",
  "and a final review on Jan 13; each review updates the probability, impact, score, actions and status of every risk in the register. Lessons learned")
R("the risk register is reviewed at least every two weeks.",
  "the risk register (R-01 to R-12) is reviewed at least every two weeks (1.7.5), and every risk with a score of 15 or more is reported in the monthly report (1.7.4).")
for n, e, o in R.log: print(f"  {n}/{e}  {o}")
doc.save(DST)
print("saved", DST, "groups", GROUPS)
