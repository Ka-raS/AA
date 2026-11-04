"""REFERENCE ONLY - do not run: this script built 4.docx from source files that no longer exist.
It shows how forms were made: clone the Charter title bar / header row / label+box tables,
clone a data table (Requirements Documentation was cloned from the Assumption Log), set column
widths (dlib.set_grid), add group headings (dlib.heading_par) and section breaks (SB)."""
"""Edit 4.docx: fixes, SRS -> Requirements Documentation + Project Scope Statement, sections, schedule placeholder."""
import sys, copy
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, __file__.rsplit("\\", 1)[0])
from docx import Document
from lxml import etree
from dlib import *

SRC, DST = r"E:\AA\4.docx", sys.argv[1]
MINH_MONTHS = [int(x) for x in sys.argv[2].split(",")]       # e.g. 152,152,40,64,104
TEAM_MONTHS = [int(x) for x in sys.argv[3].split(",")]

doc = Document(SRC)
body = doc.element.body
kids = list(body.iterchildren())
T = lambda i: el_text(kids[i])
assert T(8) == "Software Requirements Specification", T(8)
assert T(41) == "PROJECT CHARTER" and T(116).startswith("ASSUMPTION LOG") and T(119).startswith("REQUIREMENTS TRACEABILITY")
assert T(128).startswith("WORK BREAKDOWN") and T(131).startswith("WBS DICTIONARY") and T(152).startswith("WBS DICTIONARY: RESOURCE")
body_sectpr = kids[-1]
assert body_sectpr.tag == q("sectPr")

R = Repl(body, "4.docx")
# ---------------- cover ----------------
replace_in_p(kids[8], "Software Requirements Specification", "Project Management Documents")

# ---------------- Project Charter ----------------
R("Fully deliver the SRS functions (Function List) on iOS, Android and Web:", "Deliver the functions of the Requirements Documentation (FR-01 to FR-14) on iOS, Android and Web:")
R("2,464 h of the team's 4,400 available hours", "2,464 h of the team's 4,320 available hours (5 members x 108 working days x 8 h)")
R("Also includes a Flutter Web-based admin dashboard with full CRUD operations,", "Also includes a Flutter Web-based admin dashboard with CRUD operations for songs, albums, artists and user accounts,")
R("and the full lifecycle from requirement gathering", "and the lifecycle from requirement gathering")
R("music content licensing; customer support and maintenance beyond the 1-year post-launch warranty period.",
  "music content licensing; User Acceptance Testing (performed by the Customer on staging); customer support and maintenance beyond the 1-year post-launch warranty period.")
R("a Flutter Web-based admin dashboard with full CRUD operations for managing content and users.",
  "a Flutter Web-based admin dashboard with CRUD operations for songs, albums, artists and user accounts.")
R("Ensure product quality and user experience by strictly adhering to Apple Human Interface Guidelines and Google Material Design 3, resolving 100% of critical bugs, and maintaining a crash-free rate of >= 99% upon release.",
  "Follow Apple Human Interface Guidelines and Google Material Design 3 on every screen, resolve 100% of critical bugs before the release candidate and keep a crash-free rate of >= 99% after release.")
R("Achieve a minimum User Acceptance Testing (UAT) score of 4.2/5.",
  "User Acceptance Testing (UAT) score of at least 4.2/5; no open critical defect at the release candidate (1.5.8); crash-free rate >= 99% during hypercare (1.6.5).")
R("Total pre-approved budget: 757,000,000 VND (items total 756,968,012 VND, rounded up), allocated as follows:",
  "Total pre-approved budget: 757,000,000 VND = sum of the items below (406,794,048 + 55,500,000 + 48,900,000 + 137,904,988 + 71,822,880 + 36,046,096 = 756,968,012 VND), rounded up to the next million VND, allocated as follows:")
R("- The app has been published on the App Store and Google Play and is running stably.",
  "- The apps have been published on the App Store and Google Play and keep a crash-free rate of at least 99% during hypercare (1.6.5).")
R("- The Sponsor/instructor has signed off on project acceptance.", "- The Sponsor has signed off the project closure (1.7.9).")
R("- Requirement changes cause the project to significantly exceed scope, cost, schedule.",
  "- Approved requirement changes need more than the contingency and management reserves or move the handover more than 14 days past the baseline.")
R("approve minor scope and schedule changes (under 2% of budget, i.e. 15,140,000 VND, or 3 days of schedule)",
  "approve scope and schedule changes below 2% of budget (15,140,000 VND) or 3 days of schedule")
R("has the authority to approve or reject major changes to project scope", "has the authority to approve or reject changes to project scope")

# ---------------- RTM ----------------
rtm1, rtm2 = kids[119], kids[122]
def rtm_cell(rid, col, text):
    for t in (rtm1, rtm2):
        try:
            tr = find_row(t, rid)
        except KeyError:
            continue
        set_cell_text(cells(tr)[col], text); return
    raise KeyError(rid)
rtm_cell("FR-05", 6, "WBS 1.3.2, 1.4.4")
rtm_cell("FR-10", 6, "WBS 1.3.2, 1.4.4")
rtm_cell("FR-08", 6, "WBS 1.4.1, 1.4.3")
rtm_cell("NFR-12", 6, "WBS 1.2.4, 1.3.4, 1.4.3, 1.4.4")
R("Content Administration: full CRUD for songs", "Content Administration: CRUD for songs")
R("Profile changes persist and display correctly after re-login", "Profile changes are saved and shown after re-login")
R("Relevant results returned; API p90 latency < 200 ms", "Every result matches the keyword in the title, artist, album or genre; API p90 latency < 200 ms")
R("Playback controls function correctly at all supported quality levels", "Play, pause, seek, skip and loop work at every supported quality level")
R("Playlist CRUD operations persist correctly", "Playlist create, edit, delete and share operations are saved and shown after re-login")
R("Downloaded songs play back correctly with network disabled", "Downloaded songs play from start to end with the network disabled")
R("Favorite, rate and comment actions persist; history displays correctly", "Favorite, rate and comment actions are saved; the history lists the songs played, newest first")
R("All CRUD operations succeed and reflect correctly for end-users", "Every CRUD operation succeeds and the change is shown to end-users")
R("Load test sustains 15,000 concurrent connections without degradation", "Load test with 15,000 virtual users: error rate < 1% and API p90 < 200 ms")
R("Every screen displays correctly in both languages; the selected language persists after restart",
  "Every screen shows its text in Vietnamese and English; the selected language persists after restart")
R("Measured cold/warm start times meet targets on reference devices", "Median of 10 launches on each of the 4 reference devices meets the targets")
R("Passwords hashed with bcrypt (cost >= 10); JWT access token expiry <= 1 h; TLS 1.2+ on all traffic",
  "Passwords hashed with bcrypt (cost >= 10); JWT access token expiry <= 1 h; TLS 1.2+ on all traffic; no open critical or high finding")
# RTM values reused by the Requirements Documentation, so both forms say the same thing
RTM_INFO = {}
for t in (rtm1, rtm2):
    for tr in rows(t):
        cs = [cell_text(c).strip() for c in cells(tr)]
        if cs and (cs[0].startswith("FR-") or cs[0].startswith("NFR-")):
            RTM_INFO[cs[0]] = {"src": cs[2], "pri": cs[3], "cat": cs[4], "ver": cs[7], "val": cs[8]}
assert len(RTM_INFO) == 26

# ---------------- WBS / WBS Dictionary ----------------
R("Build login, profile, search, playlist, library and social screens", "Build login, profile, search, recommendation, playlist, library and social screens", expect=2)
R("login, profile, search, playlist, library and social screens with Vietnamese/English localization",
  "login, profile, search, recommendation, playlist, library and social screens with Vietnamese/English localization")
R("User Acceptance Testing is performed by the Customer and is outside the WBS.",
  "User Acceptance Testing is performed by the Customer and is outside the WBS. 1.7.4 and 1.7.5 are recurring work packages, scheduled as 1-day and 2-day sessions (see the Project Schedule).")
R("Labor rates in VND/hour (2 years of experience, cost.docx):", "Labor rates in VND/hour (2 years of experience, Basis of Estimates):", expect=7)
R("Saturday Sep 5, 2026 is worked in place of the second National Day holiday (Team Charter), so",
  "National Day falls on Sep 1-2, 2026 and Saturday Sep 5 is worked in place of Sep 1 (Team Charter), so")
R("SRS Section 2 (Project Scope), Section 3 (Function List) and Section 4 (Non-functional Requirements); Project Charter;",
  "SRS; Requirements Documentation; Project Scope Statement; Project Charter;")
R("SRS Section 3 (Function List); Flutter and Flutter Web;", "Requirements Documentation (FR-01 to FR-14); Flutter and Flutter Web;")
R("Apps published and running stably on the App Store and Google Play;", "Apps published on the App Store and Google Play with a crash-free rate >= 99% during hypercare (1.6.5);")
R("Sponsor/instructor sign-off per the Charter's Project Exit Criteria.", "Sponsor sign-off of the project closure (1.7.9) per the Charter's Project Exit Criteria.")
R("1.7.4, 1.7.5 and 1.7.6 are recurring: their hours are shown as one block in the Project Schedule, but the work is spread over the project (a status and cost report at each month end, a risk review every two weeks, lessons learned at each phase end).",
  "1.7.4 and 1.7.5 are recurring and are scheduled as sessions on days without other PM work: 1.7.4 as five month-end status and cost reports (Sep 11; Oct 15-16; Nov 13 and 16; Dec 10-11; Jan 14-15), 1 day in Month 1 (before the cost baseline) and 2 days in Months 2-5; 1.7.5 as eight 1-day risk and change-request reviews, every two weeks from Oct 14 to Jan 5, and a final review on Jan 13. Lessons learned are collected in a 30-minute review at each phase end and compiled in 1.7.6.")
R("as in the Project Charter and cost.docx.", "as in the Project Charter and the Cost Estimating Worksheet.")
R("Workloads are balanced: every member has 448-544 hours of work spread over all five months (see the Project Schedule in Microsoft Project).",
  "Workload: every member has 448-544 hours of work, spread over all five months (see the Project Schedule).")
R("so the work packages of one person never overlap in the Project Schedule.",
  "so the work packages (and the sessions of 1.7.4 and 1.7.5) of one person never overlap in the Project Schedule.")

# formula headers in the WBS Dictionary cost tables ("use excel": show how each column is computed)
wbsd = [kids[i] for i in range(131, 152, 3)]
assert all(el_text(t).startswith("WBS DICTIONARY") for t in wbsd) and len(wbsd) == 7
LAB = {3: "Hours (A)", 4: "Rate (B)", 5: "Total (C = A x B)", 6: "Units (D)", 7: "Cost (E)", 8: "Total (F = D x E)"}
for t in wbsd:
    rs = rows(t)
    h5 = [r for r in rs if [cell_text(c) for c in cells(r)][:3] == ["ID", "Activity", "Resource"]][0]
    h6 = h5.getnext()
    c6 = cells(h6)
    assert [cell_text(c) for c in c6][3:9] == ["Hours", "Rate", "Total", "Units", "Cost", "Total"], [cell_text(c) for c in c6]
    for k, v in LAB.items():
        set_cell_text(c6[k], v)
    c5 = cells(h5)
    assert cell_text(c5[-1]) == "Total Cost"
    set_cell_text(c5[-1], "Total Cost (G = C + F)")
    for r in rs:
        cs = cells(r)
        if len(cs) > 2 and cell_text(cs[1]) == "Total" and cell_text(cs[0]) == "":
            set_cell_text(cs[1], "Total (SUM)")

# Resource summary: Minh's months after the 1.7.4/1.7.5 sessions, team totals, formula headers
rsum = kids[152]
hdr = find_row(rsum, "Team Member")
hc = cells(hdr)
assert cell_text(hc[7]) == "Total Hours" and cell_text(hc[9]) == "Labor Cost (VND)"
set_cell_text(hc[7], "Total Hours (SUM)")
set_cell_text(hc[9], "Labor Cost (VND) (= Total Hours x rate)")
def fmt(n): return f"{n:,}"
mr = cells(find_row(rsum, "Trần Vũ Tiến Minh"))
for i, v in enumerate(MINH_MONTHS):
    set_cell_text(mr[2 + i], str(v))
assert sum(MINH_MONTHS) == 512 and cell_text(mr[7]) == "512"
tr_ = cells(find_row(rsum, "Team Total"))
for i, v in enumerate(TEAM_MONTHS):
    set_cell_text(tr_[2 + i], fmt(v))
assert sum(TEAM_MONTHS) == 2464

# ---------------- new forms built from the SRS content ----------------
al = kids[116]
def clone_form_header(title, date):
    """Title bar + 'Project Title | Date Prepared' row, cloned from the Charter."""
    t_title = copy.deepcopy(kids[41]); set_cell_text(cells(rows(t_title)[0])[0], title)
    gap = copy.deepcopy(kids[42])
    hdr_t = copy.deepcopy(kids[44]); rr = rows(hdr_t)
    hdr_t.remove(rr[1])
    c0, c1 = cells(rr[0])
    set_cell_lines(c0, [[("Project Title: ", True), ("Development and Deployment of a Music Streaming Platform", False)]])
    set_cell_lines(c1, [[("Date Prepared: ", True), (date, False)]])
    return [t_title, gap, hdr_t]
def label(text):
    p = copy.deepcopy(kids[45])
    ts = texts(p); ts[0].text = text
    for t in ts[1:]: t.text = ""
    return p
def box(lines):
    b = copy.deepcopy(kids[46])
    set_cell_lines(cells(rows(b)[0])[0], lines)
    return b
def gap():
    return copy.deepcopy(kids[42])

# Requirements Documentation (Book of Forms 2.6): data from the SRS Function List / NFR list and the RTM
REQ = [
 ("FR-01", "Registration, Login: create an account or log in with email/phone number and password.", "Customer", "Functional", "Must-have", "Account is created and login succeeds for valid credentials; invalid credentials are rejected.", "UAT", "1.4 Full Features (API in 1.3.1)"),
 ("FR-02", "Profile Management: update the profile picture, bio and password.", "Customer", "Functional", "Should-have", "Profile changes are saved and shown after re-login.", "Functional test + UAT", "1.4 Full Features (API in 1.3.1)"),
 ("FR-03", "Search Songs, Artists, Albums: search by song title, artist, album or genre, with keyword suggestions while typing.", "Customer", "Functional", "Must-have", "Every result matches the keyword in the title, artist, album or genre; API p90 latency < 200 ms.", "Functional test + performance test", "1.4 Full Features"),
 ("FR-04", "Music Streaming (Playback): stream songs with selectable audio quality; play, pause, seek, skip and loop on mobile and web.", "Customer", "Functional", "Must-have", "Play, pause, seek, skip and loop work at every supported quality level.", "Functional test + UAT", "1.3 MVP"),
 ("FR-05", "Music Video Playback: stream the music videos of songs at 360p, 720p or 1080p, with fullscreen landscape playback.", "Customer", "Functional", "Could-have", "Video plays at each listed resolution and rotates to fullscreen landscape.", "Functional test", "1.4 Full Features"),
 ("FR-06", "Playlist Management: create, edit, delete and share personal playlists; add or remove songs.", "Customer", "Functional", "Should-have", "Playlist create, edit, delete and share operations are saved and shown after re-login.", "Functional test", "1.4 Full Features"),
 ("FR-07", "Library Management: add artists, albums, songs and other users' playlists to the personal library.", "Customer", "Functional", "Should-have", "Added items appear in and persist in the user's library.", "Functional test", "1.4 Full Features"),
 ("FR-08", "Recommendations: suggest music from the user's listening history, song data (genre, artist) and the preferences of similar listeners; effectiveness is measured by likes, skips and plays.", "Customer", "Functional", "Should-have", "Recommendations reflect the stated inputs; effectiveness tracked via likes, skips and plays.", "UAT + analytics review", "1.4 Full Features"),
 ("FR-09", "Live Premiere: join the scheduled premiere of a new release, listen at the same time as other users and chat live.", "Customer", "Functional", "Must-have", "Users join the premiere, exchange chat and stay in sync in real time under concurrent load.", "Load test (WebSocket concurrency) + UAT", "1.3 MVP"),
 ("FR-10", "Offline Download: download songs to the device and play them without an internet connection.", "Customer", "Functional", "Could-have", "Downloaded songs play from start to end with the network disabled.", "Functional test", "1.4 Full Features"),
 ("FR-11", "Favorites, Ratings, Commenting: mark songs as favorites, rate and comment on songs and view the recent listening history.", "Customer", "Functional", "Should-have", "Favorite, rate and comment actions are saved; the history lists the songs played, newest first.", "Functional test", "1.4 Full Features"),
 ("FR-12", "Content Administration: admins upload, edit or remove songs, albums and artists and manage user accounts in the web dashboard.", "Customer", "Functional", "Must-have", "Every CRUD operation succeeds and the change is shown to end-users.", "Functional + integration test", "1.4 Full Features"),
 ("FR-13", "System Notifications: push notifications about new releases, new recommendations and replies to the user's comments.", "Customer", "Functional", "Could-have", "Notifications are delivered for each stated trigger event.", "Functional test", "1.4 Full Features"),
 ("FR-14", "In-App Purchase: users buy the paid features defined by the Sponsor through App Store / Google Play billing; the backend verifies the store receipt and unlocks the feature; the system processes no payments and stores no card data.", "Sponsor", "Functional", "Should-have", "Sandbox purchase unlocks the feature only after receipt verification; no card data stored.", "Functional test (store sandbox)", "1.4 Full Features"),
 ("NFR-01", "The server supports at least 9,000 concurrent users at launch and up to 15,000 without performance degradation.", "Sponsor", "Non-functional (Scalability)", "Must-have", "Load test with 15,000 virtual users: error rate < 1% and API p90 < 200 ms.", "Load testing", "1.5 Testing and QA (1.5.3)"),
 ("NFR-02", "Backend API response latency < 200 ms for 90% of queries.", "Sponsor", "Non-functional (Performance)", "Must-have", "Measured p90 response time < 200 ms under load.", "Load and performance testing", "1.5 Testing and QA (1.5.3)"),
 ("NFR-03", "App launch speed: cold start < 2 s, warm start < 1 s.", "Sponsor", "Non-functional (Performance)", "Should-have", "Median of 10 launches on each reference device meets the targets.", "Performance testing", "1.5 Testing and QA (1.5.3)"),
 ("NFR-04", "Battery consumption < 6% per hour of audio streaming over Wi-Fi, screen off, volume 50%, on a device with about 4,000 mAh battery capacity.", "Sponsor", "Non-functional (Performance)", "Should-have", "Measured battery drain meets the target under the stated test conditions.", "Device testing", "1.5 Testing and QA (1.5.3)"),
 ("NFR-05", "Passwords hashed with bcrypt (cost factor >= 10); JWT access/refresh tokens, access token expiry <= 1 hour; TLS 1.2 or higher; refresh tokens in platform secure storage (Keychain / Keystore / httpOnly cookie on Web).", "Project Manager", "Non-functional (Security)", "Must-have", "bcrypt cost >= 10; JWT access token expiry <= 1 h; TLS 1.2+ on all traffic; no open critical or high finding.", "Security testing / code review", "1.5 Testing and QA (1.5.3)"),
 ("NFR-06", "UI/UX follows Apple Human Interface Guidelines and Google Material Design 3; UAT satisfaction score >= 4.2/5.", "Sponsor", "Non-functional (Usability)", "Must-have", "Design review confirms guideline conformance; UAT score >= 4.2/5.", "UAT", "1.4 Full Features; UAT score before 1.6.3"),
 ("NFR-07", "Containerized architecture, horizontally scalable to 50,000 concurrent users in future phases.", "Sponsor", "Non-functional (Scalability)", "Should-have", "Architecture review confirms the horizontal scaling path; a full-scale load test is outside the current scope.", "Architecture review", "1.6 Deployment (1.6.4)"),
 ("NFR-08", "Backend REST API documented with the OpenAPI 3.0 specification; all endpoints versioned under /api/v1.", "Project Manager", "Non-functional (Maintainability)", "Must-have", "OpenAPI 3.0 specification published; all endpoints versioned.", "Design review", "1.2 Design (1.2.6)"),
 ("NFR-09", "Automated backend test coverage of at least 70%.", "Project Manager", "Non-functional (Quality)", "Should-have", "Coverage report shows >= 70% backend coverage.", "Automated coverage report", "1.4 Full Features"),
 ("NFR-10", "PostgreSQL backed up daily; Recovery Point Objective (RPO) 24 hours, Recovery Time Objective (RTO) 4 hours.", "Project Manager", "Non-functional (Reliability)", "Must-have", "Daily backup job configured; restore drill completes within the RTO.", "Backup and restore drill", "1.6 Deployment (1.6.4)"),
 ("NFR-11", "Centralized logging and uptime monitoring for all backend services, with alerts on error-rate and latency threshold breaches.", "Project Manager", "Non-functional (Operability)", "Should-have", "Logs and alerts configured and verified before go-live.", "Operational readiness review", "1.6 Deployment (1.6.4)"),
 ("NFR-12", "GUI available in Vietnamese and English on the mobile apps, the web client and the admin dashboard.", "Customer", "Non-functional (Usability)", "Must-have", "Every screen shows its text in Vietnamese and English; the selected language persists after restart.", "UAT", "1.4 Full Features"),
]
rd = copy.deepcopy(al)
rr = rows(rd)
set_cell_text(cells(rr[0])[0], "REQUIREMENTS DOCUMENTATION")
assert replace_all(rr[1], "September 29, 2026", "September 9, 2026 (revised September 17, 2026)") == 1
HEAD = ["ID", "Requirement", "Stakeholder", "Category", "Priority", "Acceptance Criteria", "Test or Verification", "Phase or Release"]
for c, h in zip(cells(rr[2]), HEAD):
    set_cell_text(c, h)
tmpl = rr[3]
for r_ in rr[4:]:
    rd.remove(r_)
anchor = tmpl
for rec in REQ:
    info = RTM_INFO[rec[0]]
    rec = list(rec)
    rec[2], rec[3], rec[4] = info["src"], info["cat"], info["pri"]
    rec[5] = info["ver"] if info["ver"].endswith(".") else info["ver"] + "."
    rec[6] = info["val"]
    nr = copy.deepcopy(tmpl)
    for c, v in zip(cells(nr), rec):
        set_cell_text(c, v)
    anchor.addnext(nr); anchor = nr
rd.remove(tmpl)
set_grid(rd, [560, 2200, 1100, 950, 700, 2050, 1200, 1204])
reqdoc = [rd]

# Project Scope Statement (Book of Forms 2.8): SRS scope + Charter + Assumption Log
scope_p1 = clone_form_header("PROJECT SCOPE STATEMENT", "September 7, 2026 (revised October 8, 2026: budget)") + [
 label("Project Scope Description:"),
 box(["Build and deploy a music streaming application for iOS, Android and Web that lets users discover, listen to, rate and comment on music and manage their favorite music on their devices. The system uses a client-server architecture with a Node.js (Express.js) backend API and a PostgreSQL database. A team of 5 members delivers it in 5 months (August 17, 2026 - January 15, 2027) with a budget of 757,000,000 VND.",
      [("Product scope:", True)],
      "- Clients: iOS and Android apps and a Flutter Web client for end-users, built on a single Flutter codebase to reduce development cost and time and keep one UI; GUI in Vietnamese and English.",
      "- Backend: Node.js (Express.js) REST API and PostgreSQL for API management, user authentication and data queries; WebSocket real-time messaging for Live Premiere; audio and music-video files kept in object storage and served through a Content Delivery Network (CDN); PostgreSQL stores only application data and media metadata.",
      "- Web: the Flutter Web client offers the same core functions as the mobile apps (registration/login, playback, search, playlist and library management, recommendations, Live Premiere, favorites/ratings/comments, notifications; offline download and video playback where the browser supports them); a Flutter Web admin dashboard offers CRUD operations for songs, albums, artists and user accounts.",
      "- User functions (Requirements Documentation FR-01 to FR-14): registration/login and profile management; music and music-video playback; Live Premiere; rating, comments and favorites; playlist and library management, offline download; search and recommendations; system notifications; in-app purchases through App Store / Google Play billing, where the backend only verifies the store receipt and unlocks the paid feature.",
      "- Supported software: iOS 15.0 and above, Android 8.0 and above, and Chrome 96+, Edge 96+, Firefox 99+ and Safari 15.6+ (the oldest versions Flutter Web runs on) with JavaScript enabled.",
      "- Supported hardware: mobile devices with at least 2 GB RAM, screens of 4.7 to 6.7 inches, resolutions 720p, 1080p and 1440p and aspect ratios 16:9, 18:9 and 20:9; web devices (desktop, laptop or tablet) with at least 2 GB RAM and screens from 1024x768 (768x1024 portrait tablet) up to 1920x1080.",
      "- Server infrastructure: cloud-hosted, containerized backend sized for at least 9,000 concurrent users at launch (up to 15,000 without performance degradation) and horizontally scalable to 50,000 concurrent users in future phases.",
      [("Project scope:", True)],
      "- Requirement gathering, system analysis, UI/UX design, frontend and backend development, QA testing (integration, load, security and device), deployment to production, store publishing, administrator training and handover documentation.",
      "- 1-year warranty with technical support and system maintenance after the production deployment."]),
 label("Project Deliverables:"),
 box(["- 1.1 Requirements and Planning: Project Charter, SRS, Requirements Traceability Matrix with acceptance criteria, Project Scope Statement, WBS and WBS Dictionary.",
      "- 1.2 System and Database Design: system architecture, ER diagram and PostgreSQL schema, detailed design (modules, class and sequence diagrams), general API specification (OpenAPI 3.0), wireframes, high-fidelity mockups and clickable prototype; repository, CI/CD pipeline, staging, backups and CDN.",
      "- 1.3 MVP Development and 1.4 Full Feature Development: source code with unit tests for the backend API, the Live Premiere WebSocket service, the Flutter apps for iOS, Android and Web and the admin dashboard.",
      "- 1.5 Testing and QA: test plan and quality metrics, test cases, integration, load, security and device test reports and the test summary report.",
      "- 1.6 Deployment and Closeout: release candidate on staging, production release with monitoring and backups, apps on the App Store and Google Play, user guide, admin guide and release notes (Vietnamese/English), deployment and operations documentation, administrator training.",
      "- 1.7 Project Management: Project Management Plan (schedule, cost baseline, risk, quality, communications and procurement plans), monthly status and cost reports, final project report."]),
 label("Product Acceptance Criteria:"),
 box(["- Every Must-have and Should-have requirement (FR-01 to FR-14, NFR-01 to NFR-12) meets its acceptance criteria in the Requirements Documentation and passes its RTM verification; Could-have functions (FR-05, FR-10, FR-13) are delivered as capacity allows.",
      "- Load test (1.5.3) with 15,000 virtual users: error rate < 1% and API p90 latency < 200 ms; cold start < 2 s and warm start < 1 s; battery consumption < 6% per hour.",
      "- No open critical or high defect at the release candidate (1.5.8); automated backend test coverage >= 70%; restore drill completed within the 4-hour RTO.",
      "- UAT score of at least 4.2/5 and the Customer's acceptance sign-off (1.6.3); crash-free rate >= 99% during hypercare (1.6.5).",
      "- Apps published on the App Store and Google Play; handover deliverables accepted by the Customer (1.6.12); project closure signed by the Sponsor (1.7.9)."]),
 label("Project Exclusions:"),
 box(["- Separate native codebases for iOS and Android; native desktop applications (Windows, macOS, Linux).",
      "- Processing payments directly (own payment gateway, card details, e-wallets such as MoMo, ZaloPay, VNPay).",
      "- Designing music subscription plans and prices (the paid products are defined by the Sponsor).",
      "- GUI languages other than Vietnamese and English.",
      "- Product marketing and user acquisition; music content licensing.",
      "- User Acceptance Testing: planned and performed by the Customer on the staging environment; the team only deploys the release candidate and fixes the defects the Customer reports (1.6.1, 1.6.2).",
      "- Customer support and maintenance after the 1-year warranty period."]),
]
pb_tmpl = kids[58]
scope_p2 = [copy.deepcopy(pb_tmpl), copy.deepcopy(kids[41]), gap()]
set_cell_text(cells(rows(scope_p2[1])[0])[0], "PROJECT SCOPE STATEMENT")
scope_p2 += [
 label("Project Constraints:"),
 box(["- Budget: 757,000,000 VND = cost baseline 720,921,916 VND + management reserve 36,046,096 VND, released only by the Sponsor (Project Charter).",
      "- Schedule: August 17, 2026 - January 15, 2027 (108 working days); the Summary Milestones of the Project Charter; schedule variance of at most 14 days.",
      "- Team: 5 fixed members; each member works on one work package at a time and at most 176 hours in any month.",
      "- Technology: Flutter (single codebase for iOS, Android and Web), Node.js (Express.js), PostgreSQL and WebSocket (Project Charter, Technical Decisions); minimum iOS 15.0, Android 8.0 and the browser versions in the scope description.",
      "- Payments only through App Store / Google Play billing; no card data stored by the system.",
      "- App Store Review Guidelines, Google Play Developer Program Policies and the Vietnamese Law on Personal Data Protection.",
      "- Change control: the PM approves changes below 2% of the budget (15,140,000 VND) or 3 days of schedule; larger changes go to the Sponsor."]),
 label("Project Assumptions:"),
 box(["- The Sponsor and the Customer attend the meetings, reviews, demos and sign-offs on the scheduled dates (Assumption Log A-01 to A-12).",
      "- The App Store and Google Play review the apps within 3 working days (A-13); the cloud and CDN accounts are active within 2 working days of the purchase (A-14).",
      "- The requirements collected in 1.1.1 are complete: 14 functions and 12 non-functional requirements.",
      "- The Customer performs UAT on staging from December 23, 2026 and signs the acceptance on January 4, 2027 (A-09).",
      "- The Sponsor defines the paid features sold through in-app purchase.",
      "- The music catalog loaded during the project fits in 1,000 GB of object storage (Resource Requirements PR-03)."]),
]

# Project Schedule (Book of Forms 2.18): placeholder for the Gantt chart exported from 4.xml
sched = clone_form_header("PROJECT SCHEDULE", "September 29, 2026") + [
 label("Gantt Chart (Microsoft Project, 4.xml):"),
 box(["[Insert here the Gantt chart exported from 4.xml.]",
      "Columns: WBS | Task Name | Duration | Start | Finish | Predecessors | Resource Names, with the Gantt bars from August 17, 2026 to January 15, 2027 and the resource name next to each bar. Milestones (0 days) are shown as diamonds; 1.7.4 and 1.7.5 are summary rows of their recurring sessions."]),
]

# ---------------- reassemble the body ----------------
def SB():
    p = section_break_par(body_sectpr)
    # keep the section-break paragraph tiny so it never spills onto an empty page
    ppr = p.find(q("pPr"))
    sp = etree.Element(q("spacing")); sp.set(q("before"), "0"); sp.set(q("after"), "0"); sp.set(q("line"), "20"); sp.set(q("lineRule"), "exact")
    ppr.insert(0, sp)
    rpr = etree.Element(q("rPr")); sz = etree.SubElement(rpr, q("sz")); sz.set(q("val"), "2")
    ppr.insert(1, rpr)
    return p
H = heading_par
K = lambda a, b: kids[a:b + 1]
new = (K(0, 9) + [SB(), H("1. CREATE PROJECT CHARTER")]
       + K(41, 56) + [kids[58]] + K(59, 71) + [kids[73]] + K(74, 90) + [kids[92]] + K(93, 113) + [SB()]
       + [kids[116], SB()]
       + reqdoc + [SB()]
       + scope_p1 + scope_p2 + [SB()]
       + [H("2. CREATE WBS")] + [kids[119], kids[121], kids[122], kids[124], kids[125], SB()]
       + [kids[128], SB()])
wd = []
for i in range(131, 153):
    if i in (132, 135, 138, 141, 144, 147, 150, 153):
        continue  # "Page x of 8" labels
    wd.append(kids[i])
new += wd + [SB(), H("3. CREATE PROJECT SCHEDULE")] + sched
for k in list(body):
    body.remove(k)
for k in new:
    body.append(k)
body.append(body_sectpr)
# no "Page x of y" text may remain in the body (footers carry the page numbers)
import re as _re
left = [el_text(p) for p in body.iter(q("p")) if _re.fullmatch(r"Pages? [\d-]+ of \d+", el_text(p).strip())]
assert not left, left
# Arial (font of the Book of Forms templates used by cost/quality/resource) is wider than the
# Times New Roman fallback of the old 4.docx: body text goes 1 pt smaller so every form keeps its pages.
SHRINK = int(sys.argv[4]) if len(sys.argv) > 4 else 2
for tag in ("sz", "szCs"):
    for e in body.iter(q(tag)):
        v = int(e.get(q("val")))
        if 14 <= v <= 22:
            e.set(q("val"), str(v - SHRINK))
doc.save(DST)
print("saved", DST)
for n, e, o in R.log:
    if n != e: print("COUNT", n, e, o)
