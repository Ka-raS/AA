"""Apply the November notes, keeping the approved numeric/schedule baseline."""
import sys, re, copy, json, shutil
from pathlib import Path
from lxml import etree
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from dlib import *

ROOT = Path(__file__).resolve().parents[1]
TMP = ROOT / 'tmp/revision_20261118'
TMP.mkdir(parents=True, exist_ok=True)
SRC = TMP / '4.original.docx'
if not SRC.exists(): shutil.copy2(ROOT/'4.docx', SRC)
DATA = json.loads((TMP/'source.json').read_text(encoding='utf-8'))
MINH, DAI, HIEU, A, THU = 'Trần Vũ Tiến Minh', 'Nguyễn Phú Đại', 'Trần Đình Hiếu', 'Trần Văn A', 'Lê Quang Thu'
PROJECT = 'Development and Deployment of a Music Streaming Platform'

STAKES = [
['Nguyễn Trung B / PTIT Innovation Fund','Project Sponsor','Via PM; contact details to be recorded','Approve funding, PMP and changes beyond PM authority.','Budget <= 757,000,000 VND; closure sign-off Jan 15, 2027.','External; high power / high interest'],
['Đỗ Thị Bích Ngọc','Customer; end-user representative','Via PM; contact details to be recorded','Approve requirements/design; perform UAT; accept handover.','Must/Should requirements pass; UAT >= 4.2/5; acceptance Jan 4.','External; high power / high interest'],
[MINH,'Project Manager','Jira, team chat; email to be recorded','Plan scope, schedule, cost, resources, risks and stakeholder communication.','One owner per WP; changes within delegated authority; timely decisions.','Internal; high power / high interest'],
[DAI,'Business Analyst / Full-Stack Developer','Jira, GitHub, team chat','Maintain requirements/RTM; implement assigned APIs and Flutter features.','Approved API contract, measurable acceptance and controlled changes.','Internal; medium power / high interest'],
[HIEU,'Backend Developer / Technical Lead','Jira, GitHub, team chat','Design architecture/API; develop backend; review code and fix defects.','API p90 < 200 ms; backend coverage >= 70%; documented interfaces.','Internal; medium power / high interest'],
[A,'QA Engineer / DevOps Engineer','Jira, GitHub, team chat','Trace tests to requirements; test integration/load/security/devices; operate CI/CD.','15,000-VU test passes; no open S1/S2 at RC; reproducible releases.','Internal; medium power / high interest'],
[THU,'UI/UX Designer / Flutter Developer','Design workspace, Jira, team chat','Design approved screens; build Flutter clients; submit store releases.','One Flutter codebase; Vietnamese/English UI; Apple HIG / Material 3.','Internal; medium power / high interest'],
['End-users','Listeners on iOS, Android and Web','Customer representation; UAT feedback','Discover, play and organize music; use Live Premiere; protect personal data.','Playback/chat meet FR-09; account/data deletion works; UI >= 4.2/5.','External; low power / high interest'],
['Customer administrators','Content and account administrators','Via Customer and PM; names to be confirmed','Manage songs, artists, albums and users through the web dashboard.','FR-12 tests pass; training Jan 12-13; accepted admin/operations guides.','External; medium power / high interest'],
['Bizfly Cloud','Cloud, managed database and CDN provider','Provider support portal via PM / DevOps','Activate purchased services within 2 working days (A-14); supply PR-01/02/04.','Agreed service capacity and billing; 5,000 GB/month CDN allowance.','External; medium power / medium interest'],
['VNG Cloud','Object storage provider','Provider support portal via PM / DevOps','Supply 1,000 GB/month object storage for 3 months (PR-03).','Accounts active before setup; documented access and charges.','External; medium power / medium interest'],
['Apple App Store','iOS distribution and billing platform','App Store Connect via Flutter developer','Compliant build, privacy/account deletion, store billing and developer account.','Complete review submission; working TestFlight build; policy compliance.','External; high power / low interest'],
['Google Play','Android distribution and billing platform','Play Console via Flutter developer','Compliant build, Data safety/account deletion, store billing and verified account.','Complete review submission; working internal-test build; policy compliance.','External; high power / low interest'],
]

def font_run(r, size=7.5, bold=False, color='000000'):
    r.font.name='Arial';r.font.size=Pt(size);r.font.bold=bold;r.font.color.rgb=RGBColor.from_string(color)
    rp=r._element.get_or_add_rPr();rf=rp.find(q('rFonts'))
    if rf is None:rf=etree.SubElement(rp,q('rFonts'))
    for k in ['ascii','hAnsi','cs','eastAsia']:rf.set(q(k),'Arial')

def paragraph(doc,text='',size=8,bold=False,keep=False):
    p=doc.add_paragraph();p.paragraph_format.space_after=Pt(4);p.paragraph_format.space_before=Pt(0)
    p.paragraph_format.line_spacing=1.05;p.paragraph_format.keep_with_next=keep
    font_run(p.add_run(text),size,bold);return p

def table(doc, headers, data, widths, size=7.5):
    t=doc.add_table(rows=1,cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
    grid=t._tbl.find(q('tblGrid'))
    for g,w in zip(grid,widths):g.set(q('w'),str(w))
    pr=t._tbl.tblPr
    w=pr.find(q('tblW'));w.set(q('w'),str(sum(widths)));w.set(q('type'),'dxa')
    borders=etree.SubElement(pr,q('tblBorders'))
    for side in ['top','left','bottom','right','insideH','insideV']:
        e=etree.SubElement(borders,q(side));e.set(q('val'),'single');e.set(q('sz'),'4');e.set(q('color'),'D9D9D9')
    for j,h in enumerate(headers):t.rows[0].cells[j].text=h
    for vals in data:
        row=t.add_row()
        for c,v in zip(row.cells,vals):c.text=str(v)
    for i,row in enumerate(t.rows):
        rp=row._tr.get_or_add_trPr();etree.SubElement(rp,q('cantSplit'))
        if i==0:etree.SubElement(rp,q('tblHeader'))
        for j,c in enumerate(row.cells):
            c.width=Inches(widths[j]/1440);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cp=c._tc.get_or_add_tcPr();mar=etree.SubElement(cp,q('tcMar'))
            for k,x in [('top',45),('bottom',45),('left',55),('right',55)]:
                el=etree.SubElement(mar,q(k));el.set(q('w'),str(x));el.set(q('type'),'dxa')
            sh=etree.SubElement(cp,q('shd'));sh.set(q('fill'),'334155' if i==0 else ('F1F5F9' if i%2==0 else 'FFFFFF'))
            for p in c.paragraphs:
                p.paragraph_format.space_after=Pt(0);p.paragraph_format.space_before=Pt(0);p.paragraph_format.line_spacing=1.03
                p.paragraph_format.keep_with_next=(i==0)
                for r in p.runs:font_run(r,size,i==0,'FFFFFF' if i==0 else '000000')
    return t

def form_header(doc,title,date):
    t=table(doc,[title],[],[9360],16)
    for r in t.cell(0,0).paragraphs[0].runs:font_run(r,16,True,'FFFFFF')
    meta=doc.add_table(rows=1,cols=2);meta.autofit=False;meta.alignment=WD_TABLE_ALIGNMENT.CENTER
    widths=[6540,2820]
    for c,w,txt in zip(meta.rows[0].cells,widths,[f'Project Title: {PROJECT}',f'Date Prepared: {date}']):
        c.width=Inches(w/1440);c.text=txt
        for p in c.paragraphs:
            p.paragraph_format.space_after=Pt(4);p.paragraph_format.space_before=Pt(0);p.paragraph_format.keep_with_next=True
            for r in p.runs:font_run(r,8)
    meta.rows[0].cells[1].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.RIGHT
    for g,w in zip(meta._tbl.find(q('tblGrid')),widths):g.set(q('w'),str(w))
    return t

def base_doc():
    d=Document();s=d.sections[0];s.page_width=Inches(8.2677);s.page_height=Inches(11.6929)
    s.top_margin=s.bottom_margin=Inches(.65);s.left_margin=s.right_margin=Inches(.65)
    st=d.styles['Normal'];st.font.name='Arial';st.font.size=Pt(8);st.paragraph_format.space_after=Pt(4)
    for nm,size in [('Title',18),('Heading 1',13),('Heading 2',10)]:
        d.styles[nm].font.name='Arial';d.styles[nm].font.size=Pt(size);d.styles[nm].font.color.rgb=RGBColor(0,0,0)
    for border in list(d.styles.element.iter(q('pBdr'))):border.getparent().remove(border)
    return d

def append_stakeholder(doc, standalone=False):
    if not standalone:doc.element.body.insert(len(doc.element.body)-1,heading_par('8. CREATE STAKEHOLDER REGISTER'))
    form_header(doc,'STAKEHOLDER REGISTER','September 29, 2026')
    paragraph(doc,'Classification uses decision power and project interest. Contact channels are listed where verified addresses are unavailable; the PM records individual contact details before use. Review at each 1.7.5 session and after an approved change.',8)
    table(doc,['Name','Position/Role','Contact Information','Requirements','Expectations','Classification'],STAKES,[1260,1410,1360,1880,2100,1350])

def revised_risks():
    upper=[];lower=[]
    for n in [69,71,73,75]:upper += [r[:] for r in DATA['tables'][n] if re.fullmatch(r'R-\d\d',r[0])]
    for n in [70,72,74,76]:lower += [r[:] for r in DATA['tables'][n] if re.fullmatch(r'R-\d\d',r[0])]
    updates={
    'R-01':'Assumption Log A-04 to A-12; Project Schedule and milestone list; stakeholder contact/availability records.',
    'R-02':'Architecture/OpenAPI design; WBS Dictionary 1.3/1.5; PR-01; test plan/results; Schedule and cost forecast.',
    'R-03':'WBS/Dictionary; SRS, RTM and Scope Statement if features change; Schedule, resource plan and cost forecast.',
    'R-04':'SRS, Requirements Documentation, RTM and Scope Statement; WBS/Dictionary, Schedule and cost baseline if approved.',
    'R-05':'Assumption Log A-13; store submission checklist; release/handover dates in the Schedule.',
    'R-06':'Assumption Log A-14; procurement/service activation records; Resource Requirements and Schedule.',
    'R-07':'PR-04; monthly cost forecast and status report; cost baseline only if a change is approved.',
    'R-08':'Defect log; security checklist/test report; QMP and WBS Dictionary 1.5 if acceptance/process changes; Schedule/cost forecast.',
    'R-09':'Requirements Documentation, SRS and RTM (FR-01/02/08); privacy notice; design/security/store checklists.',
    'R-10':'Team Charter backup list; Resource Requirements; Schedule/assignments; affected WBS Dictionary and cost forecast.',
    'R-11':'Flutter design/code; device test report and defect log; WBS Dictionary 1.3/1.5; Schedule/cost forecast.',
    'R-12':'Deployment/operations guide; backup/restore configuration and drill report; release checklist and Schedule.',
    }
    for row in lower:
        row[10]+=' Update if triggered: '+updates[row[0]]
    # Make the existing staffing response feasible instead of assuming pooled hours are interchangeable.
    lower[9][8]='1. Record one qualified backup per WP in the PMP/Team Charter (1.7.1).\n2. Confirm handover and individual free dates in 1.7.5; keep one active owner per WP.\n3. Before reassignment, check skill, dependencies, no task overlap and <= 176 h/month; approve changed assignments/dates/costs through change control.'
    upper += [
    ['R-13','If backend/API decisions or code reviews wait more than 1 working day for the Technical Lead, dependent development may stop, delaying 1.3.8 or 1.4.8.',MINH,'3','1','2','3','2','9','Mitigate: use the shared API contract (1.2.6), identify review cover in 1.7.1 and check individual workload in 1.7.4/1.7.5.'],
    ['R-14','If a developer lacks the skills required by an assigned package, code may fail acceptance or introduce repeated defects, increasing rework and delaying the release candidate (1.5.8).',HIEU,'3','1','4','3','2','12','Mitigate: check skills before assignment; review code and developer unit tests within the development WP; check integration in 1.3.5/1.4.5.']]
    lower += [
    ['R-13','2','1','1','2','1','4',MINH,'1. Check review queues and remaining effort in the Monday meeting and 1.7.4/1.7.5.\n2. If a review waits > 1 working day, assign a qualified reviewer with free capacity; another developer reviews the Technical Lead\'s own code.\n3. Replan affected dependencies and costs through change control; keep one task per person and <= 176 h/month.','Open','Trigger: review/API decision waiting > 1 working day, or remaining effort exceeds individual capacity before due date. Update: Team Charter review cover; Resource Requirements; Schedule; affected WBS Dictionary and cost forecast.'],
    ['R-14','2','1','3','2','1','6',HIEU,'1. Match assignments to Resource Requirements skills before work starts.\n2. In development WPs, require a peer-approved PR, passing unit tests and backend coverage >= 70%; reject failed submissions.\n3. On S1/S2 defects or >= 3 defects with one root cause, use the QMP root cause process; revise the assignment/checklist and use approved defect-fix WPs. Extra work needs change approval.','Open','Trigger: failed quality gate, S1/S2 defect, or >= 3 defects with one root cause. Update: Resource Requirements/Team Charter; QMP/checklists; defect log/test cases; affected WBS Dictionary, Schedule and cost forecast.']]
    short_statements={
    'R-01':'If the Customer or Sponsor misses a review or sign-off (A-04 to A-12), approval may be late, delaying dependent WPs and milestones 1.7.2, 1.2.11, 1.6.3 or 1.6.12.',
    'R-02':'If WebSocket scaling in 1.3.3 is insufficient, the Nov 12-20 load test (1.5.3) may miss 15,000 users, error rate < 1% or p90 < 200 ms, requiring rework before RC (1.5.8).',
    'R-03':'If development in 1.3/1.4 exceeds its hours (up to +40%, Medium-Low estimate confidence), 1.4.8 (Dec 7) may slip, reducing Could-have scope or time for testing.',
    'R-04':'If the Customer changes requirements after SRS sign-off (1.1.8, Sep 18), design and development rework may increase cost and delay delivery.',
    'R-05':'If store review exceeds 3 working days or rejects a build (A-13), publication may miss Jan 15, 2027 and delay 1.6.12.',
    'R-06':'If cloud/CDN activation exceeds 2 working days after purchase (A-14), setup (1.2.7, Oct 13-20) may start late, delaying 1.3.2 and the MVP.',
    'R-07':'If streaming traffic exceeds 5,000 GB/month (PR-04), CDN overage may increase the cost of 1.2.7.',
    'R-08':'If ZAP or ASVS/MASVS checks in 1.5.3 (Nov 12-20) find a critical/high vulnerability, fixes may exceed the 5 days in 1.5.4, delaying RC and failing NFR-05.',
    'R-10':'If a member is absent for > 3 working days, their singly owned WPs may stop and delay dependent milestones.',
    'R-11':'If Flutter misses device targets (cold < 2 s, warm < 1 s, battery < 6%/h; NFR-03/04), rework after 1.5.3 may delay RC (1.5.8).',
    'R-12':'If the restore drill before go-live (1.6.4) exceeds RTO 4 h or RPO 24 h (NFR-10), production deployment may be postponed, delaying 1.6.12.'}
    for r in upper:
        if r[0] in short_statements:r[1]=short_statements[r[0]]
    for r in lower:r[10]=r[10].replace(' Update if triggered: ',' Updates if triggered: ')
    return upper,lower

def shorten_repetition(d):
    b=d.element.body
    common='One responsible person per work package, working on one task at a time; hours = duration x 8 h (8-80 h each).'
    for t in b.findall(q('tbl')):
        first=cell_text(cells(rows(t)[0])[0])
        if first=='WBS DICTIONARY':replace_all(t,' '+common,'')
        if first=='RESOURCE REQUIREMENTS':
            for p in t.iter(q('p')):
                text=p_text(p)
                text2=re.sub(r' Skills: .*\.$','',text)
                text2=text2.replace(', one task at a time.','.')
                if text2!=text:replace_in_p(p,text,text2)
            rp=first_rpr(cells(rows(t)[1])[0])
            note=make_par('Common rules: one active WP per person; 8 h per day; one owner per WP. Skills by role: PM - facilitation/planning/reporting; BA/Full-Stack - requirements/Node.js/Flutter; Technical Lead - Node.js/PostgreSQL/WebSocket; QA/DevOps - test design/k6/ZAP/Docker/CI; UI/Flutter - Figma/HIG/Material 3/Flutter. These skills apply to every assignment below.',rp)
            t.addprevious(note)
    # Keep the shared Dictionary rules once rather than repeating them in all seven accounts.
    for p in b.iter(q('p')):
        if p_text(p).startswith('Dictionary tailored at control-account level'):
            replace_in_p(p,'they are not single work packages.','they are not single work packages. Each WP has one owner, 8-80 h at 8 h/day; one active task per person. Costs and dates are detailed in the associated estimate/resource/schedule records.')

def risk_section(doc):
    form_header(doc,'RISK REGISTER','October 8, 2026')
    paragraph(doc,'Probability and Impact Scale (1-5) used in this register',9,True,True)
    legend=DATA['tables'][77]
    table(doc,legend[1],legend[2:7],[500,900,1820,1820,2020,2300],7)
    paragraph(doc,legend[7][0],7.5)
    paragraph(doc,'Revised ratings are target residual ratings after the actions, not measured results. All risks remain Open. PM authority requires BOTH cost <= 15,140,000 VND AND schedule delay <= 3 calendar days; exceeding either limit requires Sponsor approval.',7.5)
    paragraph(doc,'Planned prevention is included in the referenced WPs and their estimates. Triggered work beyond those estimates requires impact analysis and approval in 1.7.5, then updates to affected scope, schedule, resources and cost records. Reserves fund approved responses; they do not add unlimited effort to a WP.',7.5)
    u,l=revised_risks()
    paragraph(doc,'Risk Statements',9,True,True)
    table(doc,['ID','Risk Statement','Owner','Probability','Scope','Quality','Schedule','Cost','Score','Response'],u,[500,2570,1050,600,420,420,500,420,480,2400],7)
    paragraph(doc,'Actions',9,True,True)
    table(doc,['ID','Revised Probability','Scope','Quality','Schedule','Cost','Revised Score','Responsible Party','Actions','Status','Comments'],l,[500,630,420,420,500,420,570,1000,2500,570,1830],7)

def main():
    d=Document(SRC);b=d.element.body
    # Date Prepared carries a single date, using the latest existing preparation/revision date.
    for tc in b.iter(q('tc')):
        for p in tc.findall(q('p')):
            t=p_text(p)
            if 'Date Prepared:' in t and '(revised ' in t:
                m=re.search(r'Date Prepared:.*?\(revised ([A-Z][a-z]+ \d{1,2}, \d{4})[^)]*\)',t)
                if m:replace_in_p(p,m.group(0),'Date Prepared: '+m.group(1))
    # Remove only event rows in the WBS; do not alter schedule IDs, logic or baseline dates.
    wt=next(t for t in b.findall(q('tbl')) if cell_text(cells(rows(t)[0])[0])=='WORK BREAKDOWN STRUCTURE')
    for p in list(wt.iter(q('p'))):
        if re.match(r'1\.\d\.\d+\.',p_text(p)) and '(milestone)' in p_text(p):p.getparent().remove(p)
    last=next(p for p in wt.iter(q('p')) if p_text(p).startswith('Level 2 ='))
    for child in list(last):
        if child.tag!=q('pPr'):last.remove(child)
    last.append(make_par('Level 2 = phase control accounts for scope, schedule and cost reporting; level 3 = 55 work packages, each with one owner and 8-80 h. Milestones are zero-duration schedule events, listed in the Charter/Schedule and Dictionary milestone fields, not work packages. Existing codes are retained for traceability. Customer UAT is outside the WBS. Recurring WPs 1.7.4/1.7.5 use dated sessions. Development WPs include developer unit tests. The Dictionary summarizes each control account and lists its work packages.').find(q('r')))
    dictionaries=[t for t in b.findall(q('tbl')) if cell_text(cells(rows(t)[0])[0])=='WBS DICTIONARY']
    for t in dictionaries:
        tc=cells(rows(t)[2])[0]
        set_cell_text(tc,cell_text(tc).replace('Work Package Name:','Control Account Name:'))
    note=make_par('Dictionary tailored at control-account level (Book of Forms, Section 2.10, pp. 52-53). Codes 1.1-1.7 summarize the work packages listed below; they are not single work packages.',first_rpr(cells(rows(dictionaries[0])[1])[0]))
    dictionaries[0].addprevious(note)
    replace_all(b,'R-01 to R-12','R-01 to R-14')
    # Include the new register in existing management/planning scope, without creating effort.
    replace_all(b,'Project Schedule, Assumption Log, Risk Register','Project Schedule, Assumption Log, Risk Register, Stakeholder Register')
    replace_all(b,'with the risk register (R-01 to R-14)','with the risk register (R-01 to R-14) and Stakeholder Register')
    shorten_repetition(d)
    # Replace the complete risk section body with one scale and two continuous datasets.
    heading=next(p for p in b.findall(q('p')) if p_text(p)=='7. CREATE RISK REGISTER')
    for el in list(b)[list(b).index(heading)+1:]:
        if el.tag!=q('sectPr'):b.remove(el)
    risk_section(d)
    # End the risk form using the same page system as other forms.
    br=section_break_par(b.find(q('sectPr')))
    ppr=br.find(q('pPr'));sp=etree.SubElement(ppr,q('spacing'));sp.set(q('after'),'0');sp.set(q('before'),'0')
    rp=etree.SubElement(ppr,q('rPr'));etree.SubElement(rp,q('sz')).set(q('val'),'2')
    b.insert(len(b)-1,br)
    append_stakeholder(d)
    d.save(TMP/'4.revised.docx')
    standalone=base_doc();append_stakeholder(standalone,True)
    fp=standalone.sections[0].footer.paragraphs[0];fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    font_run(fp.add_run('Page '),8)
    fld=etree.SubElement(fp._p,q('fldSimple'));fld.set(q('instr'),'PAGE')
    font_run(fp.add_run(' of '),8)
    fld=etree.SubElement(fp._p,q('fldSimple'));fld.set(q('instr'),'NUMPAGES')
    standalone.save(ROOT/'Stakeholder_Register.docx')
    (TMP/'risks.json').write_text(json.dumps(revised_risks(),ensure_ascii=False,indent=2),encoding='utf-8')
    print('Saved revised report and Stakeholder Register; schedule baseline unchanged.')

if __name__=='__main__':main()
