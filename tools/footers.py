"""Give every form section (all but the cover) a centered 'Page {PAGE} of {SECTIONPAGES}' footer and restart numbering."""
import sys, copy
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, __file__.rsplit("\\", 1)[0])
from docx import Document
from lxml import etree
from dlib import *

SRC, DST = sys.argv[1], sys.argv[2]
d = Document(SRC)

def fld(instr, cached):
    rpr = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="17"/></w:rPr>'
    return (f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>{rpr}<w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>{rpr}<w:t>{cached}</w:t></w:r><w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>')
def txt(s):
    return f'<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:sz w:val="17"/></w:rPr><w:t xml:space="preserve">{s}</w:t></w:r>'
FOOT_P = (f'<w:p xmlns:w="{W_NS}"><w:pPr><w:jc w:val="center"/></w:pPr>' + txt("Page ") + fld("PAGE", "1") + txt(" of ") + fld("SECTIONPAGES", "1") + "</w:p>")

ORDER = ["headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols"]
for i, s in enumerate(d.sections):
    sp = s._sectPr
    s.different_first_page_header_footer = False
    # numbering restarts at 1 in every section
    for x in sp.findall(q("pgNumType")):
        sp.remove(x)
    if i > 0:
        pn = etree.Element(q("pgNumType")); pn.set(q("start"), "1")
        # insert before the first element that must follow pgNumType
        after = [sp.find(q(t)) for t in ORDER[ORDER.index("pgNumType") + 1:]]
        after = [a for a in after if a is not None]
        if after: after[0].addprevious(pn)
        else: sp.append(pn)
    f = s.footer
    f.is_linked_to_previous = False
    fe = f._element
    for k in list(fe):
        fe.remove(k)
    if i == 0:
        fe.append(etree.fromstring(f'<w:p xmlns:w="{W_NS}"/>'))   # cover: empty footer
    else:
        fe.append(etree.fromstring(FOOT_P))
d.save(DST)
print("sections:", len(d.sections))
