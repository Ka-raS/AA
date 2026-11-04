"""Write the real page count of each section into the cached result of its SECTIONPAGES footer field,
so viewers that do not update fields (previews, some editors) show 'Page x of N' instead of 'of 1'.
Word recalculates the field anyway. Run after footers.py and topdf.ps1 on the same file.
usage: python pagecount.py <in.docx> <in.pdf> <out.docx>
"""
import sys, os, re
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pymupdf
from docx import Document
from dlib import *

SRC, PDF, DST = sys.argv[1], sys.argv[2], sys.argv[3]
counts = []
for page in pymupdf.open(PDF):
    m = None
    for line in page.get_text().splitlines():
        mm = re.fullmatch(r"\s*Page (\d+) of (\d+)\s*", line)
        if mm: m = mm
    if m and m.group(1) == "1":
        counts.append(int(m.group(2)))
d = Document(SRC)
secs = d.sections[1:]                      # the cover has no page number
assert len(secs) == len(counts), (len(secs), counts)
for s, n in zip(secs, counts):
    runs = list(s.footer._element.iter(q("r")))
    done = False
    for i, r in enumerate(runs):
        it = r.find(q("instrText"))
        if it is not None and it.text.strip() == "SECTIONPAGES":
            for r2 in runs[i + 1:]:
                fc = r2.find(q("fldChar"))
                if fc is not None and fc.get(q("fldCharType")) == "end":
                    break
                t = r2.find(q("t"))
                if t is not None:
                    t.text = str(n); done = True
    assert done, "SECTIONPAGES field not found"
d.save(DST)
print("section page counts:", counts)
