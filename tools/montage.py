import sys, pymupdf
# usage: montage.py file.pdf out.png first last cols dpi
f, out = sys.argv[1], sys.argv[2]
first, last = int(sys.argv[3]), int(sys.argv[4])
cols = int(sys.argv[5]) if len(sys.argv) > 5 else 4
dpi = int(sys.argv[6]) if len(sys.argv) > 6 else 50
doc = pymupdf.open(f)
last = min(last, doc.page_count)
pix = [doc[i - 1].get_pixmap(dpi=dpi) for i in range(first, last + 1)]
w, h = max(p.width for p in pix), max(p.height for p in pix)
n = len(pix); rows = (n + cols - 1) // cols
sheet = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, cols * (w + 6), rows * (h + 6)), False)
sheet.clear_with(120)
for k, p in enumerate(pix):
    r, c = divmod(k, cols)
    p2 = pymupdf.Pixmap(pymupdf.csRGB, p) if p.n != 3 or p.alpha else p
    sheet.copy(p2, pymupdf.IRect(c * (w + 6), r * (h + 6), c * (w + 6) + p2.width, r * (h + 6) + p2.height)) if False else None
    p2.set_origin(c * (w + 6), r * (h + 6))
    sheet.copy(p2, p2.irect)
sheet.save(out)
print(out, n, "pages")
