import sys, pymupdf
sys.stdout.reconfigure(encoding="utf-8")
doc = pymupdf.open(sys.argv[1])
print("pages:", doc.page_count)
spec = sys.argv[2]
pages = []
for part in spec.split(","):
    if "-" in part:
        a, b = part.split("-"); pages += list(range(int(a), int(b) + 1))
    else:
        pages.append(int(part))
for p in pages:
    print(f"\n======== PAGE {p} ========")
    print(doc[p - 1].get_text())
