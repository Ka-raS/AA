"""Re-balance table column widths so long-text columns get the room and 1-character columns stay narrow.
Measures text with the real Arial metrics, wraps it like Word (word by word) and greedily gives width to the
column that shortens the table most. Tables of one form share one set of widths.
usage: python colwidth.py <in.docx> <out.docx> [--dry]
"""
import sys, os, math, copy
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from PIL import ImageFont
from dlib import *

FONT = {False: ImageFont.truetype("C:/Windows/Fonts/arial.ttf", 100), True: ImageFont.truetype("C:/Windows/Fonts/arialbd.ttf", 100)}
_cache = {}
def tw(text, size, bold):
    """text width in points"""
    k = (text, size, bold)
    if k not in _cache:
        _cache[k] = FONT[bold].getlength(text) * size / 100.0
    return _cache[k]

def run_props(r, default=9.0):
    rp = r.find(q("rPr"))
    size, bold = default, False
    if rp is not None:
        s = rp.find(q("sz"))
        if s is not None: size = int(s.get(q("val"))) / 2
        b = rp.find(q("b"))
        bold = b is not None and b.get(q("val")) not in ("0", "false")
    return size, bold

def para_words(p):
    """words of a paragraph as lists of (text, size, bold) pieces, with glue flags:
    a word that follows a hyphen of the previous word has no space before it (Word may break after '-')."""
    words, glue, cur, sizes = [], [], [], []
    nospace = False
    def flush(next_glued):
        nonlocal cur, nospace
        if cur:
            words.append(cur); glue.append(nospace); cur = []
        nospace = next_glued
    for r in p.iter(q("r")):
        size, bold = run_props(r)
        for t in r.iter(q("t")):
            txt = t.text or ""
            if txt.strip(): sizes.append(size)
            for i, part in enumerate(txt.split(" ")):
                if i > 0:
                    flush(False)
                bits = part.split("-")
                for j, bit in enumerate(bits):
                    piece = bit + ("-" if j < len(bits) - 1 else "")
                    if piece:
                        cur.append((piece, size, bold))
                    if j < len(bits) - 1 and bit and bits[j + 1]:
                        flush(True)
    flush(False)
    return words, glue, (max(sizes) if sizes else 7.5)

def is_num(word_pieces):
    txt = "".join(p[0] for p in word_pieces)
    return any(ch.isdigit() for ch in txt) and all(ch.isdigit() or ch in ",.%-+/()MhVND<>=x" for ch in txt)

class Cell:
    """paras: list of (tokens, size) where tokens = [(width_pt, gap_pt)]"""
    def __init__(self, tc, col, span, pad):
        self.col, self.span, self.pad, self.tc = col, span, pad, tc
        self.paras, caps = [], []
        for p in tc.findall(q("p")):
            words, glue, size = para_words(p)
            sp = tw(" ", size, False)
            toks = []
            full, fx = [], 0.0                      # the space-delimited word being built from hyphen parts
            def close():
                if full:
                    if fx <= 45 or (is_num(full) and "-" not in "".join(x[0] for x in full) and fx <= 90):
                        caps.append(fx)              # short words and plain numbers must fit whole
                    else:
                        caps.extend(v for v in part_w if v <= 45 or v <= 90 and is_num(full))
            part_w = []
            for wd, g in zip(words, glue):
                x = sum(tw(*piece) for piece in wd)
                toks.append((x, 0.0 if g else sp))
                if not g:
                    close(); full, fx, part_w = [], 0.0, []
                full = full + wd; fx += x; part_w.append(x)
            close()
            self.paras.append((toks, size))
        self.longest_cap = max(caps or [0])
        txt = cell_text(tc).strip()
        self.short = 0 < len(txt) <= 24 and not any(ch.isdigit() for ch in txt) and all(len(t) <= 4 for t, _ in self.paras)
        bold = any(run_props(r)[1] for r in tc.iter(q("r")) if (r.find(q("t")) is not None))
        # tiny plain values such as "3 months", "4 devices", "Open": one line
        self.tiny = (not bold and 0 < len(txt) <= 12 and len(self.paras) == 1 and len(self.paras[0][0]) <= 2
                     and txt.count(" ") <= 1)
        self.one_line = sum(w + g for w, g in self.paras[0][0]) * 20 + pad if self.paras and self.paras[0][0] else 0
    def lines_par(self, par, width_dxa):
        avail = max(width_dxa - self.pad, 40) / 20.0
        toks, size = par
        lines, cur = 1, 0.0
        for w, gap in toks:
            if w > avail:                       # Word breaks a too-long word by characters
                if cur > 0: lines += 1
                extra = math.ceil(w / avail) - 1
                lines += extra; cur = w - extra * avail
                continue
            need = w if cur == 0 else cur + gap + w
            if need <= avail: cur = need
            else: lines += 1; cur = w
        return lines
    def max_lines(self, width_dxa):
        return max([self.lines_par(p, width_dxa) for p in self.paras] or [0])
    def two_line_width(self):
        lo, hi = 200, 4000
        while hi - lo > 20:
            mid = (lo + hi) // 2
            if self.max_lines(mid) <= 2: hi = mid
            else: lo = mid
        return hi
    def height(self, width_dxa):
        return sum(self.lines_par(p, width_dxa) * p[1] * 1.15 for p in self.paras)
    def soft_height(self, width_dxa):
        avail = max(width_dxa - self.pad, 40) / 20.0
        return sum(max(1.0, sum(w + g for w, g in toks) / avail) * size * 1.15 for toks, size in self.paras)

def table_cells(t):
    pr = t.find(q("tblPr"))
    mar = pr.find(q("tblCellMar")) if pr is not None else None
    pad = 216
    if mar is not None:
        l = mar.find(q("left")); r = mar.find(q("right"))
        pad = (int(l.get(q("w"))) if l is not None else 108) + (int(r.get(q("w"))) if r is not None else 108)
    out = []
    for tr in rows(t):
        rc, col = [], 0
        for tc in cells(tr):
            tp = tc.find(q("tcPr")); span = 1
            if tp is not None and tp.find(q("gridSpan")) is not None:
                span = int(tp.find(q("gridSpan")).get(q("val")))
            cpad = pad
            if tp is not None and tp.find(q("tcMar")) is not None:
                m = tp.find(q("tcMar")); l = m.find(q("left")); r = m.find(q("right"))
                cpad = (int(l.get(q("w"))) if l is not None else pad // 2) + (int(r.get(q("w"))) if r is not None else pad // 2)
            rc.append(Cell(tc, col, span, cpad + 10))   # +10 dxa safety for Word's rounding
            col += span
        out.append(rc)
    return out

def total_height(rows_, w):
    pref = [0]
    for x in w: pref.append(pref[-1] + x)
    return sum(max([c.height(pref[c.col + c.span] - pref[c.col]) for c in rc] or [0]) for rc in rows_)

def surrogate(rows_, w, p=6):
    pref = [0]
    for x in w: pref.append(pref[-1] + x)
    tot = 0.0
    for rc in rows_:
        acc = sum(c.soft_height(pref[c.col + c.span] - pref[c.col]) ** p for c in rc)
        tot += acc ** (1.0 / p)
    return tot

def hill(rows_, w, mins, cur):
    n = len(w)
    for delta in (600, 300, 150, 80, 40, 20):
        improved = True
        while improved:
            improved = False
            for i in range(n):
                for j in range(n):
                    if i == j or w[j] - delta < mins[j]: continue
                    w[i] += delta; w[j] -= delta
                    h = total_height(rows_, w)
                    if h < cur - 1e-6:
                        cur = h; improved = True
                    else:
                        w[i] -= delta; w[j] += delta
    return w, cur

def optimize(rows_, n, total, start=None):
    mins = [300] * n
    for rc in rows_:
        for c in rc:
            if c.span == 1:
                mins[c.col] = max(mins[c.col], int(math.ceil(c.longest_cap * 20 + c.pad)))
                if c.short:
                    mins[c.col] = max(mins[c.col], c.two_line_width())
                if c.tiny:
                    mins[c.col] = max(mins[c.col], int(math.ceil(c.one_line)))
    if sum(mins) > total:
        f = total / sum(mins); mins = [int(m * f) for m in mins]
    # 1. greedy on a smooth surrogate (soft max of cell heights per row) from the minimum widths
    w = mins[:]
    left = total - sum(w)
    step = 40
    while left > 0:
        st = min(step, left)
        base = surrogate(rows_, w)
        best = None
        for i in range(n):
            w[i] += st
            g = base - surrogate(rows_, w)
            w[i] -= st
            if best is None or g > best[0]: best = (g, i)
        w[best[1]] += st; left -= st
    cands = [w]
    if start is not None and len(start) == n:
        s0 = [max(a, b) for a, b in zip(start, mins)]
        extra = sum(s0) - total
        while extra > 0:                           # take the excess from the widest columns
            k = max(range(n), key=lambda i: s0[i] - mins[i]); take = min(extra, s0[k] - mins[k]); s0[k] -= take; extra -= take
        s0[max(range(n), key=lambda i: s0[i])] += total - sum(s0)
        cands.append(s0)
    best = None
    for c in cands:
        c, h = hill(rows_, c[:], mins, total_height(rows_, c))
        if best is None or h < best[1]: best = (c, h)
    return best

def set_span(tc, span):
    tp = tc.find(q("tcPr"))
    if tp is None:
        tp = etree.Element(q("tcPr")); tc.insert(0, tp)
    g = tp.find(q("gridSpan"))
    if span == 1:
        if g is not None: tp.remove(g)
        return
    if g is None:
        g = etree.Element(q("gridSpan"))
        tcw = tp.find(q("tcW"))
        if tcw is not None: tcw.addnext(g)
        else: tp.insert(0, g)
    g.set(q("val"), str(span))

def rebalance_pairs(t, w):
    n = len(w)
    pref = [0]
    for x in w: pref.append(pref[-1] + x)
    for tr, rc in zip(rows(t), table_cells(t)):
        if len(rc) != 2 or rc[0].span + rc[1].span != n or n < 3:
            continue
        if any(c.tc.find(q("tcPr")) is not None and c.tc.find(q("tcPr")).find(q("vMerge")) is not None for c in rc):
            continue
        a, b = rc
        k0 = a.span
        best = None
        for k in range(1, n):
            h = max(a.height(pref[k]), b.height(pref[n] - pref[k]))
            key = (round(h, 1), abs(k - k0))
            if best is None or key < best[0]: best = (key, k)
        k = best[1]
        if k != k0:
            set_span(a.tc, k); set_span(b.tc, n - k)

def grid(t):
    return [int(g.get(q("w"))) for g in t.find(q("tblGrid")).findall(q("gridCol"))]

def apply(t, w):
    set_grid(t, w)
    pr = t.find(q("tblPr"))
    tw_ = pr.find(q("tblW"))
    if tw_ is None:
        tw_ = etree.SubElement(pr, q("tblW"))
    tw_.set(q("w"), str(sum(w))); tw_.set(q("type"), "dxa")
    lay = pr.find(q("tblLayout"))
    if lay is None:
        lay = etree.Element(q("tblLayout"))
        after = [pr.find(q(x)) for x in ("tblCellMar", "tblLook", "tblCaption", "tblDescription")]
        after = [a for a in after if a is not None]
        if after: after[0].addprevious(lay)
        else: pr.append(lay)
    lay.set(q("type"), "fixed")

if __name__ == "__main__":
    from lxml import etree
    SRC, DST = sys.argv[1], sys.argv[2]
    dry = "--dry" in sys.argv
    d = Document(SRC)
    body = d.element.body
    tbls = body.findall(q("tbl"))
    def first(t): return cell_text(cells(rows(t)[0])[0]).strip()
    groups, used = [], set()
    def group(pred, name):
        g = [t for t in tbls if id(t) not in used and pred(t)]
        if g:
            for t in g: used.add(id(t))
            groups.append((name, g))
    group(lambda t: first(t).startswith("REQUIREMENTS TRACEABILITY MATRIX"), "RTM")
    group(lambda t: first(t).startswith("WBS DICTIONARY") and "RESOURCE SUMMARY" not in first(t), "WBS Dictionary")
    group(lambda t: first(t).startswith("RISK REGISTER"), "Risk Register (upper)")
    group(lambda t: first(t) == "ID" and "Revised Probability" in el_text(rows(t)[0]), "Risk Register (lower)")
    group(lambda t: first(t).startswith(("Project budget (bottom-up)", "1.2.7 Cloud", "1.5.3 Reference", "1.5.3 Load-testing")), "Basis of Estimates sub-tables")
    for t in tbls:
        if id(t) in used: continue
        n = len(grid(t))
        ft = first(t)
        if n >= 3 and not ft.startswith(("Class:",)):
            used.add(id(t)); groups.append((ft[:40], [t]))
        elif ft.startswith("QUALITY MANAGEMENT PLAN") and n >= 2:   # Team Charter keeps 50/50 (signature space)
            used.add(id(t)); groups.append((ft[:40], [t]))
    for name, g in groups:
        n = len(grid(g[0])); total = sum(grid(g[0]))
        assert all(len(grid(t)) == n for t in g), name
        rows_ = [rc for t in g for rc in table_cells(t)]
        before = total_height(rows_, grid(g[0]))
        w, after = optimize(rows_, n, total, start=grid(g[0]))
        print(f"{name:42s} {before:8.0f} -> {after:8.0f} pt  {grid(g[0])} -> {w}")
        if not dry:
            for t in g:
                rebalance_pairs(t, w)
                apply(t, w)
    if not dry:
        d.save(DST)
