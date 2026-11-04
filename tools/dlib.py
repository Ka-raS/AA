"""Helpers for editing WordprocessingML with lxml (run-aware replace, cell rewrite, cloning)."""
import copy, re
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

def q(tag):
    return W + tag

def texts(p):
    return [t for t in p.iter(q("t"))]

def p_text(p):
    return "".join(t.text or "" for t in p.iter(q("t")))

def el_text(e):
    return "".join(t.text or "" for t in e.iter(q("t")))

def replace_in_p(p, old, new):
    ts = texts(p)
    if not ts:
        return 0
    n = 0
    pos = 0
    while True:
        full = "".join(t.text or "" for t in ts)
        i = full.find(old, pos)
        if i < 0:
            break
        j = i + len(old)
        # locate t nodes
        starts = []
        acc = 0
        for t in ts:
            starts.append(acc)
            acc += len(t.text or "")
        a = None
        for k in range(len(ts)):
            if starts[k] <= i < starts[k] + len(ts[k].text or ""):
                a = k; break
        b = a
        for k in range(len(ts)):
            if starts[k] < j <= starts[k] + len(ts[k].text or ""):
                b = k; break
        ta, tb = ts[a], ts[b]
        ia, jb = i - starts[a], j - starts[b]
        if a == b:
            ta.text = (ta.text or "")[:ia] + new + (ta.text or "")[jb:]
        else:
            ta.text = (ta.text or "")[:ia] + new
            for k in range(a + 1, b):
                ts[k].text = ""
            tb.text = (tb.text or "")[jb:]
        for t in (ta, tb):
            t.set(XML_SPACE, "preserve")
        n += 1
        pos = i + len(new)
    return n

def replace_all(root, old, new):
    n = 0
    for p in root.iter(q("p")):
        if old in p_text(p):
            n += replace_in_p(p, old, new)
    return n

class Repl:
    """Collects replacements and asserts expected counts."""
    def __init__(self, root, name):
        self.root, self.name, self.log = root, name, []
    def __call__(self, old, new, expect=1):
        n = replace_all(self.root, old, new)
        self.log.append((n, expect, old[:70]))
        if expect is not None and n != expect:
            raise AssertionError(f"[{self.name}] expected {expect} got {n}: {old[:90]!r}")
        return n

def rows(tbl):
    return tbl.findall(q("tr"))

def cells(tr):
    return tr.findall(q("tc"))

def cell_text(tc):
    return "\n".join(p_text(p) for p in tc.findall(q("p")))

def _mk_run(rpr, text, bold=None):
    r = etree.Element(q("r"))
    if rpr is not None:
        rp = copy.deepcopy(rpr)
        if bold is not None:
            for tag in ("b", "bCs"):
                x = rp.find(q(tag))
                if x is not None:
                    rp.remove(x)
            if bold:
                b = etree.Element(q("b"))
                # keep schema order roughly: rFonts first, then b
                fonts = rp.find(q("rFonts"))
                rStyle = rp.find(q("rStyle"))
                anchor = fonts if fonts is not None else rStyle
                if anchor is not None:
                    anchor.addnext(b)
                else:
                    rp.insert(0, b)
        r.append(rp)
    elif bold:
        rp = etree.SubElement(r, q("rPr")); etree.SubElement(rp, q("b"))
    t = etree.SubElement(r, q("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")
    return r

def first_rpr(tc_or_p):
    for r in tc_or_p.iter(q("r")):
        if r.find(q("t")) is not None:
            rp = r.find(q("rPr"))
            return rp
    for r in tc_or_p.iter(q("r")):
        return r.find(q("rPr"))
    # fall back to paragraph mark run properties
    for p in tc_or_p.iter(q("p")):
        ppr = p.find(q("pPr"))
        if ppr is not None and ppr.find(q("rPr")) is not None:
            return ppr.find(q("rPr"))
    return None

def set_cell_lines(tc, lines, rpr=None, ppr=None):
    """lines: list of str or list of (text, bold) segments. Keeps first paragraph formatting."""
    ps = tc.findall(q("p"))
    tmpl = ps[0]
    if rpr is None:
        rpr = first_rpr(tc)
    if ppr is None:
        ppr = tmpl.find(q("pPr"))
    for p in ps:
        tc.remove(p)
    for line in lines:
        p = etree.SubElement(tc, q("p"))
        if ppr is not None:
            p.append(copy.deepcopy(ppr))
        segs = [(line, None)] if isinstance(line, str) else line
        for text, bold in segs:
            p.append(_mk_run(rpr, text, bold))
    return tc

def set_cell_text(tc, text, **kw):
    return set_cell_lines(tc, text.split("\n"), **kw)

def make_par(text="", rpr=None, ppr=None, bold=None):
    p = etree.Element(q("p"))
    if ppr is not None:
        p.append(copy.deepcopy(ppr))
    if text:
        p.append(_mk_run(rpr, text, bold))
    return p

def page_break_par():
    p = etree.Element(q("p"))
    r = etree.SubElement(p, q("r"))
    br = etree.SubElement(r, q("br"))
    br.set(q("type"), "page")
    return p

def section_break_par(body_sectpr):
    """Empty paragraph that ends a section (next page) with the page setup of body_sectpr."""
    p = etree.Element(q("p"))
    ppr = etree.SubElement(p, q("pPr"))
    sp = copy.deepcopy(body_sectpr)
    for x in list(sp):
        if x.tag in (q("headerReference"), q("footerReference"), q("type"), q("titlePg")):
            sp.remove(x)
    typ = etree.Element(q("type")); typ.set(q("val"), "nextPage")
    sp.insert(0, typ)
    ppr.append(sp)
    return p

def heading_par(text):
    """Group heading (outline level 1), Arial bold 13 pt."""
    p = etree.Element(q("p"))
    ppr = etree.SubElement(p, q("pPr"))
    etree.SubElement(ppr, q("keepNext"))
    sp = etree.SubElement(ppr, q("spacing")); sp.set(q("before"), "0"); sp.set(q("after"), "120")
    ol = etree.SubElement(ppr, q("outlineLvl")); ol.set(q("val"), "0")
    r = etree.SubElement(p, q("r"))
    rpr = etree.SubElement(r, q("rPr"))
    f = etree.SubElement(rpr, q("rFonts"))
    for a in ("ascii", "hAnsi", "cs", "eastAsia"):
        f.set(q(a), "Arial")
    etree.SubElement(rpr, q("b"))
    sz = etree.SubElement(rpr, q("sz")); sz.set(q("val"), "26")
    t = etree.SubElement(r, q("t")); t.text = text
    return p

def find_row(tbl, first_cell_text):
    for tr in rows(tbl):
        cs = cells(tr)
        if cs and cell_text(cs[0]).strip() == first_cell_text:
            return tr
    raise KeyError(first_cell_text)

def set_grid(tbl, widths):
    """Set gridCol widths and each cell tcW from the grid (handles gridSpan)."""
    grid = tbl.find(q("tblGrid"))
    gcs = grid.findall(q("gridCol"))
    assert len(gcs) == len(widths), (len(gcs), len(widths))
    for g, w in zip(gcs, widths):
        g.set(q("w"), str(int(w)))
    for tr in rows(tbl):
        col = 0
        for tc in cells(tr):
            tcpr = tc.find(q("tcPr"))
            span = 1
            if tcpr is not None and tcpr.find(q("gridSpan")) is not None:
                span = int(tcpr.find(q("gridSpan")).get(q("val")))
            w = sum(widths[col:col + span])
            col += span
            if tcpr is not None:
                tcw = tcpr.find(q("tcW"))
                if tcw is not None:
                    tcw.set(q("w"), str(int(w))); tcw.set(q("type"), "dxa")
