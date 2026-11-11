"""JCLC three-line tables: a line above the table, a line under the header row, a line under the last row.
No vertical lines and no other horizontal lines."""
from docx.oxml import parse_xml
from docx.oxml.ns import qn, nsdecls

SIDES = ("top", "left", "bottom", "right")
LINE = '<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
NIL = '<w:{s} w:val="nil"/>'


def _header_rows(rows):
    n = 0
    for tr in rows:
        trpr = tr.find(qn("w:trPr"))
        if trpr is not None and trpr.find(qn("w:tblHeader")) is not None:
            n += 1
        else:
            break
    return max(n, 1)


def _line_below(rows):
    h = _header_rows(rows)
    last = len(rows) - 1
    return [(i == h - 1) or (i == last) for i in range(len(rows))]


def _edge(el, side):
    if el is None:
        return None
    e = el.find(qn("w:" + side))
    return None if e is None else e.get(qn("w:val"))


def _is_line(v):
    return v not in (None, "nil", "none")


def check_table(tbl):
    """Return a list of problems (empty if the table is a proper three-line table)."""
    rows = tbl.findall(qn("w:tr"))
    if not rows:
        return []
    below = _line_below(rows)
    problems = []
    for ri, tr in enumerate(rows):
        for tc in tr.findall(qn("w:tc")):
            pr = tc.find(qn("w:tcPr"))
            cb = pr.find(qn("w:tcBorders")) if pr is not None else None
            if cb is None:
                continue
            top = _is_line(_edge(cb, "top")); bot = _is_line(_edge(cb, "bottom"))
            want_top = ri == 0 or below[ri - 1]
            if _is_line(_edge(cb, "left")) or _is_line(_edge(cb, "right")):
                problems.append("vertical line"); break
            if bot != below[ri]:
                problems.append(f"row {ri+1}: bottom line should be {'present' if below[ri] else 'absent'}"); break
            if top != want_top:
                problems.append(f"row {ri+1}: top line should be {'present' if want_top else 'absent'}"); break
    tpr = tbl.find(qn("w:tblPr")); tb = tpr.find(qn("w:tblBorders")) if tpr is not None else None
    for ri, tr in enumerate(rows):
        ex = tr.find(qn("w:tblPrEx")); eb = ex.find(qn("w:tblBorders")) if ex is not None else None
        if eb is not None:
            problems.append(f"row {ri+1}: row-level border override"); break
    if tb is not None and any(_is_line(_edge(tb, s)) for s in ("left", "right", "insideH", "insideV")):
        problems.append("table-level inner/vertical lines")
    return problems


def normalise_table(tbl):
    rows = tbl.findall(qn("w:tr"))
    if not rows:
        return False
    before = bool(check_table(tbl))
    below = _line_below(rows)
    tpr = tbl.find(qn("w:tblPr"))
    if tpr is None:
        return False
    old = tpr.find(qn("w:tblBorders"))
    if old is not None:
        tpr.remove(old)
    xml = (f'<w:tblBorders {nsdecls("w")}>' + LINE.format(s="top") + NIL.format(s="left") + LINE.format(s="bottom")
           + NIL.format(s="right") + NIL.format(s="insideH") + NIL.format(s="insideV") + "</w:tblBorders>")
    new = parse_xml(xml)
    nxt = None
    for tag in ("shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription"):
        nxt = tpr.find(qn("w:" + tag))
        if nxt is not None:
            break
    (nxt.addprevious(new) if nxt is not None else tpr.append(new))
    for tr in rows:
        ex = tr.find(qn("w:tblPrEx"))
        if ex is not None:
            eb = ex.find(qn("w:tblBorders"))
            if eb is not None:
                ex.remove(eb)
            if len(ex) == 0:
                tr.remove(ex)
    for ri, tr in enumerate(rows):
        top = ri == 0 or below[ri - 1]
        bot = below[ri]
        for tc in tr.findall(qn("w:tc")):
            pr = tc.find(qn("w:tcPr"))
            if pr is None:
                pr = parse_xml(f'<w:tcPr {nsdecls("w")}/>'); tc.insert(0, pr)
            cb = pr.find(qn("w:tcBorders"))
            if cb is not None:
                pr.remove(cb)
            cb = parse_xml(f'<w:tcBorders {nsdecls("w")}>' + (LINE if top else NIL).format(s="top") + NIL.format(s="left")
                           + (LINE if bot else NIL).format(s="bottom") + NIL.format(s="right") + "</w:tcBorders>")
            anchor = None
            for tag in ("vAlign", "noWrap", "tcMar", "textDirection", "tcFitText", "hideMark", "shd"):
                anchor = pr.find(qn("w:" + tag))
                if anchor is not None:
                    break
            if anchor is not None and anchor.tag == qn("w:shd"):
                anchor.addprevious(cb)
            elif anchor is not None:
                # shd (if any) comes after tcBorders, before noWrap/tcMar/vAlign
                shd = pr.find(qn("w:shd"))
                (shd if shd is not None else anchor).addprevious(cb)
            else:
                pos = 0
                for i, ch in enumerate(list(pr)):
                    if ch.tag in (qn("w:cnfStyle"), qn("w:tcW"), qn("w:gridSpan"), qn("w:hMerge"), qn("w:vMerge")):
                        pos = i + 1
                pr.insert(pos, cb)
    return before
