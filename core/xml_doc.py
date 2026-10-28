"""DART 원문 XML → 문서 순서 스트림(제목·문단·표)과 표 격자. 파서들은 여기서 받은 Node/Grid 만 본다."""
from dataclasses import dataclass, field
import re
import xml.etree.ElementTree as ET

from core import normalize as N
from xml_fix import fix_xml   # vendor (core/__init__ 이 sys.path 에 넣음)

_CELL_TAGS = ("TD", "TH", "TE", "TU")


@dataclass
class Node:
    kind: str        # "title" | "p" | "table"
    elem: object
    idx: int
    title: str       # 직전 TITLE 의 clean_text


@dataclass
class Grid:
    header: list = field(default_factory=list)   # list[list[str]]
    body: list = field(default_factory=list)
    ncols: int = 0
    border: str = "0"
    attrs: dict = field(default_factory=dict)

    def text(self):
        return N.clean_text(" ".join(c for r in self.header + self.body for c in r if c))

    def rows(self):
        return self.header + self.body


def load_xml(path):
    with open(path, encoding="utf-8", errors="replace") as f:
        raw = f.read()
    return ET.fromstring(fix_xml(raw).encode("utf-8"))


def text(elem):
    return N.clean_text("".join(elem.itertext()).replace(N.BR, " "))


def lines(elem):
    return [N.clean_text(s) for s in "".join(elem.itertext()).split(N.BR) if N.clean_text(s)]


def compact_title(node):
    return N.compact(text(node.elem))


def stream(root):
    """TITLE·P·TABLE 을 문서 순서로. 표 안의 P 는 표에 속하므로 제외한다."""
    out, cur = [], ""
    def walk(e):
        nonlocal cur
        for c in e:
            tag = c.tag
            if tag == "TITLE":
                cur = text(c)
                out.append(Node("title", c, len(out), cur))
            elif tag == "TABLE":
                out.append(Node("table", c, len(out), cur))
                continue                      # 표 내부는 내려가지 않음
            elif tag == "P":
                out.append(Node("p", c, len(out), cur))
                continue
            walk(c)
    walk(root)
    return out


def _span(cell, name):
    v = cell.get(name) or cell.get(name.lower())
    try:
        return max(1, int(v)) if v else 1
    except ValueError:
        return 1


def _own_rows(table):
    """표 자신의 행만: TABLE 직계 TR + THEAD/TBODY/TFOOT 직계 TR. 셀 안의 중첩 표 행은 제외."""
    out = []
    for c in table:
        if c.tag == "TR":
            out.append(c)
        elif c.tag in ("THEAD", "TBODY", "TFOOT"):
            out.extend(x for x in c if x.tag == "TR")
    return out


def table_grid(table):
    """rowspan/colspan 을 펼쳐 직사각 격자로. 병합된 칸은 같은 텍스트를 반복한다.
    헤더 = TH 로만 된 선두 행들(THEAD 포함). TE/TU 는 TD 로 본다."""
    trs = _own_rows(table)
    occupied = {}           # (r, c) -> text
    is_head = {}            # r -> bool
    ncols = 0
    for r, tr in enumerate(trs):
        c = 0
        cells = [x for x in tr if x.tag in _CELL_TAGS]
        head = bool(cells) and all(x.tag == "TH" for x in cells)
        is_head[r] = head
        for cell in cells:
            while (r, c) in occupied:
                c += 1
            rs, cs = _span(cell, "ROWSPAN"), _span(cell, "COLSPAN")
            t = text(cell)
            for dr in range(rs):
                for dc in range(cs):
                    occupied[(r + dr, c + dc)] = t
            c += cs
        ncols = max(ncols, c)
    rows = [[occupied.get((r, c), "") for c in range(ncols)] for r in range(len(trs))]
    header, body, in_head = [], [], True
    for r, row in enumerate(rows):
        if in_head and is_head.get(r):
            header.append(row)
        else:
            in_head = False
            body.append(row)
    return Grid(header, body, ncols, table.get("BORDER", "0"), dict(table.attrib))


_FN_NUM = re.compile(r"주\s*\d+\s*\)")


def is_caption_table(grid):
    """'(단위: 천원)'·'(당기) (단위 : 천원)'·'(주1) 설명' 같은 짧은 테두리 없는 표."""
    if grid.header or len(grid.body) > 3:
        return False
    t = grid.text()
    if not t or len(t) > 120:
        return False
    cells = {c for r in grid.rows() for c in r if c}
    if not (t.startswith("(") or len(cells) == 1):
        return False
    return ("(단위" in t or N.period_role(t.split("(단위")[0]) is not None
            or t.startswith("(주") or t.startswith("(*") or bool(_FN_NUM.match(t)))


def section(nodes, start_pred, end_pred):
    """start_pred 가 참인 첫 노드부터, 그 뒤 end_pred 가 참인 첫 노드 직전까지."""
    out, on = [], False
    for n in nodes:
        if not on:
            if start_pred(n):
                on = True
                out.append(n)
            continue
        if end_pred(n):
            break
        out.append(n)
    return out
