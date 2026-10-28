"""주석 섹션 → 목차(NoteEntry)와 표 레코드(NoteRecord)."""
import re
from dataclasses import dataclass, replace
from core import normalize as N
from core import xml_doc as X
from core.model import NoteEntry, NoteRecord

_DATE_HDR = re.compile(r"(20\d{2})\s*[.년]\s*(\d{1,2})\s*[.월]")


@dataclass
class NoteTable:
    idx: int
    caption: str
    grid: object
    node_idx: int
    footnote: str = ""


def note_section(nodes):
    return X.section(nodes,
                     lambda n: n.kind == "title" and N.compact(n.title).startswith("주석"),
                     lambda n: n.kind == "title")


def note_toc(nodes):
    sec = note_section(nodes)
    if not sec:
        return []
    entries, last_no = [], 0
    for n in sec:
        if n.kind != "p":
            continue
        ls = X.lines(n.elem)
        if not ls:
            continue
        sh = N.split_heading(ls[0])
        if not sh:
            continue
        no, title, _ = sh
        if not (last_no < int(no) <= last_no + 2):   # 본문 속 "7. 어쩌구" 오탐 방지: 증가하되 건너뜀은 1개까지
            continue
        last_no = int(no)
        entries.append(NoteEntry(no, title, N.norm_label(title), None, False, n.idx, sec[-1].idx + 1))
    for i, e in enumerate(entries):
        if i + 1 < len(entries):
            e.end = entries[i + 1].start
        e.table_count = sum(1 for n in sec if n.kind == "table" and e.start < n.idx < e.end
                            and not X.is_caption_table(X.table_grid(n.elem)))
    return entries


# 각주 표식으로 시작: (주1)·(*)·(*1)·주1). 캡션표와 본문 P 둘 다 이 모양이면 앞 표의 각주다.
_FOOTNOTE = re.compile(r"^\s*(?:\(\s*주\s*\d+\s*\)|\(\*+\s*\d*\s*\)|주\s*\d+\s*\))")


_SENT_END = re.compile(r"다\.")


def _attach_footnote(out, txt):
    if out:
        out[-1].footnote = N.clean_text((out[-1].footnote + " " + txt).strip())


def _split_footnote_p(t):
    """각주로 시작하는 P 를 (각주 문장들, 나머지) 로 가른다. 원문은 각주와 다음 소제목이 한 P 에 붙기도 한다:
    '(*) 손상차손누계액을 합산한 금액입니다.2) 전기' → (['(*) … 입니다.'], '2) 전기'). 첫 '다.' 까지가 각주 한 문장이고,
    나머지가 또 각주 표식으로 시작하면 되풀이한다. '다.' 가 없으면 P 전체가 각주."""
    notes = []
    while t and _FOOTNOTE.match(t):
        m = _SENT_END.search(t)
        if not m:
            notes.append(t)
            return notes, ""
        notes.append(t[:m.end()].strip())
        t = t[m.end():].strip()
    return notes, t


def note_tables(nodes, entry):
    """entry 범위 안의 본문표. 캡션 = 직전 캡션표들 + 직전 소제목 P. 각주 모양('(*) …'·'(주1) …')의 P·캡션표는 앞 표의
    각주로 붙이고 캡션에는 섞지 않는다(앞 표의 각주 문단이 다음 표 캡션 앞에 결합되던 결함)."""
    out, pending_caption, sub_heading = [], [], ""
    for n in nodes[entry.start + 1:entry.end]:
        if n.kind == "p":
            t = X.text(n.elem)
            if not t:
                continue
            notes, t = _split_footnote_p(t)
            for fn in notes:
                _attach_footnote(out, fn)
            if not t:
                continue
            sub_heading = t
            pending_caption = []
        elif n.kind == "table":
            g = X.table_grid(n.elem)
            if X.is_caption_table(g):
                txt = g.text()
                if _FOOTNOTE.match(txt):     # 각주표: 앞 표에 붙이고 다음 표 캡션에는 섞지 않음
                    _attach_footnote(out, txt)
                else:
                    pending_caption.append(txt)
                continue
            caption = N.clean_text(" ".join([sub_heading] + pending_caption))
            out.append(NoteTable(len(out), caption, g, n.idx))
            pending_caption = []
    return out


_DASHES = {"-", "－", "—", "―", "–"}


def _col_kind(grid, c):
    """열 성격: 대시류('-'·'－'·'—'·'―'·빈칸)는 분모에서 뺀다. 남은 값이 없으면(대시뿐) num, 남은 값의 숫자 비율 ≥ 0.6 이면 num.
    (값이 대부분 '-' 인 숫자 열 — 유형자산 변동표의 '토지' — 이 라벨 열로 붙어 행 라벨에 값이 섞이던 결함.)"""
    vals = [r[c] for r in grid.body if c < len(r) and r[c]]
    if not vals:
        return "empty"
    rest = [v for v in vals if N.clean_text(v) and N.clean_text(v) not in _DASHES]
    if not rest:
        return "num"
    numeric = sum(1 for v in rest if N.parse_amount(v) is not None)
    return "num" if numeric >= len(rest) * 0.6 else "text"


_FOOTMARK = re.compile(r"\(\s*\*+\s*\d*\s*\)|\(\s*주\s*\d+[^)]*\)")


def _period_of(text, periods):
    text = _FOOTMARK.sub("", text)          # '당기(*)'·'전기(주1)': 각주 표식을 떼고 판정
    role = N.period_role(text)
    if role is None:
        fn = N.fiscal_no(text)
        if fn is not None:
            role = next((k for k, p in periods.items() if p.fiscal_no == fn), None)
    if role is None:
        m = _DATE_HDR.search(text)
        if m:
            y = int(m.group(1))
            role = next((k for k, p in periods.items() if p.end.year == y), None)
    return periods.get(role) if role else None


_INFO_HDR = re.compile(r"이자율|연이율|이율|만기|주식수|지분율|비율|%|구분|비고|제공처|담보")


def _keep_outside(g, c, label):
    """기간 밖 열은 텍스트 열이거나 정보성 헤더(이자율·만기 등)일 때만 기록. 그 밖의 숫자 열은 기간을 알 수 없어 제외."""
    return _col_kind(g, c) == "text" or bool(_INFO_HDR.search(label or ""))


_PAREN = re.compile(r"\(([^()]{1,12})\)")
_UNIT_PAREN = re.compile(r"\(\s*단위[^)]*\)")
# 공백 앞이 접속사·'중'이면('당기와 전기'·'… 중 당기') 단독 기간 표기가 아니다
_CAP_TAIL = re.compile(r"(?:^|[\d①-⑳]\s*\)|\(\s*\d+\s*\)|[가-힣]\.|[:：]|(?<!와)(?<!과)(?<!및)(?<!중)(?<!또는)\s)\s*"
                       r"(당기|전기|전전기|당반기|전반기|당기말|전기말)\s*$")
_TWO_PERIODS = re.compile(r"(?:당|전|전전)(?:반)?기\s*(?:와|과|및|또는)\s*(?:당|전|전전)(?:반)?기")


def _caption_period(caption, periods):
    """캡션 속 괄호 표식 '(당기)'·'(전기말)' 에서 기간을 읽는다. 본문 문장 속 '당기 및 전기'는 무시."""
    for m in _PAREN.finditer(caption):
        p = _period_of(m.group(0), periods)
        if p:
            return p
    # '① 제 25(당) 기' 처럼 회차 표식만 있는 소제목: 캡션 속 회차가 한 기간만 가리킬 때
    found = {id(p): p for p in (_period_of(m.group(0), periods) for m in N._FISCAL.finditer(caption)) if p}
    if len(found) == 1:
        return next(iter(found.values()))
    # 캡션 꼬리 규칙: '…다음과 같습니다.1) 당기 (단위: 천원)'·'① 당기' 처럼 끝이 번호 뒤 기간 표기
    t = re.sub(r"\s+", " ", _FOOTMARK.sub("", _UNIT_PAREN.sub("", caption))).strip()
    # '당기와 전기 …' 두 기간을 함께 말하는 캡션은 제외. 단 마침표 뒤 마지막 절만 본다 —
    # '(1) 당기와 전기 중 … 같습니다.1) 당기' 는 앞 문장이 아니라 꼬리 '1) 당기' 가 표의 기간이다.
    if _TWO_PERIODS.search(re.split(r"[.。]", t)[-1]):
        return None
    m = _CAP_TAIL.search(t)
    return _period_of(m.group(1), periods) if m else None


def _header_periods(grid, periods):
    """기간 표식이 있는 헤더 행의 기간 그룹: ([(cols, Period)], 행 번호)."""
    for r, hdr in enumerate(grid.header):
        groups = _row_groups(hdr, periods)
        if groups:
            return groups, r
    return [], None


def _row_groups(hdr, periods):
    groups, c = [], 0
    while c < len(hdr):
        e = c
        while e + 1 < len(hdr) and hdr[e + 1] == hdr[c]:
            e += 1
        p = _period_of(hdr[c], periods) if hdr[c] else None
        if p:
            groups.append((list(range(c, e + 1)), p))
        c = e + 1
    return groups


def _cell_value(text):
    v = N.parse_amount(text)
    if v is not None:
        return v
    v = N.parse_percent(text)
    if v is not None:
        return v
    t = N.clean_text(text)
    return t if t and t not in ("-", "－") else None


def is_subheading(row, label_cols):
    """본문 행이 소제목인가: 라벨 열 밖의 값 칸이 모두 같은 글자(숫자·비율 아님)이고 그 글자가 라벨과 같다.
    예: '(1) 판매비와관리비:' 가 병합 셀이라 모든 열에 같은 글자로 들어온 행. 빈 칸은 무시."""
    labs = [N.clean_text(row[c]) for c in label_cols if c < len(row) and row[c] and N.clean_text(row[c])]
    if not labs:
        return False
    vals = {N.clean_text(row[c]) for c in range(len(row)) if c not in label_cols and row[c] and N.clean_text(row[c])}
    if len(vals) != 1:
        return False
    v = next(iter(vals))
    if N.parse_amount(v) is not None or N.parse_percent(v) is not None:
        return False
    return N.norm_label(v) == N.norm_label(" / ".join(dict.fromkeys(labs))) or N.norm_label(v) == N.norm_label(labs[0])


def _strip_unit_rows(g):
    """헤더 맨 위의 '(단위: 천원)' 만으로 된 행(전 열 병합)을 떼어 낸다. 열 라벨로 오인하지 않도록. (그리드, 단위)"""
    unit, k = None, 0
    while k < len(g.header) - 1:
        cells = {N.clean_text(c) for c in g.header[k] if c and N.clean_text(c)}
        if len(cells) != 1:
            break
        u = N.parse_unit(next(iter(cells)))
        if u is None or not N.compact(next(iter(cells))).lstrip("(").startswith("단위"):
            break
        unit, k = u, k + 1
    return (replace(g, header=g.header[k:]) if k else g), unit


def table_records(nt, entry, periods, fs_div, rcept_no):
    g, hdr_unit = _strip_unit_rows(nt.grid)
    if "cur" not in periods or not g.header or g.ncols < 2:
        return [], "other"
    unit = N.parse_unit(nt.caption) or hdr_unit or N.parse_unit(g.text())
    col_unit_ok = unit is None          # 표 단위가 없을 때만 열 라벨의 '(천원)' 을 그 열의 단위로(라벨 표시는 원문 그대로)
    unit = unit or 1
    groups, prow = _header_periods(g, periods)
    in_group = {c for cols, _ in groups for c in cols}
    label_cols = [0]                       # 라벨 열 = 0열 + (바로 옆 텍스트 열 1개)
    if g.ncols > 2 and 1 not in in_group and _col_kind(g, 1) == "text":
        label_cols.append(1)

    def col_label_of(c):
        """기간 행을 뺀 나머지 헤더 행들의 텍스트를 합친 열 라벨(같은 글자는 한 번만)."""
        parts = [N.clean_text(h[c]) for r, h in enumerate(g.header) if r != prow and c < len(h) and h[c]]
        return " / ".join(dict.fromkeys(parts)) or None
    recs = []

    def label_of(row):
        parts = [N.clean_text(row[c]) for c in label_cols if c < len(row) and row[c]]
        parts = list(dict.fromkeys(parts))
        return " / ".join(parts)

    def add(period, row, c, col_label):
        label = label_of(row)
        if not label or c >= len(row):
            return
        value = None if is_subheading(row, label_cols) else _cell_value(row[c])
        col_unit = N.split_col_unit(col_label)[1] if col_unit_ok else None
        recs.append(NoteRecord(period.key, fs_div, entry.no, entry.title, entry.topic, nt.idx, nt.caption,
                               label, N.norm_label(label), col_label, value, col_unit or unit, rcept_no,
                               "cur" if period is periods.get("cur") else "prev"))

    if groups:
        single = all(len(cols) == 1 for cols, _ in groups)
        if single and len(g.header) == 1:
            for row in g.body:
                for cols, p in groups:
                    add(p, row, cols[0], None)
            # 기간 열이 아닌 숫자/텍스트 열(이자율·만기 등)은 열 라벨로
            for c in range(g.ncols):
                if c in label_cols or any(c in cols for cols, _ in groups):
                    continue
                lab = N.clean_text(g.header[0][c])
                if not _keep_outside(g, c, lab):
                    continue
                for row in g.body:
                    add(periods["cur"], row, c, lab)
            return recs, "two_col"
        for cols, p in groups:
            for c in cols:
                col_label = col_label_of(c)
                if col_label is None and len(g.header) == 1:
                    col_label = N.clean_text(g.header[0][c]) or None
                for row in g.body:
                    add(p, row, c, col_label)
        for c in range(g.ncols):           # 기간 밖의 공통 열(담보제공처·비고)
            if c in label_cols or any(c in cols for cols, _ in groups):
                continue
            lab = col_label_of(c) or N.clean_text(g.header[0][c])
            if not _keep_outside(g, c, lab):
                continue
            for row in g.body:
                add(periods["cur"], row, c, lab)
        return recs, "grouped"

    p = _caption_period(nt.caption, periods)
    if p is None:
        return [], "other"
    hdr = g.header[-1]
    for row in g.body:
        for c in range(g.ncols):
            if c in label_cols:
                continue
            add(p, row, c, N.clean_text(hdr[c]) if c < len(hdr) else f"열{c+1}")
    return recs, "rollforward"
