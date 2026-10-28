"""회사 제시 엑셀 파일 → Collected (파일 입력 2단계). DART 경로의 collect() 와 같은 모양을 돌려주어
layout·compose·checks·finalize·refresh 가 그대로 쓰인다.

- 파일마다 xlsx_input 으로 표를 인식하고 cfg.layout 의 사용자 수정값(유형·머리글 행·라벨 열·기간 열↔기간·단위)을 덮어쓴다.
  layout 키 = "<파일명>|<시트>|<표번호(0부터)>". 같은 파일명이 둘이면 rcept_no 와 같은 꼬리표("a.xlsx#2")로 구분한다.
- 연간 키(FY)의 출처는 연간 파일, 반기 키(H1)의 출처는 반기 파일뿐(DART 와 같음). 반기 파일의 주석 표는 쓰지 않는다.
- 기간 열은 period_cols 만 쓴다(값 열 전체가 아니다). 열마다 (1) layout 지정 → (2) 당기/전기/전전기 역할 → (3) 머리글 날짜
  → (4) 회차만 있으면 가장 큰 회차를 당기로(로그 + plan.inferred 표시) → (5) 못 정하면 제외.
- 원문에 없는 값·합계를 만들지 않는다. OTHER 표·표 밖 숫자 행은 로그만 남긴다. 금액은 단위 환산 없이 원문 값 + unit.
- 재작성 기준: 결산일이 늦은 파일의 전기 열 우선(pipeline._pick_restated 재사용).
- 캐시·로그 파일·네트워크·COM 을 쓰지 않는다. 셀 내용은 col.logs(메모리) 에만 남는다.
"""
import os
import re
from dataclasses import dataclass, field, replace
from datetime import date, datetime

from core import normalize as N
from core import note_parser as NP
from core import pipeline as P
from core import xlsx_input as XI
from core.model import FsRecord, NoteEntry, NoteRecord, Period, ReportRef
from core.note_std import detect_topic
from core.xlsx_input import OFF_ROLE as _OFF_ROLE, ROLE_OFF as _ROLE_OFF

_PK = re.compile(r"^(?:FY(\d{4})|(\d{4})H1)$")
_ROLE_KO = {"cur": "당기", "prev": "전기", "prev2": "전전기"}
_YEAR_IN_LABEL = re.compile(r"(20\d{2})\s*[.년]")          # fs_parser._YEAR 와 같은 규칙(자본변동표 행 라벨의 연도)
_SCE_PREFIX = re.compile(r"^(전전기|전기|당기)(?:기초|기말|초|말)")  # '전기 기초잔액'·'당기말' 처럼 경계어가 붙을 때만 구간 전환
_SCE_OFF = {"당기": 0, "전기": 1, "전전기": 2}
_FS_KINDS = ("BS", "IS", "CIS", "CF")
NO_TABLE_NOTE = "인식된 표 없음(수식만 있는 파일이면 Excel 에서 한 번 저장한 뒤 다시 넣으세요)"
_UNIT_NAME = {1: "원", 10**3: "천원", 10**4: "만원", 10**5: "십만원", 10**6: "백만원", 10**7: "천만원",
              10**8: "억원", 10**9: "십억원", 10**10: "백억원"}
_NOTE_NO = re.compile(r"^주석\s*(\d+)\s*[.:：\-]?\s*(.*)$")
_CUR_WORDS = ("당기", "당반기", "당기말", "당반기말")
_PREV_WORDS = ("전기", "전반기", "전기말", "전전기")


@dataclass
class FilePlan:
    """selector.Plan 과 같은 속성 + inferred(기간 역할을 회차로 추정한 표 키 — UI 에서 확인 표시용)."""
    periods: list = field(default_factory=list)
    reports: dict = field(default_factory=dict)
    sources: dict = field(default_factory=dict)
    note_sources: dict = field(default_factory=dict)
    base_report: object = None
    inferred: list = field(default_factory=list)
    # rcept_no -> {"sheets": ["재무상태표(표 1)", "주석(표 1·2)"], "unit": "원·천원"} — 로그 시트 '보고서' 구역용(쓴 표만)
    file_meta: dict = field(default_factory=dict)


@dataclass
class _File:
    path: str
    basename: str
    rcept_no: str
    period_end: date
    kind: str                # "annual" | "half"
    fs_div: str
    unit: object             # int | None(파일 단위 미지정)
    mtime: date

    @property
    def cur(self):
        return _period(self.kind, self.period_end.year)

    @property
    def key_base(self):
        """layout 키의 파일 부분 = rcept_no 에서 'FILE:' 을 뺀 것(같은 파일명이 둘이면 'a.xlsx#2' 로 구분)."""
        return self.rcept_no[5:]

    def ref(self):
        return ReportRef(self.rcept_no, self.basename, self.mtime, self.kind, "파일", self.period_end, (self.fs_div,))


# ---------------------------------------------------------------- 기간

HALF_FY_NOTE = "(반기 파일 — FY 출처 아님)"     # 반기 파일 BS 전기말 열: 기록만, FY 열·출처·회차로 쓰지 않음


def _kind_of_key(pk):
    """기간 키 → 'annual'(FY…) | 'half'(…H1)."""
    return "annual" if str(pk).startswith("FY") else "half"


def _period(kind, year, fiscal=None):
    return Period.annual(year, fiscal) if kind == "annual" else Period.half(year, fiscal)


def _period_of_key(key):
    m = _PK.match(str(key or "").strip())
    if not m:
        return None
    return Period.annual(int(m.group(1))) if m.group(1) else Period.half(int(m.group(2)))


def _period_from_end(end, kind):
    """머리글 날짜 'YYYY-MM-DD'·'YYYY-MM'·'YYYY' → Period. 6월은 반기, 12월은 연간, 연도만이면 파일 종류를 따른다."""
    y = int(end[:4])
    mo = int(end[5:7]) if len(end) >= 7 else None
    if mo == 6:
        return Period.half(y)
    if mo == 12 or mo is None:
        return _period(kind, y) if mo is None else Period.annual(y)
    return None


def _period_for_role(f, role, sj):
    """파일 당기에서 역할만큼 떨어진 기간. 반기 파일의 재무상태표 전기 열은 직전 연차 기말(DART prev_bs 규칙)."""
    off = _ROLE_OFF[role]
    if f.kind == "half" and sj == "BS" and off:
        return Period.annual(f.period_end.year - off)
    return _period(f.kind, f.period_end.year - off)


def _role_of(f, p):
    """레코드 역할: 파일 당기면 cur, 그보다 n년 전이면 prev/prev2(그 밖은 prev<n>)."""
    off = f.period_end.year - p.end.year
    if off == 0 and p.kind == f.kind:
        return "cur"
    return _OFF_ROLE.get(off, f"prev{off}")


def _resolve_columns(t, f, sj, override, key, flags):
    """표의 기간 열 → ({열: (Period, role)}, 메시지들). 순서: layout → 역할 → 날짜 → 회차만(추정 표시) → 제외."""
    out, msgs, pending = {}, [], []
    ov = {str(k): v for k, v in ((override or {}).get("periods") or {}).items()}
    cols = dict(t.period_cols)
    for c in ov:
        if c.isdigit() and int(c) not in cols:
            cols[int(c)] = XI.PeriodGuess("")
    for c, g in sorted(cols.items()):
        if str(c) in ov:
            v = ov[str(c)]
            if v is None:
                msgs.append(f"열 {c}: 사용자 설정으로 제외")
                continue
            p = _period_of_key(v)
            if p is None:
                msgs.append(f"열 {c}: 기간 '{v}' 을(를) 해석하지 못해 제외")
                continue
            out[c] = (p, _role_of(f, p))
        elif g.role in _ROLE_OFF:
            out[c] = (_period_for_role(f, g.role, sj), g.role)
        elif g.end_date and _period_from_end(g.end_date, f.kind) is not None:
            p = _period_from_end(g.end_date, f.kind)
            out[c] = (p, _role_of(f, p))
        elif g.fiscal is not None:                # 날짜가 6월·12월 말이 아니어도 회차가 있으면 회차로
            pending.append(c)
        elif g.end_date:
            msgs.append(f"열 {c} '{g.text}': 6월·12월 말이 아닌 날짜, 회차도 없음 → 제외")
        else:
            msgs.append(f"열 {c} '{g.text}': 기간을 정하지 못해 제외")
    if pending:                                   # 회차만 있는 열: 역할을 추정하므로 로그 + 화면 확인 표시
        anchors = [(cols[c].fiscal, _ROLE_OFF.get(out[c][1])) for c in out
                   if cols[c].fiscal is not None and out[c][1] in _ROLE_OFF]
        if anchors:                               # 날짜·layout 으로 정해진 열의 회차에서 차이로
            af, aoff = anchors[0]
            base = lambda c: aoff + (af - cols[c].fiscal)
        else:                                     # 회차뿐: 가장 큰 회차 = 당기
            top = max(cols[c].fiscal for c in pending)
            base = lambda c: top - cols[c].fiscal
        for c in pending:
            role = _OFF_ROLE.get(base(c))
            if role is None:
                msgs.append(f"열 {c} '{cols[c].text}': 당기와 3기 이상 차이 → 제외")
                continue
            out[c] = (_period_for_role(f, role, sj), role)
        msgs.append("기간 역할을 회차로 추정 — 화면에서 확인")
        flags.append(key)
    # 같은 기간을 여러 열이 가지면 하나만: 사용자가 매긴 열 → 없으면 앞 열. 나머지는 제외(레코드 중복 방지)
    user = {int(c) for c, v in ov.items() if c.isdigit() and v is not None}
    keep = {}
    for c in sorted(out, key=lambda c: (c not in user, c)):
        keep.setdefault(out[c][0].key, c)
    for c in sorted(out):
        k = out[c][0].key
        if keep[k] != c:
            msgs.append(f"열 {c} 은 기간 {k} 와 겹쳐 제외")
            del out[c]
    return out, msgs


# ---------------------------------------------------------------- 셀·행

def _label(row, label_cols):
    parts = [XI._text(XI._cell(row, c - 1)) for c in label_cols]
    return " / ".join(dict.fromkeys(p for p in parts if p))


def _amount(v):
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    return N.parse_amount(v) if isinstance(v, str) else None


def _note_value(v):
    """주석 셀 값: 숫자 그대로, '(1,234)'·'12.5%' 는 수치로, 글자는 글자로, 대시·빈칸은 None."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    if not isinstance(v, str):
        return XI._text(v)
    a = N.parse_amount(v)
    if a is not None:
        return a
    pc = N.parse_percent(v)
    if pc is not None:
        return pc
    t = N.clean_text(v)
    return t if t and not XI._is_dash(t) else None


def _apply_layout(t, ov, rows=None, sheet_name="", mapping=None):
    """cfg.layout 의 사용자 수정값을 TableScan 복사본에 덮어쓴다(기간 열↔기간은 _resolve_columns 가 본다).
    머리글 행·라벨 열을 고치면(rows 가 있을 때) 그 행·열로 표를 다시 인식해 값 열·기간 열·본문 범위를 맞춘다."""
    if not ov:
        return t
    ch = {}
    hr = int(ov["header_row"]) if ov.get("header_row") else None
    lc = [int(c) for c in ov["label_cols"]] if ov.get("label_cols") else None
    rebuilt = False
    if rows is not None and (hr or lc):
        new = XI.rebuild(rows, t, header_row=hr, label_cols=lc, sheet_name=sheet_name, mapping=mapping)
        if new is not None:              # 다시 인식한 머리글 행(지정 행부터 2행일 수 있음)·라벨 열·본문 범위를 그대로 쓴다
            t, rebuilt = new, True
    if ov.get("kind"):
        ch["kind"] = str(ov["kind"]).upper()
    if hr and not rebuilt:
        ch["header_row"] = hr
        ch["first_row"] = int(ov.get("first_row") or hr + 1)
    elif ov.get("first_row"):
        ch["first_row"] = int(ov["first_row"])
    if ov.get("last_row"):
        ch["last_row"] = int(ov["last_row"])
    if lc and not rebuilt:
        ch["label_cols"] = lc
    if ov.get("unit"):
        ch["unit"], ch["unit_text"] = int(ov["unit"]), "사용자 설정"
    return replace(t, **ch)


def _unit_text(u, known):
    return _UNIT_NAME.get(u, str(u)) if known else "미표기"


# ---------------------------------------------------------------- 레코드

def _fs_records(t, rows, f, cols, unit):
    out, ord_ = [], 0
    for r in range(t.first_row - 1, t.last_row):
        row = rows[r]
        label = _label(row, t.label_cols)
        if not label:
            continue
        ord_ += 1
        for c, (p, role) in cols.items():
            out.append(FsRecord(p.key, f.fs_div, t.kind, ord_, label, N.norm_label(label), _amount(XI._cell(row, c - 1)),
                                unit, None, None, f.rcept_no, "XLSX", role))
    return out


def _comp_cols(t, rows):
    """구성요소 열(자본변동표·변동표): 라벨 열을 뺀 값 열 → {열: 머리글 글자}. '주석' 열은 제외."""
    hdr = rows[t.header_row - 1]
    out = {}
    for c in t.value_cols:
        if c in t.label_cols:
            continue
        h = XI._text(XI._cell(hdr, c - 1))
        if "주석" in N.compact(h):
            continue
        out[c] = h or f"열{c}"
    return out


def _sce_records(t, rows, f, cols, unit, msgs):
    """자본변동표: 행 라벨의 연도(2025.1.1(당기초)…)로 기간을 잡고 열 = 자본 구성요소(fs_parser._parse_sce 와 같은 모양).
    기간 열이 있는(가로형) 자본변동표는 일반 재무제표처럼 행마다 기간 열 값을 쓴다."""
    if cols:
        return _fs_records(t, rows, f, cols, unit)
    comps = {c: ("총계" if N.compact(h) == "총계" else h) for c, h in _comp_cols(t, rows).items()}
    by_off = {off: (_period(f.kind, f.period_end.year - off), _OFF_ROLE[off]) for off in (2, 1, 0)}
    by_year = {p.end.year: (p, r) for p, r in by_off.values()}
    body = [(rows[r], _label(rows[r], t.label_cols)) for r in range(t.first_row - 1, t.last_row)]
    labels = [N.compact(lab) for _, lab in body if lab]

    def by_year_of(label):
        m = _YEAR_IN_LABEL.search(label)
        return by_year.get(int(m.group(1))) if m else None

    def by_prefix_of(label):
        m = _SCE_PREFIX.match(N.compact(label))
        return by_off[_SCE_OFF[m.group(1)]] if m else None

    # 행의 기간: (1) 라벨의 연도 → (2) 연도 없으면 '전기초'·'당기 기말잔액' 같은 경계 행(접두어 + 기초·기말·초·말)
    #            → (3) 아무 표기도 없으면 전부 당기(로그). 한 표에 '전기초'와 '당기 기초잔액'이 섞여도 같은 규칙
    if any(by_year_of(lab) for lab in labels):
        marker, cur = by_year_of, None
    elif any(_SCE_PREFIX.match(lab) for lab in labels):
        marker, cur = by_prefix_of, None
        msgs.append("행 라벨에 연도 없음 → '전기…/당기…' 접두어로 기간을 나눔")
    else:
        marker, cur = (lambda label: None), (f.cur, "cur")
        msgs.append("행 라벨에 연도·당기/전기 표기 없음 → 전부 당기로")
    out, ord_ = [], 0
    for row, label in body:
        if not label:
            continue
        cur = marker(label) or cur
        if cur is None:
            continue
        ord_ += 1
        p, role = cur
        for c, cl in comps.items():
            out.append(FsRecord(p.key, f.fs_div, "SCE", ord_, label, N.sce_row_norm(label), _amount(XI._cell(row, c - 1)),
                                unit, cl, None, f.rcept_no, "XLSX", role))
    return out


def _note_id(title, heading, sheet_name):
    """(주석 번호, 제목): 표 제목 또는 위쪽 주석 제목이 'N. 주제'·'주석 N 주제' 꼴이면 거기서, 아니면 ('', 제목 또는 시트명)."""
    for t in (title, heading):
        if not t:
            continue
        sh = N.split_heading(t)
        if sh:
            return sh[0], sh[1]
        m = _NOTE_NO.match(N.clean_text(t))
        if m:
            return m.group(1), N.clean_text(m.group(2)) or sheet_name
    return "", N.clean_text(title or "") or sheet_name


def _caption_period(caption, f):
    """기간 열 없는 주석 표의 기간: note_parser 캡션 규칙 → 없으면 캡션에 당기(또는 전기)만 한쪽으로 적혀 있을 때."""
    per = {"cur": f.cur, "prev": _period(f.kind, f.period_end.year - 1), "prev2": _period(f.kind, f.period_end.year - 2)}
    p = NP._caption_period(caption, per)
    if p is not None:
        return p, ""
    c = N.compact(caption)
    has_cur, has_prev = any(w in c for w in _CUR_WORDS), any(w in c for w in _PREV_WORDS)
    if has_cur and not has_prev:
        return per["cur"], f"표 기간을 캡션 '{N.clean_text(caption)[:30]}' 의 당기 표기로"
    if has_prev and not has_cur:
        return per["prev"], f"표 기간을 캡션 '{N.clean_text(caption)[:30]}' 의 전기 표기로"
    return None, ""


def _note_records(t, rows, f, cols, unit, unit_known, mapping, sheet_name, caption_texts, heading, msgs):
    """주석 표 → NoteRecord. 기간 열이 있으면 열마다, 없으면 캡션 기간에 구성요소 열(열 라벨)로,
    그것도 없으면 원표 그대로(raw_cell → rN_기타). → (레코드, (no, title, topic, recommended))."""
    no, ntitle = _note_id(t.title, heading, sheet_name)
    topic, rec = detect_topic(N.norm_label(ntitle), mapping)
    caption = " ".join(caption_texts) or N.clean_text(t.title or "") or sheet_name
    body = [(rows[r], _label(rows[r], t.label_cols)) for r in range(t.first_row - 1, t.last_row)]

    def rec_(p, role, label, cl, v, cell=None):
        return NoteRecord(p.key, f.fs_div, no, ntitle, topic, t.idx, caption, label, N.norm_label(label), cl, v, unit,
                          f.rcept_no, role, cell, unit_known)
    out = []
    if cols:
        for row, label in body:
            if label:
                for c, (p, role) in cols.items():
                    out.append(rec_(p, role, label, None, _note_value(XI._cell(row, c - 1))))
        return out, (no, ntitle, topic, rec)
    comps = _comp_cols(t, rows)
    p, why = _caption_period(caption, f)
    if p is not None:
        if why:
            msgs.append(why)
        role = _role_of(f, p)
        for row, label in body:
            if label:
                for c, cl in comps.items():
                    out.append(rec_(p, role, label, cl, _note_value(XI._cell(row, c - 1))))
        return out, (no, ntitle, topic, rec)
    msgs.append("기간 표시 없음 → rN_기타")
    hdr = rows[t.header_row - 1]
    all_cols = list(t.label_cols) + [c for c in t.value_cols if c not in t.label_cols]
    for ri, (row, label) in enumerate(body):
        for ci, c in enumerate(all_cols):
            v = XI._cell(row, c - 1)
            if c in t.label_cols:
                v = XI._text(v) or None
            else:
                v = _note_value(v)
            cl = XI._text(XI._cell(hdr, c - 1)) or f"열{ci + 1}"
            out.append(rec_(f.cur, "cur", label or f"행{ri + 1}", cl, v, (ri, ci)))
    return out, (no, ntitle, topic, rec)


# ---------------------------------------------------------------- 파일·시트

def _files_of(cfg):
    out, seen = [], {}
    for d in cfg.files or []:
        path = str(d.get("path") or "")
        if not path or not os.path.isfile(path):
            raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {path}")
        pe = d.get("period_end")
        try:
            pe = date.fromisoformat(str(pe)[:10]) if not isinstance(pe, date) else pe
        except ValueError:
            pe = None
        if pe is None:
            raise ValueError(f"파일 {os.path.basename(path)}: 결산일이 없습니다")
        kind = d.get("kind") or "annual"
        if kind not in ("annual", "half"):
            raise ValueError(f"{os.path.basename(path)}: kind 는 annual 또는 half 이어야 합니다: {kind!r}")
        base = os.path.basename(path)
        n = seen[base] = seen.get(base, 0) + 1
        rc = f"FILE:{base}" + (f"#{n}" if n > 1 else "")
        unit = d.get("unit")
        unit = int(unit) if unit not in (None, "", 0, "0") else None
        mtime = datetime.fromtimestamp(os.path.getmtime(path)).date()
        out.append(_File(path, base, rc, pe, kind, d.get("fs_div") or "OFS", unit, mtime))
    return out


def _texts_above(rows, t, floor):
    """표 머리글 위(앞 표 끝 아래, 최대 4행)의 글자 셀 → (캡션 글자들, 주석 제목). 기간 행·단위 문구·제목은 캡션에서 뺀다."""
    caps, heading = [], None
    lo = max(floor, t.header_row - 1 - 4)
    for r in range(lo, t.header_row - 1):
        row = rows[r]
        if len(XI._period_cells(row)) >= 2:
            continue
        for v in row:
            if not XI._is_text(v):
                continue
            s = XI._text(v)
            if XI.unit_of(s):
                continue
            if XI._is_heading(s):
                heading = s
            else:
                caps.append(s)
    if heading is None:                      # 더 위쪽(앞 표들 밖)의 마지막 주석 제목
        for r in range(lo - 1, -1, -1):
            for v in rows[r]:
                if XI._is_text(v) and XI._is_heading(XI._text(v)):
                    heading = XI._text(v)
                    break
            if heading:
                break
    return caps, heading


def _sort_key(f):
    return (f.period_end, f.mtime, f.basename)


@dataclass
class _SheetTables:
    """파일의 시트 하나: 읽은 rows·인식 결과·layout 적용 표. items = [(layout 키, 원 표, 수정값 적용 표)]."""
    name: str
    hidden: bool
    rows: list
    note: str
    orphans: list
    items: list


def _iter_tables(path, key_base, layout, mapping):
    """파일의 시트마다 _SheetTables — preview 와 collect_files 가 같은 인식·layout 적용 규칙을 쓴다.
    숨김 시트도 돌려주므로(화면 표시용) 쓸지 말지는 부르는 쪽이 정한다."""
    layout = layout or {}
    for name, hidden, rows, note in XI.load_sheets(path):
        tables, orphans = XI.scan_sheet(rows, name, mapping)
        items = []
        for t0 in tables:
            key = f"{key_base}|{name}|{t0.idx}"
            items.append((key, t0, _apply_layout(t0, layout.get(key), rows, name, mapping)))
        yield _SheetTables(name, hidden, rows, note, orphans, items)


# ---------------------------------------------------------------- 화면 미리보기(F4 /api/file/scan)

def _guess_period_end(tables):
    """표들의 기간 열에서 파일 결산일 추정: 당기 열의 날짜 중 가장 늦은 것(없으면 모든 날짜 중 가장 늦은 것).
    연도만 있으면 12-31. → ("YYYY-MM-DD" | None, "annual" | "half")."""
    cur, any_ = [], []
    half_word = False
    for t in tables:
        for g in t.period_cols.values():
            if "반기" in N.compact(g.text or ""):
                half_word = True
            if g.end_date:
                (cur if g.role == "cur" else any_).append(g.end_date)
    pool = cur or any_
    if not pool:
        return None, "half" if half_word else "annual"
    end = max(pool, key=lambda s: (s[:4], s[5:7] or "12", s[8:10] or "31"))
    if len(end) == 4:
        end = f"{end}-12-31"
    elif len(end) == 7:
        end = f"{end}-{'30' if end[5:7] == '06' else '31'}"
    return end, ("half" if end[5:7] == "06" else "annual")


def preview(path, *, period_end=None, kind=None, fs_div="OFS", unit=None, key_base=None, layout=None, mapping=None):
    """파일 하나를 collect_files 와 같은 규칙으로 읽어 화면용 사전을 돌려준다(레코드는 만들지 않는다).
    period_end·kind 가 없으면 표에서 추정한다. layout 은 cfg.layout 과 같은 모양(사용자 수정값 반영).
    → {path, basename, key_base, scan(FileScan.to_dict + 표마다 key·auto_periods·period_msgs·inferred, 주석이면 note_no·
       note_title·title_norm), period_end_guess, kind_guess, unit_guess, inferred:[표 키], note(표가 하나도 없을 때)}.
    auto_periods 는 값 열 전부(기간을 못 매긴 열은 null) — 화면에서 다른 열로 바꿀 수 있게. 셀 내용은 담지 않는다."""
    base = os.path.basename(path)
    key_base = key_base or base
    layout = layout or {}
    scanned = list(_iter_tables(path, key_base, layout, mapping))
    all_tables = [t for s in scanned if not s.hidden for _, _, t in s.items]
    pe_guess, kind_guess = _guess_period_end(all_tables)
    units = [t.unit for t in all_tables if t.unit is not None and t.kind != "OTHER"] or [t.unit for t in all_tables if t.unit is not None]
    unit_guess = max(set(units), key=units.count) if units else None
    pe = str(period_end)[:10] if period_end else pe_guess
    kind = kind if kind in ("annual", "half") else kind_guess
    f = None
    if pe:
        try:
            f = _File(path, base, "FILE:" + key_base, date.fromisoformat(pe), kind, fs_div or "OFS",
                      int(unit) if unit not in (None, "", 0, "0") else None, date.today())
        except ValueError:
            f = None
    inferred = []
    out_sheets = []
    for s in scanned:
        floor = 0
        dicts = []
        for key, _, t in s.items:
            d = t.to_dict()
            d["key"] = key
            # 기간 열이 있는 표는 값 열 전부를 보여 준다(같은 기간의 3개월·누적 열 중 다른 열을 고를 수 있게); 구성요소 열 표는 기간 열만
            shown = set(t.value_cols) if t.period_cols else set(t.period_cols)
            if f is not None:
                sj = t.kind if t.kind in _FS_KINDS else None
                flags = []
                cols, msgs = _resolve_columns(t, f, sj, layout.get(key), key, flags)
                d["auto_periods"] = {str(c): (cols[c][0].key if c in cols else None) for c in sorted(shown | set(cols))}
                d["period_msgs"] = msgs
                d["inferred"] = bool(flags)
                if flags:
                    inferred.append(key)
            else:
                d["auto_periods"] = {str(c): None for c in sorted(shown)}
                d["period_msgs"] = ["결산일을 정하면 기간을 매깁니다"] if t.period_cols else []
                d["inferred"] = False
            if t.kind == "NOTE":
                _, heading = _texts_above(s.rows, t, floor)
                no, ntitle = _note_id(t.title, heading, s.name)
                d["note_no"], d["note_title"], d["title_norm"] = no, ntitle, N.norm_label(ntitle)
            dicts.append(d)
            floor = t.last_row
        out_sheets.append({"name": s.name, "hidden": s.hidden, "tables": dicts, "orphan_rows": list(s.orphans), "note": s.note})
    note = None if all_tables else NO_TABLE_NOTE
    return {"path": path, "basename": base, "key_base": key_base, "scan": {"path": path, "sheets": out_sheets},
            "period_end_guess": pe, "kind_guess": kind, "unit_guess": unit_guess, "inferred": inferred, "note": note}


def collect_files(cfg, mapping, log=None):
    """cfg.files 의 엑셀 파일들 → Collected. 파일이 없으면 FileNotFoundError('입력 파일을 찾을 수 없습니다: …')."""
    files = _files_of(cfg)
    plan = FilePlan()
    col = P.Collected(cfg, plan, [], mapping=mapping)
    layout = cfg.layout or {}
    by_report, note_by_report = {}, {}
    toc = {}                                  # (no, title_norm) -> [NoteEntry, 표 수]
    own = []                                  # 파일 자신의 결산 기간(항상 열이 된다)
    for f in sorted(files, key=_sort_key):    # 결산일 오래된 파일부터 → units 는 최신 파일 값이 남는다
        if f.kind == "annual" and f.period_end.month != 12 or f.kind == "half" and f.period_end.month != 6:
            P._log(col, log, f"파일 {f.basename}: 결산일 {f.period_end} — 12월 결산(반기는 6월)이 아님, 1차 범위 밖 → 건너뜀")
            continue
        if f.fs_div not in cfg.fs_divs:
            P._log(col, log, f"파일 {f.basename}: 연결/별도 구분 {f.fs_div} 이(가) 설정(fs_divs)에 없음 → 건너뜀")
            continue
        plan.reports[f.rcept_no] = f.ref()
        own.append(f.cur)
        used, units_used = {}, []             # 이 파일에서 쓴 시트별 표 번호(1부터)·단위 → plan.file_meta
        plan.file_meta[f.rcept_no] = {"sheets": [], "unit": ""}
        n_tables = 0
        for s in _iter_tables(f.path, f.key_base, layout, mapping):
            name, rows = s.name, s.rows
            if s.hidden:
                P._log(col, log, f"파일 {f.basename} 시트 {name}: 숨김 시트 → 건너뜀")
                continue
            if s.note:
                P._log(col, log, f"파일 {f.basename} 시트 {name}: {s.note}")
            if s.orphans:
                P._log(col, log, f"파일 {f.basename} 시트 {name}: 표에 속하지 않은 숫자 행 {len(s.orphans)}개"
                                 f"(행 {', '.join(str(r) for r in s.orphans[:20])})")
            floor = 0
            n_tables += len(s.items)
            for key, _, t in s.items:
                head = f"파일 {f.basename} 시트 {name} 표 {t.idx + 1}"
                if t.kind not in _FS_KINDS + ("SCE", "NOTE"):
                    P._log(col, log, f"{head}: 표 인식 못 함(기타)")
                    floor = t.last_row
                    continue
                if t.kind == "NOTE" and f.kind == "half":       # 반기는 재무제표만(DART 규칙) — 열 해석도 하지 않는다
                    P._log(col, log, f"{head}: 반기 파일의 주석 표는 쓰지 않음(반기는 재무제표만)")
                    floor = t.last_row
                    continue
                unit_known = t.unit is not None or f.unit is not None
                unit = t.unit if t.unit is not None else (f.unit or 1)
                sj = t.kind if t.kind != "NOTE" else None
                cols, msgs = _resolve_columns(t, f, sj, layout.get(key), key, plan.inferred)
                recs = []
                if t.kind in _FS_KINDS:
                    recs = _fs_records(t, rows, f, cols, unit)
                    col.units[(f.fs_div, t.kind)] = unit
                elif t.kind == "SCE":
                    recs = _sce_records(t, rows, f, cols, unit, msgs)
                    col.units[(f.fs_div, "SCE")] = unit
                else:
                    caps, heading = _texts_above(rows, t, floor)
                    recs, (no, ntitle, topic, rec) = _note_records(t, rows, f, cols, unit, unit_known, mapping, name,
                                                                   caps, heading, msgs)
                    # 목차는 제목(정규화)으로 묶고 번호·제목은 결산일이 늦은 파일 것(파일을 오름차순으로 돌므로 덮어쓴다)
                    e = toc.setdefault(N.norm_label(ntitle), [NoteEntry(no, ntitle, N.norm_label(ntitle), topic, rec, 0, 0), 0])
                    e[0].no, e[0].title = no, ntitle
                    e[1] += 1
                for c, g in t.period_cols.items():
                    # 회차도 출처 규칙과 같이: 연간 키는 연간 파일, 반기 키는 반기 파일에서만(반기 BS 전기말 열이 연간 회차를 덮지 않게)
                    if c in cols and g.fiscal is not None and _kind_of_key(cols[c][0].key) == f.kind:
                        col.units[("fiscal", cols[c][0].key)] = g.fiscal
                used.setdefault(name, []).append(t.idx + 1)
                if _unit_text(unit, unit_known) not in units_used:
                    units_used.append(_unit_text(unit, unit_known))
                if t.kind == "NOTE":
                    note_by_report.setdefault(f.rcept_no, []).extend(recs)
                else:
                    by_report.setdefault(f.rcept_no, []).extend(recs)
                pers = {}
                for r in recs:
                    pers.setdefault(r.period_key, r.role)
                ptxt = "·".join(f"{pk}({_ROLE_KO.get(role, role)})"
                                + (HALF_FY_NOTE if _kind_of_key(pk) != f.kind else "")
                                for pk, role in pers.items()) or "기간 없음"
                P._log(col, log, f"{head}: {t.kind} {ptxt} 단위 {_unit_text(unit, unit_known)}")
                for m in msgs:
                    P._log(col, log, f"{head}: {m}")
                floor = t.last_row
        if not n_tables:
            P._log(col, log, f"파일 {f.basename}: {NO_TABLE_NOTE}")
        plan.file_meta[f.rcept_no] = {"sheets": [f"{sn}(표 {'·'.join(map(str, ns))})" for sn, ns in used.items()],
                                      "unit": "·".join(units_used)}
    if not plan.reports:
        raise RuntimeError("쓸 수 있는 입력 파일이 없습니다(로그 확인).")

    # 기간: 파일 자신의 결산 기간은 항상, 머리글에서 온 전기·전전기는 연도 범위 안(반기는 include_half)일 때만
    def in_range(p):
        return cfg.start_year <= p.end.year <= cfg.end_year and (p.kind == "annual" or cfg.include_half)
    def same_kind(rc, pk):
        """연간 키(FY)는 연간 파일에서만, 반기 키(H1)는 반기 파일에서만(DART: 12월 보고서만 FY 출처) — 반기 파일의
        재무상태표 전기말 열은 기록하되 FY 열을 만들거나 그 출처가 되지 않는다."""
        return plan.reports[rc].kind == ("annual" if pk.startswith("FY") else "half")
    found = {}
    for src in (by_report, note_by_report):
        for rc, lst in src.items():
            for r in lst:
                if same_kind(rc, r.period_key):
                    found.setdefault(r.period_key, r)
    keep = {p.key for p in own}
    for pk in found:
        p = _period_of_key(pk)
        if pk in keep:
            continue
        if p is not None and in_range(p):
            keep.add(pk)
        else:
            P._log(col, log, f"기간 {pk}: 연도 범위({cfg.start_year}~{cfg.end_year}{', 반기 제외' if not cfg.include_half else ''}) 밖 → 제외")
    fiscal = {pk: col.units.get(("fiscal", pk)) for pk in keep}
    pers = [_period_of_key(pk) for pk in keep]
    pers = [_period(p.kind, p.end.year, fiscal.get(p.key)) for p in pers]
    plan.periods = sorted([p for p in pers if p.kind == "annual"], key=lambda p: p.end) \
        + sorted([p for p in pers if p.kind == "half"], key=lambda p: p.end)
    col.periods = [p.key for p in plan.periods]

    # 출처 순서: 결산일 늦은 파일 먼저(재작성 기준), 같은 파일 안에서는 당기·전기·전전기
    order = {f.rcept_no: i for i, f in enumerate(sorted(files, key=_sort_key, reverse=True))}

    def sources_of(recs_by_rc):
        src = {}
        for rc, lst in recs_by_rc.items():
            for r in lst:
                pair = (rc, r.role)
                if (r.period_key in col.periods and same_kind(rc, r.period_key)
                        and pair not in src.setdefault(r.period_key, [])):
                    src[r.period_key].append(pair)
        for pk in src:
            src[pk].sort(key=lambda x: (order.get(x[0], 99), _ROLE_OFF.get(x[1], 9), x[1]))
        return src
    plan.sources = sources_of(by_report)
    annual = [pk for pk in col.periods if pk.startswith("FY")]
    plan.note_sources = {pk: v for pk, v in sources_of(note_by_report).items() if pk in annual}
    annual_files = [f for f in files if f.kind == "annual" and f.rcept_no in plan.reports]
    plan.base_report = plan.reports[max(annual_files, key=_sort_key).rcept_no] if annual_files else None

    col.fs_records, col.restate_diffs = P._pick_restated(
        col.periods, plan.sources, by_report, lambda r: (r.fs_div, r.sj_div), P._fs_cmp,
        lambda r: (r.label_norm, r.detail or ""))

    # 주석: 목차(include 는 cfg.note_include 규칙) → 선택된 주석만 재작성 기준으로 확정, 원표는 연도별 그대로
    entries = sorted(toc.values(), key=lambda e: (int(e[0].no) if str(e[0].no).isdigit() else 999, e[0].title))
    for e, n in entries:
        e.table_count = n
        e.include = P._wanted(e, cfg)
    col.toc = [e for e, _ in entries]
    wanted = {e.title_norm for e in col.toc if e.include}
    note_by_report = {rc: [r for r in lst if N.norm_label(r.note_title) in wanted] for rc, lst in note_by_report.items()}
    col.note_records, note_diffs = P._pick_restated(
        annual, plan.note_sources, note_by_report, lambda r: (r.fs_div, N.norm_label(r.note_title)),
        P._note_cmp, lambda r: (r.row_label, r.col_label or ""))
    col.restate_diffs.extend(note_diffs)
    seen = set()
    for rc in sorted(note_by_report, key=lambda rc: order.get(rc, 99), reverse=True):
        for r in note_by_report[rc]:
            k = (r.rcept_no, r.fs_div, r.note_no, r.table_idx, r.raw_cell, r.row_label, r.col_label)
            if r.raw_cell is not None and r.period_key in annual and k not in seen:
                seen.add(k)
                col.note_records.append(r)
    return col
