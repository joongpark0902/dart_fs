"""회사 제시 엑셀 재무제표의 표 인식(파일 입력 1단계).

시트를 위→아래로 훑어 머리글 행(기간 표식 셀 2개 이상, 또는 자본변동표·변동표처럼 구성요소 열 머리글)을 찾고,
그 아래 본문 행·라벨 열·값 열·기간 열·단위를 읽어 표 유형(BS/IS/CIS/CF/SCE/NOTE/OTHER)을 판정한다.
순수 openpyxl(read_only·data_only) — 원본을 고치지 않고, 캐시·로그 파일·네트워크·COM 을 쓰지 않는다.
행·열 번호는 엑셀과 같은 1부터.
"""
import datetime as _dt
import re
from dataclasses import dataclass, field, asdict

from openpyxl import load_workbook

from core import normalize as N
from core import mapping as _mapping
from core.note_std import detect_topic

_DASHES = {"-", "－", "—", "―", "–"}           # note_parser._DASHES 와 같은 대시 규칙
_MAX_ROWS = 20000
_MAX_COLS = 80

_DATE = re.compile(r"(\d{4})\s*[.\-/년]\s*(\d{1,2})\s*[.\-/월]\s*(\d{1,2})\s*일?")
_YEAR = re.compile(r"(?:FY\s*(\d{4})|(\d{4})\s*년)", re.I)
_DATE_WORDS = re.compile(r"부터|까지|현재|\(\s*\)")          # 날짜를 뗀 뒤 남는 꼬리말('…부터 …까지', '… 현재')
_UNIT_BARE = re.compile(r"^\(\s*([^()]{1,6}?)\s*\)$")
_NOTE_NO = re.compile(r"주석\s*\d+")
_NOTE_REF_CELL = re.compile(r"^(주석\s*)?[\d,\s~\-·]+$")     # 머리글 없는 주석 번호 열의 셀('주석 4'·'4,5'·'4~6')
_EXCL_HDR = re.compile(r"주석|코드|비고|^no\.?$|번호", re.I)   # 숫자여도 값 열이 아닌 머리글
_MAX_FORMULA_ROWS = 2000                                    # 수식만 있는 파일 판정 때 읽는 행 수
ROLE_OFF = {"cur": 0, "prev": 1, "prev2": 2}
OFF_ROLE = {v: k for k, v in ROLE_OFF.items()}

_SHEET_KIND = [("재무상태표", "BS"), ("대차대조표", "BS"), ("포괄손익계산서", "CIS"), ("손익계산서", "IS"),
               ("현금흐름표", "CF"), ("자본변동표", "SCE")]
_BS_KW = ["자산총계", "부채총계", "자본총계"]
_CF_KW = ["영업활동", "투자활동", "재무활동"]
_SCE_KW = ["자본금", "이익잉여금", "결손금"]
_SGA_ITEMS = ["급여", "퇴직급여", "복리후생비", "감가상각비", "지급수수료", "광고선전비", "접대비", "임차료",
              "원재료사용", "종업원급여", "운반비", "세금과공과"]


@dataclass
class PeriodGuess:
    text: str
    role: str = None          # "cur" | "prev" | "prev2" | None
    end_date: str = None      # "YYYY-MM-DD"(날짜) 또는 "YYYY"(연도만), 없으면 None
    fiscal: int = None        # '제 N 기' 의 N

    def to_dict(self):
        return asdict(self)


@dataclass
class TableScan:
    idx: int
    kind: str
    title: str
    header_row: int
    first_row: int
    last_row: int
    label_cols: list
    value_cols: list
    period_cols: dict
    unit: int = None
    unit_text: str = None
    confidence: str = "low"
    reason: str = ""
    col_texts: dict = field(default_factory=dict)   # 값 열(1부터) → 머리글 글자(상위 기간 + 하위 '3개월'·'누적' 등)
    header_start: int = None                        # 머리글 첫 행(1부터). 2행 머리글이면 header_row(아래 행)보다 위

    def to_dict(self):
        d = asdict(self)
        d["period_cols"] = {str(c): g.to_dict() for c, g in self.period_cols.items()}
        d["col_texts"] = {str(c): t for c, t in self.col_texts.items()}
        return d


@dataclass
class SheetScan:
    name: str
    tables: list = field(default_factory=list)
    hidden: bool = False
    orphan_rows: list = field(default_factory=list)   # 라벨+숫자가 있지만 어느 표에도 속하지 않은 행(1부터)
    note: str = None                                   # 표를 못 찾은 이유(수식만 있어 계산값이 없는 시트 등)

    def to_dict(self):
        return {"name": self.name, "hidden": self.hidden, "tables": [t.to_dict() for t in self.tables],
                "orphan_rows": list(self.orphan_rows), "note": self.note}


@dataclass
class FileScan:
    path: str
    sheets: list = field(default_factory=list)

    def to_dict(self):
        return {"path": self.path, "sheets": [s.to_dict() for s in self.sheets]}


# ---------------------------------------------------------------- 셀 판정

def _text(v):
    if v is None:
        return ""
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y.%m.%d")
    return N.clean_text(str(v))


def _is_dash(v):
    return isinstance(v, str) and N.clean_text(v) in _DASHES


def _is_num(v):
    if isinstance(v, bool) or v is None:
        return False
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        return N.parse_amount(v) is not None or N.parse_percent(v) is not None
    return False


def _is_text(v):
    return isinstance(v, str) and bool(N.clean_text(v)) and not _is_num(v) and not _is_dash(v)


def unit_of(text):
    """'(단위 : 천원)'·'단위: 백만원'·'(천원)' → 배수. 단위 문구가 아니면 None."""
    t = N.clean_text(text)
    if not t:
        return None
    u = N.parse_unit(t)
    if u is None and "단위" in t:
        u = N.parse_unit(t + ")")
    if u is None:
        m = _UNIT_BARE.match(t)
        if m:
            u = N.unit_multiplier(m.group(1))
    return u


def period_guess(v):
    """머리글 셀 하나 → PeriodGuess. 기간 표식(당기/전기·제 N 기·날짜·연도)이 없으면 None."""
    if isinstance(v, (_dt.datetime, _dt.date)):
        return PeriodGuess(_text(v), end_date=v.strftime("%Y-%m-%d"))
    if not isinstance(v, str):
        return None
    t = N.clean_text(v)
    if not t:
        return None
    # 날짜·'부터/까지/현재' 를 뗀 핵심 글자로 길이 제한·역할·회차를 본다 — '제 12(당) 기 2025년 1월 1일부터 2025년 12월 31일까지'
    core = N.clean_text(_DATE_WORDS.sub(" ", _DATE.sub(" ", t)))
    if len(N.compact(core)) > 30:
        return None
    role = N.period_role(core)
    fiscal = N.fiscal_no(core)
    end = None
    for m in _DATE.finditer(t):             # 기간 표기('…부터 …까지')는 마지막 날짜가 기말
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 1 <= mo <= 12 and 1 <= d <= 31:
            end = f"{y:04d}-{mo:02d}-{d:02d}"
    if end is None:
        m = _YEAR.search(t)
        if m:
            end = m.group(1) or m.group(2)
    if role is None:   # '당기 금액'·'전기말(2024.12.31)': '금액'·'합계'를 떼고 다시
        rest = re.sub(r"금액|합계", "", N.compact(core))
        role = N.period_role(rest) if rest else None
    if role is None and fiscal is None and (end is None or len(N.compact(core)) > 10):
        return None    # 날짜뿐인 셀은 짧은 꼬리말('현재'·'기말')까지만 — 날짜가 든 문장은 머리글이 아니다
    return PeriodGuess(t, role, end, fiscal)


def _merge_guess(gs):
    """2행 머리글: 같은 열의 판정을 합친다(빈 항목은 다른 행 값으로)."""
    gs = [g for g in gs if g]
    if not gs:
        return None
    out = PeriodGuess(" ".join(g.text for g in gs))
    for g in gs:
        out.role = out.role or g.role
        out.end_date = out.end_date or g.end_date
        out.fiscal = out.fiscal if out.fiscal is not None else g.fiscal
    return out


# ---------------------------------------------------------------- 시트 읽기

def _cell(row, c):
    return row[c] if c < len(row) else None


def _empty(row):
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in row)


_YEAR_STR = re.compile(r"(\d{4})(?:\s*[.\-]\s*(\d{1,2}))?")


def _year_like(v):
    """머리글의 연도·연월(int 2025, '2025', '2026.06', '2026-06') → end_date('YYYY' 또는 'YYYY-MM'), 아니면 None.
    본문 숫자와 구분이 안 되므로 머리글 행 판정·기간 열에서만 쓴다."""
    if isinstance(v, bool):
        return None
    if isinstance(v, int):
        return str(v) if 1990 <= v <= 2100 else None
    if not isinstance(v, str):
        return None
    m = _YEAR_STR.fullmatch(N.clean_text(v))
    if not m or not 1990 <= int(m.group(1)) <= 2100:
        return None
    if m.group(2) is None:
        return m.group(1)
    mo = int(m.group(2))
    return f"{m.group(1)}-{mo:02d}" if 1 <= mo <= 12 else None


def _year_seq(ends):
    """연도만인 머리글(2025·2024)은 1년 간격 연속(오름 또는 내림)일 때만 인정 — 금액 2,100·1,995 와 구분.
    연월('2026-06')이 섞이면 연속 조건을 적용하지 않는다."""
    if any("-" in e for e in ends):
        return True
    ys = [int(e) for e in ends]
    d = [b - a for a, b in zip(ys, ys[1:])]
    return all(x == 1 for x in d) or all(x == -1 for x in d)


def _hdr_guess(v):
    end = _year_like(v)
    if end:
        return PeriodGuess(_text(v), end_date=end)
    return period_guess(v)


def _period_cells(row):
    return [c for c, v in enumerate(row) if period_guess(v)]


def _col_kind(rows, c):
    """note_parser._col_kind 와 같은 규칙: 대시·빈칸은 분모에서 빼고, 남은 값이 없으면 num, 숫자 비율 ≥ 0.6 이면 num."""
    vals = [_cell(r, c) for r in rows]
    vals = [v for v in vals if v is not None and _text(v)]
    if not vals:
        return "empty"
    rest = [v for v in vals if not _is_dash(v)]
    if not rest:
        return "num"
    numeric = sum(1 for v in rest if _is_num(v))
    return "num" if numeric >= len(rest) * 0.6 else "text"


def _is_heading(text):
    return bool(text) and (N.split_heading(text) is not None or bool(_NOTE_NO.match(N.clean_text(text))))


def _comp_header(rows, i):
    """구성요소 열 머리글(자본변동표·변동표·시산표): 숫자 없는 행에 글자 셀 ≥ 2,
    다음 비어 있지 않은 행에 라벨 글자 + 그 머리글 열들 중 숫자(또는 '-') 셀 ≥ 2."""
    row = rows[i]
    texts = [c for c, v in enumerate(row) if _is_text(v)]
    if len(texts) < 2 or any(_is_num(v) for v in row):
        return False
    for j in range(i + 1, min(i + 3, len(rows))):
        nxt = rows[j]
        if _empty(nxt):
            continue
        if not any(_is_text(v) for v in nxt):
            return False
        hits = sum(1 for c in texts if _is_num(_cell(nxt, c)) or _is_dash(_cell(nxt, c)))
        return hits >= 2
    return False


def _header_at(rows, i):
    """i 행이 머리글이면 (머리글 행 목록, 'period'|'comp'), 아니면 None."""
    row = rows[i]
    pc = _period_cells(row)
    years = [c for c, v in enumerate(row) if _year_like(v)]
    year_row = (len(years) >= 2 and _year_seq([_year_like(row[c]) for c in years]) and not pc and any(_is_text(v) for v in row)
                and all(_year_like(v) for v in row if _is_num(v)) and i + 1 < len(rows)
                and any(_is_text(v) for v in rows[i + 1]) and any(_is_num(v) for v in rows[i + 1]))
    if year_row:                                        # 머리글이 연도·연월(2025·'2026.06')
        return [i], "period"
    if len(pc) >= 2 and not any(_is_num(v) for v in row):
        hdr = [i]
        if i + 1 < len(rows):
            nxt = rows[i + 1]
            if len(_period_cells(nxt)) >= 2 and not any(_is_num(v) for v in nxt):
                hdr.append(i + 1)
            elif (not _empty(nxt) and not any(_is_num(v) for v in nxt)
                  and all(not _text(v) for v in nxt[:min(pc)])
                  and any(_text(v) for v in nxt[min(pc):])):
                hdr.append(i + 1)                       # '금액'·'구성비' 같은 보조 머리글 행
        return hdr, "period"
    if _comp_header(rows, i):
        return [i], "comp"
    return None


def _header_soon(rows, r, span=3):
    """r 행 바로 아래 span 개의 비어 있지 않은 행 안에 머리글이 오는가(주석 제목 + 새 표)."""
    seen = 0
    for j in range(r + 1, len(rows)):
        if _empty(rows[j]):
            continue
        if _header_at(rows, j):
            return True
        seen += 1
        if seen >= span:
            break
    return False


def _body_end(rows, start):
    """본문 마지막 행(0부터). 마지막 행은 숫자 또는 '-' 가 있는 행까지만 — 설명 문장·구분 행은 last 로 치지 않는다.
    끝내는 조건: 빈 행 2개, 다음 기간 머리글, 단위 문구, 구성요소 머리글, 주석 제목(바로 아래 3행 안에 머리글이 오는
    경우만 — 빈 행 뒤 '2. 비유동자산' 같은 번호 붙은 구분 행은 제목이 아니다), 빈 행 뒤 글자 행 중 다음 행도 숫자·'-' 가
    없는 것(설명 문장). 그 밖의 글자 행(빈 행 뒤 '부채' 등)은 구분 행으로 보고 계속 읽는다."""
    last, blanks, r = start - 1, 0, start
    while r < len(rows):
        row = rows[r]
        if _empty(row):
            blanks += 1
            if blanks >= 2:
                break
            r += 1
            continue
        has_num = any(_is_num(v) for v in row)
        if len(_period_cells(row)) >= 2 and not has_num:
            break
        if not has_num and not any(_is_dash(v) for v in row):
            if any(unit_of(_text(v)) for v in row if _is_text(v)) or _comp_header(rows, r):
                break
            if _is_heading(next((_text(v) for v in row if _is_text(v)), "")) and _header_soon(rows, r):
                break
            if blanks:
                nxt = next((x for x in rows[r + 1:] if not _empty(x)), None)
                if nxt is None or not any(_is_num(v) or _is_dash(v) for v in nxt):
                    break                   # 빈 행 뒤 설명 문장
            blanks = 0                      # 구분 행('자산'·빈 행 뒤 '부채'): 계속 보되 마지막 행으로는 치지 않음
            r += 1
            continue
        last, blanks = r, 0
        r += 1
    return last


# ---------------------------------------------------------------- 유형 판정

def _kind(labels, header_texts, title, sheet_name, n_period, heading, mapping):
    """(kind, 키워드 점수, 근거 문구 목록)."""
    lab = [N.norm_label(x) for x in labels if x]
    hdr = [N.norm_label(x) for x in header_texts if x]
    ctitle = N.compact(title or "")
    csheet = N.compact(sheet_name or "")

    def weight(kind):
        w, why = 0, []
        for src, text in (("시트명", csheet), ("제목", ctitle)):
            for word, k in _SHEET_KIND:
                if word in text:          # 가장 먼저 맞는 이름만('포괄손익계산서' ⊃ '손익계산서')
                    if k == kind:
                        w, why = w + 1, why + [f"{src} '{word}'"]
                    break
        return w, why

    if any("차변" in h for h in hdr) and any("대변" in h for h in hdr):
        return "OTHER", 0, ["머리글 차변·대변(시산표)"]

    # SCE: 제목 '자본변동' 또는 머리글 구성요소 2개 이상(열 라벨이 구성요소)
    sce_hit = [k for k in _SCE_KW if any(k in h for h in hdr)]
    w, why = weight("SCE")
    if "자본변동" in ctitle or len(sce_hit) >= 2 or (sce_hit and w):
        kws = sce_hit + (["자본변동"] if "자본변동" in ctitle else [])
        return "SCE", len(kws), (["머리글 " + "·".join(sce_hit)] if sce_hit else []) + why

    def has(word):
        return any(word in x for x in lab)

    is_groups = []
    if any(x in ("매출액", "매출", "수익", "영업수익") or x.startswith(("매출액", "수익(")) for x in lab):
        is_groups.append("매출액")
    if has("영업이익") or has("영업손실"):
        is_groups.append("영업이익")
    if any(has(w_) for w_ in ("당기순이익", "당기순손실", "반기순이익", "반기순손실", "분기순이익", "분기순손실")):
        is_groups.append("당기순이익")
    cis_hit = [k for k in ("총포괄", "기타포괄") if has(k)]
    bs_hit = [k for k in _BS_KW if has(k)]
    cf_hit = [k for k in _CF_KW if has(k)]

    if "총포괄" in cis_hit:
        hits = cis_hit + [g for g in is_groups if g == "당기순이익"]
        w, why = weight("CIS")
        return "CIS", len(hits), ["라벨 " + "·".join(hits)] + why

    cands = []
    for kind, hits in (("BS", bs_hit), ("IS", is_groups), ("CF", cf_hit)):
        w, why = weight(kind)
        if kind == "IS":
            w2, why2 = weight("CIS")
            w, why = w + w2, why + why2
        if hits and len(hits) + w >= 2:
            cands.append((len(hits) + w, len(hits), kind, ["라벨 " + "·".join(hits)] + why))
    if cands:
        cands.sort(key=lambda x: (-x[0], -x[1]))
        _, n, kind, why = cands[0]
        return kind, n, why

    # NOTE: 제목(또는 위쪽 주석 제목)이 'N. 주제' / '주석 N' + 주제 키워드, 또는 라벨에 판관비·성격별 항목
    for t in (title, heading):
        if t and _is_heading(t):
            sh = N.split_heading(t)
            body = sh[1] if sh else _NOTE_NO.sub("", N.clean_text(t))
            topic, _ = detect_topic(N.norm_label(body), mapping)
            if topic:
                return "NOTE", 2, [f"주석 제목 '{N.clean_text(t)}'", f"주제 {topic}"]
    items = [k for k in _SGA_ITEMS if k in lab]
    if len(items) >= 2 and n_period >= 1:
        return "NOTE", len(items), ["라벨 " + "·".join(items)]
    return "OTHER", 0, ["유형 키워드 없음"]


# ---------------------------------------------------------------- 표 만들기

def _title_of(rows, h0, floor):
    """머리글 위 3행 안(앞 표 끝 아래)의 글자 셀. 주석 제목·재무제표 이름을 먼저, 없으면 가장 가까운 글자."""
    cands = []
    for r in range(h0 - 1, max(floor, h0 - 4), -1):
        for v in rows[r]:
            if _is_text(v) and unit_of(_text(v)) is None:
                cands.append(_text(v))
    for t in cands:
        if _is_heading(t) or any(w in N.compact(t) for w, _ in _SHEET_KIND):
            return t
    return cands[0] if cands else None


def _infer_roles(guesses):
    """'제 N 기' 만 있고 당/전 표기가 없는 열: 같은 표에 역할과 회차가 함께 적힌 열이 있을 때만 회차 차이로 채운다.
    그런 기준 열이 없으면 None 으로 둔다(상위에서 파일 결산일·기수로 정함)."""
    known = [g.fiscal + ROLE_OFF[g.role] for g in guesses if g.role and g.fiscal is not None]
    if not known:
        return
    base = known[0]
    for g in guesses:
        if g.role is None and g.fiscal is not None:
            g.role = OFF_ROLE.get(base - g.fiscal)


def _period_guesses(rows, hdr, width):
    """열마다 머리글 행들의 판정을 합친다 → ({열: PeriodGuess|None}, {열: 머리글 글자}).
    기간 셀 오른쪽의 빈 머리글 칸은 그 기간의 하위 열로 본다('당반기 | (빈칸)' 위에 '3개월 | 누적')."""
    per_col = {c: [] for c in range(width)}
    texts = {c: [] for c in range(width)}
    for r in hdr:
        span = None
        for c in range(width):
            v = _cell(rows[r], c)
            if _text(v):
                span = _hdr_guess(v)
                per_col[c].append(span)
                texts[c].append(_text(v))
            elif span:
                per_col[c].append(span)
                texts[c].append(span.text)
    guesses = {}
    for c in range(width):
        g = _merge_guess(per_col[c])
        if g:
            g.text = " ".join(texts[c])
        guesses[c] = g
    return guesses, {c: " ".join(texts[c]) for c in range(width)}


def _note_ref_col(body, c):
    """글자 열의 비어 있지 않은 셀이 대부분 주석 참조('주석 4'·'4,5')면 라벨 열도 값 열도 아니다."""
    vals = [_text(_cell(r, c)) for r in body]
    vals = [v for v in vals if v]
    return bool(vals) and sum(1 for v in vals if _NOTE_REF_CELL.match(v)) * 2 > len(vals)


def _pick_period_cols(value_cols, guesses, hdr_text):
    """같은 기간이 여러 숫자 열에 걸치면(금액·구성비, 3개월·누적) 하위 머리글에 '누적' 이 있는 열, 없으면 첫 열."""
    groups = {}
    for c in value_cols:
        g = guesses[c]
        if g:
            groups.setdefault((g.role, g.end_date, g.fiscal), []).append(c)
    out = {}
    for cs in groups.values():
        pick = next((c for c in cs if "누적" in N.compact(hdr_text[c])), cs[0])
        out[pick + 1] = guesses[pick]
    return dict(sorted(out.items()))


def _build(rows, hdr, mode, floor, sheet_name, heading, unit_state, mapping, label_cols=None):
    """머리글 행(0부터)들 → (TableScan, 마지막 행). label_cols(0부터)를 주면 라벨 열은 그대로 쓴다(사용자 수정값)."""
    h0, hl = hdr[0], hdr[-1]
    last = _body_end(rows, hl + 1)
    if last < hl + 1:
        return None, hl
    body = rows[hl + 1:last + 1]
    width = max(len(r) for r in rows[h0:last + 1])

    kinds = {c: _col_kind(body, c) for c in range(width)}
    if mode == "period":
        guesses, hdr_text = _period_guesses(rows, hdr, width)
        anchor = min(c for c, g in guesses.items() if g)
        lo = 0
    else:
        guesses = {c: None for c in range(width)}
        hdr_text = {c: " ".join(_text(_cell(rows[r], c)) for r in hdr if _text(_cell(rows[r], c))) for c in range(width)}
        t0 = next((c for c in range(width) if kinds[c] == "text"), None)
        if t0 is None:
            return None, hl
        anchor = next((c for c in range(t0 + 1, width) if kinds[c] == "num"), width)
        lo = t0
    if label_cols is None:
        label_cols = [c for c in range(lo, anchor) if kinds[c] == "text" and not _note_ref_col(body, c)][:2]
    else:
        label_cols = [c for c in label_cols if 0 <= c < width]
    start = (max(label_cols) + 1) if label_cols else 0
    value_cols = [c for c in range(start, width)
                  if kinds[c] == "num" and not _EXCL_HDR.search(N.compact(hdr_text[c]))]
    if not value_cols:
        return None, hl

    _infer_roles([guesses[c] for c in value_cols if guesses[c]])
    period_cols = _pick_period_cols(value_cols, guesses, hdr_text)

    # 단위: 앞 표 끝 ~ 머리글 행 사이 문구, 없으면 시트 위쪽에서 마지막으로 본 단위
    unit, unit_text = None, None
    for r in range(floor + 1, hl + 1):
        for v in rows[r]:
            if _is_text(v) and unit_of(_text(v)):
                unit, unit_text = unit_of(_text(v)), _text(v)
    carried = False
    if unit is None and unit_state:
        unit, unit_text = unit_state
        carried = True

    title = _title_of(rows, h0, floor)
    labels = [" ".join(_text(_cell(r, c)) for c in label_cols if _text(_cell(r, c))) for r in body]
    hdr_labels = [hdr_text[c] for c in range(width) if c not in label_cols and hdr_text[c]]
    kind, n_kw, why = _kind(labels, hdr_labels, title, sheet_name, len(period_cols), heading, mapping)

    good = (n_kw >= 2) + (len(period_cols) >= 2)
    confidence = "low" if kind == "OTHER" else {2: "high", 1: "mid"}.get(good, "low")
    reason = " / ".join(why + [f"기간 열 {len(period_cols)}개"] + (["단위 시트 상단에서 이월"] if carried else []))
    t = TableScan(idx=0, kind=kind, title=title, header_row=hl + 1, first_row=hl + 2, last_row=last + 1,
                  label_cols=[c + 1 for c in label_cols], value_cols=[c + 1 for c in value_cols],
                  period_cols=period_cols, unit=unit, unit_text=unit_text, confidence=confidence, reason=reason,
                  col_texts={c + 1: hdr_text[c] for c in value_cols}, header_start=h0 + 1)
    return t, last


def rebuild(rows, t, *, header_row=None, label_cols=None, sheet_name="", mapping=None):
    """사용자 수정값(머리글 행·라벨 열, 둘 다 1부터)으로 표를 다시 만든다 — 기간 열·값 열·본문 범위가 수정값과 맞게.
    다시 만들지 못하면(값 열 없음 등) None. idx 는 원 표 것, 단위는 새로 못 찾으면 원 표 것."""
    mapping = mapping or _mapping.load()
    if header_row is None or int(header_row) == t.header_row:
        # 머리글 행을 바꾸지 않았으면 머리글 첫 행부터(2행 머리글의 위 행 '당기·전기' 표식을 잃지 않게)
        h = int(t.header_start or t.header_row) - 1
    else:
        h = int(header_row) - 1
    if not 0 <= h < len(rows):
        return None
    hdr = _header_at(rows, h)
    if hdr:
        hdr_rows, mode = hdr
    else:
        hdr_rows, mode = [h], ("period" if _period_cells(rows[h]) else "comp")
    lc = [int(c) - 1 for c in label_cols] if label_cols else None
    above = [_text(v) for r in range(h - 1, -1, -1) for v in rows[r] if _is_text(v)]      # 가까운 행부터
    heading = next((s for s in above if _is_heading(s)), None)
    unit_state = next(((unit_of(s), s) for s in above if unit_of(s)), None)
    new, _ = _build(rows, hdr_rows, mode, max(-1, h - 4), sheet_name, heading, unit_state, mapping, label_cols=lc)
    if new is None:
        return None
    new.idx = t.idx
    if new.unit is None and t.unit is not None:
        new.unit, new.unit_text = t.unit, t.unit_text
    return new


def _read_rows(ws):
    try:
        ws.reset_dimensions()          # 더존 출력처럼 dimension 이 'A1' 로 깨진 파일도 끝까지 읽는다
    except AttributeError:
        pass
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i >= _MAX_ROWS:
            break
        row = list(row[:_MAX_COLS])
        while row and (row[-1] is None or (isinstance(row[-1], str) and not row[-1].strip())):
            row.pop()
        rows.append(row)
    return rows


def scan_rows(rows, sheet_name="", mapping=None):
    """값 2차원 목록(행마다 list) → [TableScan]. scan() 의 시트 단위 본체."""
    return _scan_sheet(rows, sheet_name, mapping)[0]


def _scan_sheet(rows, sheet_name="", mapping=None):
    """→ (tables, orphan_rows). orphan_rows = 표 범위(머리글~마지막 행) 밖에서 글자와 숫자가 함께 있는 행 번호(1부터)."""
    mapping = mapping or _mapping.load()
    tables, i, floor, heading, unit_state = [], 0, -1, None, None
    covered = set()
    while i < len(rows):
        row = rows[i]
        if _empty(row):
            i += 1
            continue
        hdr = _header_at(rows, i)
        if hdr:
            t, end = _build(rows, hdr[0], hdr[1], floor, sheet_name, heading, unit_state, mapping)
            if t:
                t.idx = len(tables)
                tables.append(t)
                if t.unit is not None:
                    unit_state = (t.unit, t.unit_text)
                covered.update(range(hdr[0][0], end + 1))
                floor, i = end, end + 1
                continue
            i = hdr[0][-1] + 1
            continue
        for v in row:
            if _is_text(v):
                if unit_of(_text(v)):
                    unit_state = (unit_of(_text(v)), _text(v))
                elif _is_heading(_text(v)):
                    heading = _text(v)
        i += 1
    orphans = [r + 1 for r, row in enumerate(rows) if r not in covered
               and any(_is_text(v) for v in row) and any(_is_num(v) for v in row)]
    return tables, orphans


def scan_sheet(rows, sheet_name="", mapping=None):
    """시트 한 장 → (tables, orphan_rows). file_collect 가 읽어 둔 rows 로 다시 쓰기 위한 공개 이름."""
    return _scan_sheet(rows, sheet_name, mapping)


FORMULA_NOTE = "글자만 있고 숫자가 없음 — 수식 셀의 계산값이 없는 파일(Excel 에서 한 번 저장한 뒤 다시 넣으세요)"


def _has_formula(path, names):
    """계산값 없이 수식만 저장된 파일인가: data_only=False 로 다시 열어 해당 시트들에 '=' 로 시작하는 셀이 있는지."""
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        for ws in wb.worksheets:
            if ws.title not in names:
                continue
            try:
                ws.reset_dimensions()
            except AttributeError:
                pass
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i >= _MAX_FORMULA_ROWS:
                    break
                if any(isinstance(v, str) and v.startswith("=") for v in row[:_MAX_COLS]):
                    return True
        return False
    finally:
        wb.close()


def load_sheets(path, *, max_sheets=60):
    """엑셀 파일 → [(시트명, 숨김 여부, rows, note)]. 원본은 읽기만 한다(read_only·data_only).
    note: 글자 셀은 있는데 숫자 셀이 하나도 없고 파일에 수식이 있으면 FORMULA_NOTE(계산값 없는 수식 파일), 아니면 None."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        sheets = [(ws.title, getattr(ws, "sheet_state", "visible") != "visible", _read_rows(ws))
                  for ws in wb.worksheets[:max_sheets]]
    finally:
        wb.close()
    suspect = {name for name, _, rows in sheets
               if any(_is_text(v) for r in rows for v in r) and not any(_is_num(v) for r in rows for v in r)}
    formula = _has_formula(path, suspect) if suspect else False
    return [(name, hidden, rows, FORMULA_NOTE if formula and name in suspect else None) for name, hidden, rows in sheets]


def scan(path, *, max_sheets=60, mapping=None):
    """엑셀 파일 → FileScan."""
    mapping = mapping or _mapping.load()
    out = FileScan(path=str(path))
    for name, hidden, rows, note in load_sheets(path, max_sheets=max_sheets):
        tables, orphans = _scan_sheet(rows, name, mapping)
        out.sheets.append(SheetScan(name, tables, hidden, orphans, note))
    return out
