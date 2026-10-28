"""Collected → WorkbookModel. 시트 배치 규칙은 설계 §3.5·§5.

raw 시트(rFS_*·rN_*)는 블록마다 공시 단위를 그대로 두고, 단위 환산은 mFS 수식에서만 한다.
mFS 값 수식은 `=raw*공시단위/cfg_unit` — 설정 시트 cfg_unit(1·1000·1000000)을 바꾸면 [업데이트] 없이 즉시 바뀐다.
표시 단위 0(자동)이면 주 재무제표(연결 선택 시 연결, 아니면 별도) 재무상태표의 공시 단위로 판정해 cfg_unit 에 쓴다.
mFS 는 실제로 매핑된 raw 행이 있는 표준계정만 만들고(합계를 지어내지 않음), 매핑 안 된 raw 행은
"미매핑" 플래그로 뒤에 붙여 아무것도 버리지 않는다.
검증식은 `검증` 시트(excel/checks.py)에 두고 mFS 에는 fs_div 별 요약 3행(TRUE·FALSE·해당 없음 수)만 남긴다.
"""
import dataclasses, json, re
from collections import OrderedDict, namedtuple
from datetime import date, datetime, timedelta
from core import layout as L, normalize as N
from core.model import Period
from core.note_std import sheet_for
from excel.grid import Sheet, WorkbookModel

RAW_LABEL_COL, RAW_FIRST_VALUE_COL, RAW_GAP_ROWS = 4, 5, 5
BLOCK_MARK_COL, BLOCK_MARK = 3, "§"          # 블록 제목 행 C열 표식(업데이트 때 블록 경계 탐지용)
DATA_SHEETS = ("목차", "mFS", "검증", "RAW >>", "데이터", "매핑표", "로그")
ETC_SHEET = "rN_기타"
_SJ_TITLE = {"BS": "재무상태표", "IS": "손익계산서", "CIS": "포괄손익계산서", "CF": "현금흐름표", "SCE": "자본변동표"}
_UNIT_TXT = {1: "원", 1000: "천원", 10**4: "만원", 10**6: "백만원", 10**8: "억원", 10**9: "십억원"}
_RAW_WIDTHS = {1: 2.5, 2: 2.7, 3: 3.6, 4: 44}
VALUE_COL_WIDTH = 15                          # 값(기간) 열 폭 — raw·mFS·검증 공통
_VALUE_FMTS = ("num", "bold", "hdr", "date")
_TOTAL_NORMS = ("합계", "총계", "소계")
UNIT_UNKNOWN = "(단위 : 원문 미표기)"
_SUMMARY_ROWS = (("검증 TRUE 수", "TRUE"), ("검증 FALSE 수", "FALSE"), ("해당 없음 수", '"해당 없음"'))

# raw 시트에 놓인 행 하나: mFS·검증식이 참조한다.
Placed = namedtuple("Placed", "sheet block section row rownum k")


def is_total(norm):
    """합계·총계류 행(굵게 + 아래 실선): 정규화 라벨이 '계' 이거나 합계·총계·소계로 끝난다(자산총계·부채와자본총계 포함).
    당기순이익 같은 이익 행은 아니다."""
    return bool(norm) and (norm == "계" or norm.endswith(_TOTAL_NORMS))


def is_data_sheet(name):
    return name in DATA_SHEETS or name.startswith("rFS_") or name.startswith("rN_")


def _col_letter(c):
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


def _ref(sheet, r, c):
    return f"'{sheet}'!{_col_letter(c)}{r}"


def _fs_name(fs_div):
    return "연결" if fs_div == "CFS" else "별도"


def _fs_divs(col):
    """항상 연결 → 별도 순서."""
    return sorted(col.cfg.fs_divs, key=lambda d: 0 if d == "CFS" else 1)


def _unit_txt(u):
    return _UNIT_TXT.get(u, str(u))


UNIT_NAME = "cfg_unit"
UNIT_CHOICES = (1, 1000, 1000000)
UNIT_LIST_F1 = '"' + ",".join(str(u) for u in UNIT_CHOICES) + '"'
# cfg_unit → 단위 글자(설정 도우미 셀·mFS 제목)
UNIT_TEXT_EXPR = (f'IF({UNIT_NAME}=1,"원",IF({UNIT_NAME}=1000,"천원",'
                  f'IF({UNIT_NAME}=1000000,"백만원","?")))')


def judge_unit(col):
    """공시 단위 자동 판정 → (나눗수, 근거). 주 재무제표 = 연결 선택 시 연결, 아니면 별도.
    그 재무상태표 단위(col.units — XML 은 제목표 '(단위 : 천원)', API 는 항상 원), 없으면 손익계산서, 그것도 없으면 1."""
    fs_div = "CFS" if "CFS" in col.cfg.fs_divs else "OFS"
    sheet = f"rFS_{_fs_name(fs_div)}"
    for sj, name in (("BS", "재무상태표"), ("IS", "손익계산서"), ("CIS", "포괄손익계산서")):
        u = col.units.get((fs_div, sj))
        if u:
            return int(u), f"{sheet} {name} 단위 행"
    return 1, f"{sheet} 단위 정보 없음 → 원"


def display_unit(col):
    """(cfg_unit 에 쓸 나눗수, 자동 여부, 설정 '단위 판정' 문구)."""
    u = col.cfg.unit_display
    if not u:
        v, basis = judge_unit(col)
        return v, True, f"자동 판정: {_unit_txt(v)} (근거: {basis})"
    return int(u), False, f"직접 선택: {_unit_txt(int(u))} (자동 판정 안 함)"


def _unique(title, used):
    t, n = title, 1
    while t in used:
        n += 1
        t = f"{title} #{n}"
    used.add(t)
    return t


# ---------------------------------------------------------------- 기존 구조(업데이트 때 행·열 안정성, 설계 §5)

_PK = re.compile(r"^(?:FY(\d{4})|(\d{4})H1)$")
_TITLE_NO = re.compile(r"^(.*) #(\d+)$")
_EXCEL_EPOCH = datetime(1899, 12, 30)
_EMPTY = {"blocks": {}, "periods": []}


def is_struct_sheet(name):
    """블록 구조(§ 표식)를 읽어 업데이트 때 위치를 지키는 시트."""
    return name == "mFS" or name.startswith("rFS_") or name.startswith("rN_")


def _period_of(v):
    """블록 헤더 2행째 셀 → 기간 키. date/datetime·Excel 일련번호(COM Value2)·'YYYY-06-30 (반기)'·ISO 문자열."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, datetime):
        d = v.date()
    elif isinstance(v, date):
        d = v
    elif isinstance(v, (int, float)):
        if not 1 <= v <= 73050:                       # 1900~2099 일련번호만
            return None
        d = (_EXCEL_EPOCH + timedelta(days=float(v))).date()
    elif isinstance(v, str):
        m = re.match(r"^\s*(\d{4})-(\d{2})-(\d{2})\s*(\(반기\))?\s*$", v)
        if not m:
            return None
        y, mo, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if m.group(4) or (mo, dd) == (6, 30):
            return f"{y}H1"
        return f"FY{y}" if (mo, dd) == (12, 31) else None
    else:
        return None
    if (d.month, d.day) == (12, 31):
        return f"FY{d.year}"
    if (d.month, d.day) == (6, 30):
        return f"{d.year}H1"
    return None


def _at(rows, i, c):
    """rows[i] 의 c열(1부터) 값. 범위 밖은 None."""
    if 0 <= i < len(rows):
        row = rows[i] or []
        if 0 <= c - 1 < len(row):
            return row[c - 1]
    return None


def _pk_text(v):
    return v.strip() if isinstance(v, str) and _PK.match(v.strip()) else None


def _header_periods(rows, i):
    """마커 행 i(0부터) 블록의 기간 키(열 순서). raw 블록 = 단위 행 다음 헤더 2행째(날짜·'(반기)'),
    mFS 구역 = 마커 다음 행의 기간 키 문자열, rN_기타 원표(단위 행 E열에 '주석' 태그) = 없음."""
    d1, e1 = _at(rows, i + 1, RAW_LABEL_COL), _at(rows, i + 1, RAW_FIRST_VALUE_COL)
    if isinstance(d1, str) and d1.lstrip().startswith("(단위"):
        if isinstance(e1, str) and "주석" in e1:
            return []
        src, conv = i + 3, _period_of
    elif _pk_text(e1):
        src, conv = i + 1, _pk_text
    else:
        return []
    out, c = [], RAW_FIRST_VALUE_COL
    while True:
        pk = conv(_at(rows, src, c))
        if pk is None:
            return out
        out.append(pk)
        c += 1


def existing_from_rows(rows):
    """시트 값 2-D(A1부터) → {"blocks": {제목: {"start": 마커 행(1부터), "labels": [...]}}, "periods": [기간 키...]}.
    C열 § 표식 + D열 제목이 블록 시작. 라벨은 마커+2행 이후 비어 있지 않은 D 값('(단위' 행 제외)."""
    blocks, periods, cur, start = {}, [], None, -1
    for i in range(len(rows)):
        c, d = _at(rows, i, BLOCK_MARK_COL), _at(rows, i, RAW_LABEL_COL)
        if c == BLOCK_MARK:
            if d not in (None, ""):
                cur = str(d)
                blocks[cur] = {"start": i + 1, "labels": []}
                start = i + 2
                for pk in _header_periods(rows, i):
                    if pk not in periods:
                        periods.append(pk)
            else:
                cur = None
            continue
        if cur is None or i < start or d in (None, ""):
            continue
        labs = blocks[cur]["labels"]
        if not labs and str(d).lstrip().startswith("(단위"):
            continue
        labs.append(str(d))
    return {"blocks": blocks, "periods": periods}


def existing_from_model(model):
    """WorkbookModel 에서 업데이트용 기존 구조(write_com.read_existing 과 같은 모양)를 만든다(테스트·비COM용)."""
    return {s.name: existing_from_rows(s.values_matrix()) for s in model.sheets if is_struct_sheet(s.name)}


def _norm_existing(existing):
    """새 모양({"blocks", "periods"})과 예전 모양({제목: [라벨]}, 시작 행·기간 없음)을 새 모양으로."""
    out = {}
    for name, v in (existing or {}).items():
        if isinstance(v, dict) and isinstance(v.get("blocks"), dict):
            out[name] = {"blocks": {t: {"start": (b or {}).get("start"), "labels": list((b or {}).get("labels") or [])}
                                    for t, b in v["blocks"].items()},
                         "periods": list(v.get("periods") or [])}
        else:
            out[name] = {"blocks": {t: {"start": None, "labels": list(lab or [])} for t, lab in (v or {}).items()},
                         "periods": []}
    return out


def _merge_periods(existing_periods, periods):
    """기존 열 순서 그대로 + 새 기간은 오른쪽에. 이번에 없는 기존 기간도 열을 유지(값은 빈칸)."""
    return list(existing_periods) + [p for p in periods if p not in existing_periods]


def _resolve_titles(items, ex_blocks):
    """items: [(기본 제목, 행 라벨들)] → 시트 안에서 유일한 블록 제목.
    기존 블록 제목과 정확히 같은 것을 먼저 준다. 기본 제목이 같은 블록이 여럿이면 기존 '제목'·'제목 #n' 을
    행 라벨이 가장 많이 겹치는 블록에 짝짓고(동률이면 앞 블록), 짝이 없는 새 중복에만 다음 #n 을 붙인다."""
    groups = OrderedDict()
    for j, (base, _) in enumerate(items):
        groups.setdefault(base, []).append(j)

    def num(t, base):
        if t == base:
            return 1
        m = _TITLE_NO.match(t)
        return int(m.group(2)) if m and m.group(1) == base else None

    cands = {base: sorted((t for t in ex_blocks if num(t, base) is not None), key=lambda t: num(t, base))
             for base in groups}
    used = {t for lst in cands.values() for t in lst}
    out = [None] * len(items)
    norms = [{N.norm_label(x) for x in labs} for _, labs in items]
    for base, idxs in groups.items():
        free = list(idxs)
        for t in cands[base]:
            if not free:
                break
            exl = {N.norm_label(x) for x in ex_blocks[t]["labels"]}
            best = max(free, key=lambda j: (len(exl & norms[j]), -free.index(j)))
            out[best] = t
            free.remove(best)
        for j in free:
            out[j] = _unique(base, used)
    return out


def _layout(sheet_name, entries, ex_blocks, warnings, shifts, r0):
    """entries: [(제목, 높이)] → 시작 행 목록.
    기존 블록은 기록된 시작 행에 둔다(앞 블록이 커져도 여유 행 RAW_GAP_ROWS 개 안이면 그대로).
    앞 블록이 여유 행을 다 쓰고 넘치면 그 끝 + RAW_GAP_ROWS 로 밀고 경고. 새 블록은 마지막 블록 뒤에."""
    starts = [None] * len(entries)
    known = sorted((j for j, (t, _) in enumerate(entries) if (ex_blocks.get(t) or {}).get("start")),
                   key=lambda j: ex_blocks[entries[j][0]]["start"])
    end = None                                   # 직전 블록 마지막 행 다음 행
    for j in known:
        t, h = entries[j]
        s = ex_blocks[t]["start"]
        free = r0 if end is None else end
        if s >= free:
            starts[j] = s
        else:
            starts[j] = r0 if end is None else end + RAW_GAP_ROWS
            msg = f"{sheet_name} '{t}': 블록이 {starts[j] - s}행 밀림 — 모델 참조 확인 필요"
            warnings.append(msg)
            shifts.append(msg)
        end = starts[j] + h
    titles = {t for t, _ in entries}
    gone_end = max((b["start"] + 4 + len(b.get("labels") or []) for t, b in ex_blocks.items()
                    if t not in titles and b.get("start")), default=None)   # 사라진 블록 영역 끝 — 새 블록이 침범 금지
    if gone_end is not None:
        end = gone_end if end is None else max(end, gone_end)
    for j, (t, h) in enumerate(entries):
        if starts[j] is None:
            starts[j] = r0 if end is None else end + RAW_GAP_ROWS
            end = starts[j] + h
    for t, b in ex_blocks.items():
        if t not in titles and b.get("start"):
            warnings.append(f"{sheet_name} '{t}': 이번 수집에 없는 블록 — 그 자리는 비워 둠")
    return starts


# ---------------------------------------------------------------- 블록 쓰기

def _period_header(col, pk):
    p = next((x for x in col.plan.periods if x.key == pk), None)
    if p is None:                    # 기존 파일에만 있는 기간(열 유지, 값 빈칸)
        m = _PK.match(pk or "")
        if not m:
            raise ValueError(f"기간 {pk!r} 이(가) 수집 계획(plan.periods)에 없습니다: {[x.key for x in col.plan.periods]}")
        p = Period.annual(int(m.group(1))) if m.group(1) else Period.half(int(m.group(2)))
    fiscal = col.units.get(("fiscal", pk))
    half = p.kind == "half"
    label = (f"제 {fiscal} 기" + (" 반기" if half else "")) if fiscal else ("반기" if half else "")
    return label, p.end, half


def _write_block(sheet, r, block, col, section):
    """블록을 r행부터 쓴다: 제목(section)·단위·헤더 2행·데이터 행(높이 4 + 행 수). [Placed] 반환."""
    sheet.set(r, BLOCK_MARK_COL, BLOCK_MARK, "text")
    sheet.set(r, RAW_LABEL_COL, section, "section")
    sheet.set(r + 1, RAW_LABEL_COL, f"(단위 : {_unit_txt(block.unit)})", "unit")
    sheet.set(r + 2, RAW_LABEL_COL, "", "hdr")
    sheet.set(r + 3, RAW_LABEL_COL, "", "hdr")
    for i, pk in enumerate(block.periods):
        label, end, half = _period_header(col, pk)
        sheet.set(r + 2, RAW_FIRST_VALUE_COL + i, label, "hdr")
        if half:
            sheet.set(r + 3, RAW_FIRST_VALUE_COL + i, f"{end.isoformat()} (반기)", "hdr")
        else:
            sheet.set(r + 3, RAW_FIRST_VALUE_COL + i, end, "date")
    last = RAW_FIRST_VALUE_COL + max(len(block.periods), 1) - 1
    sheet.rule(r + 2, RAW_LABEL_COL, last, "top")          # 머리글 2행 = 한 덩어리: 위 실선 + 아래 실선
    sheet.rule(r + 3, RAW_LABEL_COL, last, "bottom")
    rr, seen, placed = r + 4, {}, []
    flag_col = RAW_FIRST_VALUE_COL + len(block.periods) + 1
    for row in block.rows:
        total = is_total(row.norm)
        sheet.set(rr, RAW_LABEL_COL, row.label, "bold" if total else "text")
        for i, pk in enumerate(block.periods):
            v = row.values.get(pk)
            if isinstance(v, str) and N.norm_label(v) == row.norm:      # 소제목 행(값 = 라벨) → 라벨만
                v = None
            if v is not None:
                num = isinstance(v, (int, float))
                sheet.set(rr, RAW_FIRST_VALUE_COL + i, v, ("bold" if total else "num") if num else "text")
        if total:
            sheet.rule(rr, RAW_LABEL_COL, last, "bottom")
        if row.flags:
            sheet.set(rr, flag_col, ", ".join(row.flags), "flag")
        k = seen.get(row.norm, 0)
        seen[row.norm] = k + 1
        placed.append(Placed(sheet.name, block, section, row, rr, k))
        rr += 1
    return placed


class _RawTable:
    """기간 표시가 없어 펴지 못한 주석 원표 하나(rN_기타 에 연도별 그대로). 높이 = 제목·단위·헤더 + 본문 행."""
    def __init__(self, key, recs):
        self.key, self.recs = key, recs
        self.unit_known = all(getattr(x, "unit_known", True) for x in recs)
        self.labels = [x.row_label for x in recs if x.raw_cell[1] == 0]
        self.height = max(3, 4 + max((x.raw_cell[0] for x in recs), default=-1))


def _write_raw_table(s, r, t, section):
    x0 = t.recs[0]
    s.set(r, BLOCK_MARK_COL, BLOCK_MARK, "text")
    s.set(r, RAW_LABEL_COL, section, "section")
    s.set(r + 1, RAW_LABEL_COL, f"(단위 : {_unit_txt(x0.unit)})" if t.unit_known else UNIT_UNKNOWN, "unit")
    s.set(r + 1, RAW_FIRST_VALUE_COL, f"{x0.period_key} · 주석 {x0.note_no} · rcept_no {x0.rcept_no}", "italic")
    hdr = {}
    for x in t.recs:
        hdr.setdefault(x.raw_cell[1], x.col_label)
    last = RAW_LABEL_COL + max(hdr, default=0)
    for c in range(RAW_LABEL_COL, last + 1):
        s.set(r + 2, c, hdr.get(c - RAW_LABEL_COL) or "", "hdr")       # 한 줄 머리글: 위·아래 실선(hdr 기본)
    totals = {x.raw_cell[0] for x in t.recs if x.raw_cell[1] == 0 and is_total(N.norm_label(str(x.value or x.row_label)))}
    for x in t.recs:
        if x.value not in (None, ""):
            num = isinstance(x.value, (int, float))
            bold = x.raw_cell[0] in totals
            s.set(r + 3 + x.raw_cell[0], RAW_LABEL_COL + x.raw_cell[1], x.value,
                  ("bold" if bold else "num") if num else ("bold" if bold else "text"))
    for ri in totals:
        s.rule(r + 3 + ri, RAW_LABEL_COL, last, "bottom")


def _place_raw(sheet, items, col, ex_sheet, periods, warnings, shifts, r0=1):
    """items: [(기본 제목, Block 또는 _RawTable)] 을 설계 §5 안정성 규칙으로 배치해 쓴다.
    Block 은 기존 행 순서를 적용하고 열을 시트 공통 기간(periods)으로 맞춘다. ([Placed], {원표 키: 제목})."""
    ex_blocks = ex_sheet["blocks"]
    titles = _resolve_titles([(base, obj.labels if isinstance(obj, _RawTable) else [r.label for r in obj.rows])
                              for base, obj in items], ex_blocks)
    ready = []
    for (_, obj), t in zip(items, titles):
        if isinstance(obj, _RawTable):
            ready.append((t, obj, obj.height))
            continue
        ex = ex_blocks.get(t)
        b = L.apply_existing_order(obj, ex["labels"]) if ex and ex["labels"] else obj
        b = dataclasses.replace(b, periods=list(periods))
        ready.append((t, b, 4 + len(b.rows)))
    starts = _layout(sheet.name, [(t, h) for t, _, h in ready], ex_blocks, warnings, shifts, r0)
    placed, secs = [], {}
    for (t, obj, _), r in zip(ready, starts):
        if isinstance(obj, _RawTable):
            _write_raw_table(sheet, r, obj, t)
            secs[obj.key] = t
        else:
            placed.extend(_write_block(sheet, r, obj, col, t))
    return placed, secs


# ---------------------------------------------------------------- 시트별

CHECK_NOTE = ("주석 표는 항목별 천원 반올림으로 공시되어 재무제표(원)와 ±1~2천원 차이가 날 수 있습니다. "
              "검증 시트의 FALSE(파란 글자)는 차이(원) 열로 원인을 확인하세요 — 허용폭 없이 그대로 표시합니다.")


FILE_CODE_TEXT = "파일 입력"                         # 설정 corp_code 행(파일 입력일 때)
FILES_LABEL = "입력 파일"
FILES_NOTE = "경로는 설정 JSON 에 저장됨 — [업데이트] 때 다시 읽음"


def _settings_sheet(col, python_path, script_path, unit=None):
    s = Sheet("설정", is_data=False)
    cfg = col.cfg
    uval, uauto, utext = unit or display_unit(col)
    js = cfg.to_dict()
    if uauto:
        js["unit_judged"] = uval                    # 업데이트 때 셀이 이 값 그대로면 자동(다시 판정), 다르면 직접 선택
    file_src = getattr(cfg, "source", "dart") == "file"
    rows = [("회사", cfg.corp_name, None), ("corp_code", FILE_CODE_TEXT if file_src else cfg.corp_code, None),
            ("기준일", cfg.as_of, None),
            ("연도 범위", f"{cfg.start_year}~{cfg.end_year}", None),
            ("반기 포함", "예" if cfg.include_half else "아니오", None),
            ("연결/별도", ", ".join(_fs_name(d) for d in cfg.fs_divs), None)]
    if file_src:                                    # 파일 입력만 — refresh 는 라벨로 행을 찾으므로 선택 행이어도 된다
        rows.append((FILES_LABEL, "\n".join(str(d.get("path") or "") for d in cfg.files or []), None))
    rows += [("mFS 표시 단위", uval, UNIT_NAME), ("단위 판정", utext, None), ("검증 안내", CHECK_NOTE, None),
            ("프리셋", cfg.preset, None),
            ("파이썬 경로", python_path, "cfg_python"), ("refresh.py 경로", script_path, "cfg_script"),
            ("마지막 갱신", date.today(), None), ("상태", "생성 완료", "cfg_status"),
            ("설정 JSON", json.dumps(js, ensure_ascii=False), "cfg_json")]
    s.set(1, 2, "dart_fs — DART 재무제표·주석 raw 데이터", "title")
    s.set(2, 2, "[업데이트] 버튼: 설정·목차·매핑표를 읽어 데이터 시트(mFS·검증·rFS_*·rN_*·데이터·로그)를 다시 씁니다. "
                "그 시트에 손으로 적은 내용은 덮어씌워집니다. 모델 시트는 RAW >> 앞에 두세요.", "italic")
    for i, (k, v, name) in enumerate(rows):
        r = 4 + i
        s.set(r, 2, k, "bold")
        s.set(r, 3, v, "date" if isinstance(v, date) else ("italic" if k == "검증 안내" else None))
        if k == FILES_LABEL:
            s.cells[(r, 3)].fmt = "wrap"            # 경로 한 줄에 하나
            s.set(r, 4, FILES_NOTE, "italic")
        if name:
            s.names[name] = f"C{r}"
        if name == UNIT_NAME:                       # 목록(1·1000·1000000) + 오른쪽에 단위 글자
            s.validations.append((f"C{r}", UNIT_LIST_F1))
            s.cells[(r, 3)].fmt = "input"           # 입력 칸: 연노랑 채움 + 점선 테두리
            s.set(r, 4, "=" + UNIT_TEXT_EXPR, "bold", formula=True)
            s.set(r, 5, "← 1·1000·1000000 중 고르면 mFS 가 바로 바뀝니다([업데이트] 불필요)", "italic")
    s.col_widths = {2: 16, 3: 90}
    return s


def _toc_sheet(col):
    s = Sheet("목차")
    hdr = ["번호", "제목", "표준 주제", "포함", "추천", "배치 시트", "표 수", "정규화 제목"]
    for c, h in enumerate(hdr, 1):
        s.set(1, c, h, "hdr")
    for i, e in enumerate(col.toc, 2):
        sheet = (sheet_for(e.topic, col.mapping) if e.topic else None) or ETC_SHEET
        for c, v in enumerate([e.no, e.title, e.topic or "", bool(e.include), "★" if e.recommended else "",
                               sheet, e.table_count, e.title_norm], 1):
            s.set(i, c, v, "check" if c == 4 else None)
    s.col_widths = {1: 6, 2: 40, 3: 14, 4: 8, 5: 6, 6: 16, 7: 8, 8: 40}
    s.freeze = "A2"
    return s


def _block_base(title, b):
    return f"{title} — {b.col_label}" if b.col_label else title


def _raw_widths(s):
    """_RAW_WIDTHS + 값 열(RAW_FIRST_VALUE_COL 이후 숫자·머리글·날짜 셀이 있는 열)은 VALUE_COL_WIDTH.
    표식(flag)·안내(italic)만 있는 열은 기본 폭 그대로."""
    w = dict(_RAW_WIDTHS)
    for (_, c), cell in s.cells.items():
        if c >= RAW_FIRST_VALUE_COL and (cell.fmt in _VALUE_FMTS or
                                         (isinstance(cell.v, (int, float)) and not isinstance(cell.v, bool))):
            w[c] = VALUE_COL_WIDTH
    return w


def _raw_fs_sheets(col, existing, warnings, shifts):
    """rFS_연결·rFS_별도. (시트 목록, {fs_div: [Placed]})."""
    sheets, placed = [], {}
    for fs_div in _fs_divs(col):
        recs = [r for r in col.fs_records if r.fs_div == fs_div]
        if not recs:
            continue
        s = Sheet(f"rFS_{_fs_name(fs_div)}")
        ex = existing.get(s.name, _EMPTY)
        blocks = L.fs_blocks(recs, col.periods)
        for b in blocks:
            warnings.extend(b.meta.get("warnings", []))
        placed[fs_div], _ = _place_raw(s, [(_block_base(b.title, b), b) for b in blocks], col, ex,
                                       _merge_periods(ex["periods"], col.periods), warnings, shifts)
        s.col_widths = _raw_widths(s)
        sheets.append(s)
    return sheets, placed


def _note_title(b, multi_fs):
    t = f"주석 {b.note_no}. {b.note_title} | {b.title}"
    return f"[{_fs_name(b.fs_div)}] {t}" if multi_fs else t


def _note_sheet_name(topic, mapping):
    return (sheet_for(topic, mapping) if topic else None) or ETC_SHEET


def _note_key(b, fs_order):
    no = int(b.note_no) if str(b.note_no).isdigit() else 999
    return (fs_order.get(b.fs_div, 9), no)


def _raw_note_sheets(col, existing, warnings, shifts):
    """rN_* (표준 주제별) + rN_기타. 블록 순서는 주석 번호 순, 같은 주석 안에서는 공시 표·열 순서 유지.
    (시트 목록, [Placed], {원표 키: 제목})."""
    annual = [p for p in col.periods if p.startswith("FY")]
    folded = [r for r in col.note_records if r.raw_cell is None]
    raw = [r for r in col.note_records if r.raw_cell is not None]
    fs_order = {d: i for i, d in enumerate(_fs_divs(col))}
    multi = len(col.cfg.fs_divs) > 1
    by_sheet = OrderedDict()
    for b in sorted(L.note_blocks(folded, annual), key=lambda b: _note_key(b, fs_order)):
        warnings.extend(b.meta.get("warnings", []))
        by_sheet.setdefault(_note_sheet_name(b.topic, col.mapping), []).append(b)
    order = list(OrderedDict.fromkeys([t.sheet for t in col.mapping.topics if t.sheet != ETC_SHEET]))
    sheets, placed = [], []
    for name in order:
        if name not in by_sheet:
            continue
        s = Sheet(name)
        ex = existing.get(name, _EMPTY)
        got, _ = _place_raw(s, [(_block_base(_note_title(b, multi), b), b) for b in by_sheet[name]], col, ex,
                            _merge_periods(ex["periods"], annual), warnings, shifts)
        placed.extend(got)
        s.col_widths = _raw_widths(s)
        sheets.append(s)
    # rN_기타: 표준 주제가 없는 블록 + 기간 표시 없는 원표(연도별 그대로, 표마다 기간·주석·rcept_no 태그).
    # 시트 구성이 매번 같도록 항상 만든다.
    s = Sheet(ETC_SHEET)
    ex = existing.get(ETC_SHEET, _EMPTY)
    items = [(_block_base(_note_title(b, multi), b), b) for b in by_sheet.get(ETC_SHEET, [])]
    items += _raw_table_items(raw, fs_order, multi)
    got, raw_secs = _place_raw(s, items, col, ex, _merge_periods(ex["periods"], annual), warnings, shifts)
    placed.extend(got)
    if not s.cells:
        s.set(1, RAW_LABEL_COL, "해당 표 없음", "italic")
    s.col_widths = _raw_widths(s)
    sheets.append(s)
    return sheets, placed, raw_secs


def _raw_table_items(raw, fs_order, multi):
    """기간 표시가 없어 펴지 못한 주석 표 → [(기본 제목, _RawTable)] (연결→별도, 주석 번호, 표, 기간 순)."""
    tables = OrderedDict()
    for x in raw:
        tables.setdefault((x.fs_div, x.period_key, x.note_no, x.table_idx), []).append(x)

    def key(k):
        no = int(k[2]) if str(k[2]).isdigit() else 999
        return (fs_order.get(k[0], 9), no, k[3], k[1])

    out = []
    for k in sorted(tables, key=key):
        x0 = tables[k][0]
        title = f"주석 {x0.note_no}. {x0.note_title} | {x0.caption} [{x0.period_key}]"
        if multi:
            title = f"[{_fs_name(x0.fs_div)}] {title}"
        out.append((title, _RawTable(k, tables[k])))
    return out


# ---------------------------------------------------------------- mFS

def _std_groups(col, fs_div, placed):
    """raw 행 → 표준계정 묶음. {sj: [(std|None, [Placed])]} — std 는 매핑표 순서, 미매핑은 raw 순서로 뒤에."""
    m = col.mapping
    mapped, unmapped = OrderedDict(), OrderedDict()
    for p in placed:
        sjd = p.block.sj_div
        if sjd not in ("BS", "IS", "CIS", "CF"):
            continue
        sj = "IS" if sjd == "CIS" else sjd
        std, _ = m.map_account(p.row.label, p.row.account_id, sjd)
        if std:
            mapped.setdefault((sj, std), []).append(p)
        elif any(v is not None for v in p.row.values.values()):
            unmapped.setdefault(sj, []).append(p)       # 값이 하나도 없는 미매핑 행(제목 행 '자 산' 등)은 mFS 에서 뺀다
    out = {}
    for (sj, std), ps in mapped.items():
        if any(p.block.sj_div == "IS" for p in ps):      # 손익계산서와 포괄손익계산서 둘 다 있으면 손익계산서만
            ps = [p for p in ps if p.block.sj_div != "CIS"]
        out.setdefault(sj, []).append((std, ps))
    for sj in out:
        out[sj].sort(key=lambda g: m.std_order.get((sj, g[0]), 9999))
    for sj, ps in unmapped.items():
        out.setdefault(sj, []).extend((None, [p]) for p in ps)
    return out


def _sum_formula(ps, pk, warnings=None):
    """raw 셀 직참조 합을 원으로 바꿔 cfg_unit 으로 나눈다: `='rFS_별도'!E21*1000/cfg_unit` (공시 단위 1이면 *1 생략).
    값 있는 raw 셀이 없으면 None. 같은 기간에 정규화 라벨과 값이 모두 같은 행('비지배지분' 과 'Ⅰ. 비지배지분' — 머리글과
    항목이 되풀이된 원문)은 첫 행만 더하고 warnings 에 적는다. 값이나 라벨이 다르면 그대로 더한다."""
    by_unit, seen = OrderedDict(), set()
    for p in ps:
        if pk in p.block.periods and p.row.values.get(pk) is not None:
            key = (N.norm_label(p.row.label), p.row.values.get(pk), p.block.unit or 1)
            if key in seen:
                if warnings is not None:
                    warnings.append(f"중복 행 1회만 합산: {p.row.label} {pk}")
                continue
            seen.add(key)
            c = RAW_FIRST_VALUE_COL + p.block.periods.index(pk)
            by_unit.setdefault(p.block.unit or 1, []).append(_ref(p.sheet, p.rownum, c))
    if not by_unit:
        return None
    terms = []
    for u, refs in by_unit.items():
        t = "+".join(refs)
        if u != 1:
            t = (f"({t})" if len(refs) > 1 else t) + f"*{u}"
        terms.append(t)
    body = "+".join(terms)
    if len(terms) > 1 or len(by_unit.get(1, ())) > 1:          # 'a+b/cfg_unit' 이 되지 않게
        body = f"({body})"
    return f"={body}/{UNIT_NAME}"


def _mfs_rows(groups, ex_labels):
    """mFS 한 구역의 행 [(라벨, std|None, [Placed]|None)].
    기존 행 순서를 먼저(같은 라벨은 등장 순서대로 짝), 새 행은 구역 끝에. 기존에 있었는데 이번에 없는 행은
    라벨만 남긴다([Placed] 자리 None — 값 빈칸 + '공시에서 사라짐')."""
    items = [(std or ps[0].row.label, std, ps) for std, ps in groups]
    if not ex_labels:
        return items
    pool = OrderedDict()
    for it in items:
        pool.setdefault(it[0], []).append(it)
    out, used = [], set()
    for lab in ex_labels:
        lst = pool.get(lab)
        if lst:
            it = lst.pop(0)
            out.append(it)
            used.add(id(it))
        else:
            out.append((lab, None, None))
    out.extend(it for it in items if id(it) not in used)
    return out


def _mfs_sheet(col, fs_placed, ex_sheet, warnings, shifts):
    """표준계정 × 기간 수식 시트. 구역(연결/별도 × BS·IS·CF·검증 요약)마다 C열 § 표식을 두고, 업데이트 때는
    raw 시트와 같은 안정성 규칙(기존 구역 시작 행·행 순서·열 순서 유지)을 따른다.
    반환 (시트, ctx): ctx = {"periods", "rows_of", "filled", "summary"} — 검증 시트(excel/checks)가
    mFS 셀을 참조하는 데 쓰고, summary [(fs_div, 첫 요약 행)] 은 _mfs_summary 가 COUNTIF 를 채우는 자리."""
    s = Sheet("mFS")
    periods = _merge_periods(ex_sheet["periods"], col.periods)
    ex_blocks = ex_sheet["blocks"]
    s.set(1, RAW_LABEL_COL, f'="표준계정 (표시 단위: "&{UNIT_TEXT_EXPR}&", raw 시트 직참조 · 설정 시트 {UNIT_NAME} 로 전환)"',
          "title", formula=True)
    entries = []                                       # (제목, 높이, 내용)
    for fs_div in _fs_divs(col):
        placed = fs_placed.get(fs_div)
        if not placed:
            continue
        groups = _std_groups(col, fs_div, placed)
        for sj in ("BS", "IS", "CF"):
            t = f"{_fs_name(fs_div)} {_SJ_TITLE[sj]}"
            rows = _mfs_rows(groups.get(sj, []), (ex_blocks.get(t) or {}).get("labels"))
            entries.append((t, 2 + len(rows), ("data", fs_div, sj, rows)))
        entries.append((f"{_fs_name(fs_div)} 검증 요약", 2 + len(_SUMMARY_ROWS), ("check", fs_div)))
    starts = _layout("mFS", [(t, h) for t, h, _ in entries], ex_blocks, warnings, shifts, 3)
    flag_col = RAW_FIRST_VALUE_COL + len(periods) + 1
    last = RAW_FIRST_VALUE_COL + max(len(periods), 1) - 1

    def header(r):                                     # 구역 머리글 1행(기간 키): 위·아래 실선
        s.set(r, RAW_LABEL_COL, "", "hdr")
        for i, pk in enumerate(periods):
            s.set(r, RAW_FIRST_VALUE_COL + i, pk, "hdr")
        s.rule(r, RAW_LABEL_COL, last, "both")
    rows_of = {}                                       # (fs_div, sj, std) -> mFS 행
    filled = set()                                     # (fs_div, sj, std, i): 수식이 실제로 쓰인 칸
    for (t, _, spec), r in zip(entries, starts):       # 1) 표준계정 구역
        if spec[0] != "data":
            continue
        _, fs_div, sj, rows = spec
        s.set(r, BLOCK_MARK_COL, BLOCK_MARK, "text")
        s.set(r, RAW_LABEL_COL, t, "section")
        header(r + 1)
        rr = r + 2
        for lab, std, ps in rows:
            total = bool(std) and is_total(N.norm_label(lab))
            if ps is None:
                s.set(rr, RAW_LABEL_COL, lab, "flag")
                s.set(rr, flag_col, "공시에서 사라짐", "flag")
            else:
                s.set(rr, RAW_LABEL_COL, lab, ("bold" if total else "text") if std else "flag")
                if std is None:
                    s.set(rr, flag_col, "미매핑", "flag")
                for i, pk in enumerate(periods):
                    f = _sum_formula(ps, pk, warnings)
                    if f:
                        s.set(rr, RAW_FIRST_VALUE_COL + i, f, "bold" if total else "num", formula=True)
                        if std:
                            filled.add((fs_div, sj, std, i))
                if std:
                    rows_of[(fs_div, sj, std)] = rr
                if total:
                    s.rule(rr, RAW_LABEL_COL, last, "bottom")
            rr += 1
    summary = []
    for (t, _, spec), r in zip(entries, starts):       # 2) 검증 요약 구역(수식은 검증 시트가 만들어진 뒤 _mfs_summary)
        if spec[0] != "check":
            continue
        s.set(r, BLOCK_MARK_COL, BLOCK_MARK, "text")
        s.set(r, RAW_LABEL_COL, t, "section")
        header(r + 1)
        for k, (label, _) in enumerate(_SUMMARY_ROWS):
            s.set(r + 2 + k, RAW_LABEL_COL, label, "text")
        s.set(r + 2, flag_col, "자세한 내용은 '검증' 시트", "italic")
        summary.append((spec[1], r + 2))
    s.col_widths = _raw_widths(s)
    return s, {"periods": periods, "rows_of": rows_of, "filled": filled, "summary": summary}


def _mfs_summary(s, ctx, ranges):
    """mFS 요약 3행 × 기간: 검증 시트의 그 fs_div 구역(행 범위)·같은 기간 열을 COUNTIF 로 센다.
    `=COUNTIF(` 로 시작하므로 finalize.verify 의 검증 셀 집계(=ROUND( 접두)에 걸리지 않는다."""
    for fs_div, r0 in ctx["summary"]:
        rng = ranges.get(fs_div)
        if not rng or rng[0] is None:
            continue
        r1, r2 = rng
        for k, (_, crit) in enumerate(_SUMMARY_ROWS):
            for i in range(len(ctx["periods"])):
                c = _col_letter(RAW_FIRST_VALUE_COL + i)
                s.set(r0 + k, RAW_FIRST_VALUE_COL + i, f"=COUNTIF('검증'!{c}{r1}:{c}{r2},{crit})", "num", formula=True)


# ---------------------------------------------------------------- 롱포맷·매핑표·로그

def _data_sheet(col, fs_placed, note_placed, raw_secs):
    s = Sheet("데이터")
    hdr = ["시트", "블록", "열라벨", "행라벨", "기간", "값", "단위", "출처 rcept_no", "출처", "표준계정"]
    for c, h in enumerate(hdr, 1):
        s.set(1, c, h, "hdr")
    rows = []
    fs_sec = {}
    for ps in fs_placed.values():
        for p in ps:
            cl = N.norm_label(p.block.col_label) if p.block.col_label else None    # SCE 열은 정규화 키(layout.fs_blocks)
            fs_sec.setdefault((p.block.fs_div, p.block.sj_div, cl), p.section)
    note_secs = OrderedDict()                  # (fs_div, 주석, 열, 서명) -> [section...] 배치 순서
    for p in note_placed:
        b = p.block
        cl = N.norm_label(b.col_label) if b.col_label else None
        # 행 겹침으로 병합된 블록은 옛 캡션 서명의 레코드도 같은 블록으로 찾게 한다(layout.note_blocks meta['sigs'])
        sigs = [L.caption_sig(b.title)] + [re.sub(r"#\d+$", "", g) for g in b.meta.get("sigs", [])]
        for sig in dict.fromkeys(sigs):
            key = (b.fs_div, b.note_no, cl, sig)
            if p.section not in note_secs.setdefault(key, []):
                note_secs[key].append(p.section)
    tabs = {}                                  # layout.note_blocks 와 같은 #n 순번
    for x in col.note_records:
        if x.raw_cell is None:
            lst = tabs.setdefault((x.fs_div, x.note_no, x.period_key, L.caption_sig(x.caption)), [])
            if x.table_idx not in lst:
                lst.append(x.table_idx)
    for x in col.fs_records:
        std, _ = col.mapping.map_account(x.label, x.account_id, x.sj_div)
        rows.append([f"rFS_{_fs_name(x.fs_div)}",
                     fs_sec.get((x.fs_div, x.sj_div, L.sce_detail_key(x)), _SJ_TITLE.get(x.sj_div, x.sj_div)), x.detail or "", x.label,
                     x.period_key, x.amount, x.unit, x.rcept_no, x.source, std or ""])
    for x in col.note_records:
        sheet = ETC_SHEET if x.raw_cell is not None else _note_sheet_name(x.topic, col.mapping)
        if x.raw_cell is not None:
            sec = raw_secs.get((x.fs_div, x.period_key, x.note_no, x.table_idx))
        else:
            sig = L.caption_sig(x.caption)
            lst = note_secs.get((x.fs_div, x.note_no, N.norm_label(x.col_label) if x.col_label else None, sig), [])
            n = tabs[(x.fs_div, x.note_no, x.period_key, sig)].index(x.table_idx)
            sec = lst[min(n, len(lst) - 1)] if lst else None
        rows.append([sheet, sec or f"주석 {x.note_no}. {x.note_title} | {x.caption}", x.col_label or "", x.row_label,
                     x.period_key, x.value, x.unit, x.rcept_no, "XML", ""])
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            if v is not None:
                s.set(r, c, v, "num" if c == 6 and isinstance(v, (int, float)) else None)
    s.freeze = "A2"
    s.col_widths = {1: 12, 2: 40, 3: 14, 4: 36, 5: 9, 6: 18}
    return s


def _mapping_sheet(col):
    s = Sheet("매핑표")
    for r, row in enumerate(col.mapping.rows_for_sheet(), 1):
        for c, v in enumerate(row, 1):
            s.set(r, c, v, "hdr" if r == 1 else None)
    s.freeze = "A2"
    s.col_widths = {1: 10, 2: 5, 3: 44, 4: 22, 5: 10, 6: 14, 7: 6}
    return s


def _log_sheet(col, warnings):
    s = Sheet("로그")
    r = 1

    def section(title, hdr):
        nonlocal r
        s.set(r, 1, title, "section"); r += 1
        for c, h in enumerate(hdr, 1):
            s.set(r, c, h, "hdr")
        r += 1

    reps = sorted(col.plan.reports.values(), key=lambda x: (x.period_end, x.rcept_no))
    file_src = bool(reps) and all(rep.rtype == "파일" for rep in reps)
    if file_src:                                    # 파일 입력: 접수번호·정정 대신 시트(표)·수정일·단위
        meta = getattr(col.plan, "file_meta", None) or {}
        section("보고서", ["기간말", "파일", "시트(표)", "수정일", "단위"])
        for rep in reps:
            m = meta.get(rep.rcept_no, {})
            for c, v in enumerate([rep.period_end, rep.report_nm, ", ".join(m.get("sheets", [])), rep.rcept_dt,
                                   m.get("unit", "")], 1):
                s.set(r, c, v, "date" if isinstance(v, date) else None)
            r += 1
    else:
        section("보고서", ["기간말", "보고서", "rcept_no", "접수일", "정정", "대체된 접수번호"])
        for rep in reps:
            for c, v in enumerate([rep.period_end, rep.report_nm, rep.rcept_no, rep.rcept_dt,
                                   "정정" if rep.corrected else "", ", ".join(rep.superseded)], 1):
                s.set(r, c, v, "date" if isinstance(v, date) else None)
            r += 1
    r += 1
    section("재작성 차이 (채택 = 다음 연도 보고서 전기 열)", ["기간", "항목", "채택", "원보고서", "채택 출처", "원보고서 출처"])
    for d in col.restate_diffs:
        for c, k in enumerate(["period", "key", "채택", "원보고서", "채택출처", "원보고서출처"], 1):
            v = d.get(k)
            s.set(r, c, v, "num" if c in (3, 4) and isinstance(v, (int, float)) else None)
        r += 1
    r += 1
    if not file_src:                                # 파일 입력엔 감사보고서가 없다
        section("감사인·의견", ["기간말", "보고서", "구분", "rcept_no", "감사인", "의견 문장"])
        for row in col.auditor_rows:
            for c, v in enumerate(row, 1):
                s.set(r, c, v)
            r += 1
        r += 1
    s.set(r, 1, "메시지", "section"); r += 1
    for msg in list(col.logs) + list(OrderedDict.fromkeys(warnings)):
        s.set(r, 1, msg); r += 1
    s.col_widths = {1: 12, 2: 44, 3: 18, 4: 18, 5: 30, 6: 60}
    return s


def compose(col, existing, python_path, script_path):
    """existing(업데이트 때 기존 파일 구조, write_com.read_existing / existing_from_model):
    {시트명: {"blocks": {블록 제목: {"start": 마커 행, "labels": [행 라벨...]}}, "periods": [기간 키...]}}.
    예전 모양 {시트명: {블록 제목: [행 라벨...]}} 도 받는다(행 순서만 적용).
    블록 밀림 경고는 로그 시트 메시지와 wb.shift_warnings(설정 상태 문구용)에 남는다."""
    ex = _norm_existing(existing)
    wb = WorkbookModel()
    warnings, shifts = [], []
    raw_fs, fs_placed = _raw_fs_sheets(col, ex, warnings, shifts)
    raw_notes, note_placed, raw_secs = _raw_note_sheets(col, ex, warnings, shifts)
    mfs, ctx = _mfs_sheet(col, fs_placed, ex.get("mFS", _EMPTY), warnings, shifts)
    from excel import checks as CK                      # checks 가 compose 를 쓰므로 여기서 들여온다
    chk, ranges = CK.build(col, fs_placed, note_placed, ctx)
    _mfs_summary(mfs, ctx, ranges)
    wb.shift_warnings = shifts
    unit = display_unit(col)
    wb.unit_value, wb.unit_auto, wb.unit_text = unit
    if wb.unit_auto:
        warnings.insert(0, f"표시 단위 {wb.unit_text}")
    wb.add(_settings_sheet(col, python_path, script_path, unit))
    wb.add(_toc_sheet(col))
    wb.add(mfs)
    wb.add(chk)
    sep = Sheet("RAW >>")
    sep.set(2, 2, "이 뒤는 DART 원자료(raw) 시트입니다. 모델 시트는 이 앞에 두세요.", "italic")
    wb.add(sep)
    for s in raw_fs + raw_notes:
        wb.add(s)
    wb.add(_data_sheet(col, fs_placed, note_placed, raw_secs))
    wb.add(_mapping_sheet(col))
    wb.add(_log_sheet(col, warnings))
    return wb
