"""열린 Excel 통합문서(COM)에 WorkbookModel 의 데이터 시트를 다시 쓴다.

- 데이터 시트(compose.is_data_sheet)만 손댄다. 사용자 시트(설정·모델 등)는 쓰지 않는다.
- 값은 시트별 2-D 배열 하나로 Range.Value2, 수식은 셀별 .Formula, 서식은 같은 서식 셀을 묶어 한 번에.
- 문자열은 모두 ' 접두로 써서 Excel 이 숫자·날짜·수식으로 바꾸지 않게 한다(openpyxl 경로와 같은 결과).
- 0x800AC472(사용자 셀 편집 중) 같은 일시 오류는 finalize._retry 로 재시도한다(멱등 동작만).
- ScreenUpdating 은 건드리지 않는다(편집 충돌 시 복원 실패 위험). 계산은 쓰는 동안 수동 → 끝나면 원래 방식으로 복원.
- 이 모듈은 refresh 를 import 하지 않는다(refresh 가 blocks_from_columns 를 여기서 가져간다).
"""
from datetime import date, datetime
from excel import compose as C
from core.normalize import Percent
from excel.write_openpyxl import NUM_FMT, PCT_FMT, border_of, check_ranges
from excel.finalize import _retry, _col_letter

XL_CALC_MANUAL, XL_CALC_AUTOMATIC = -4135, -4105
XL_CENTER = -4108
XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_CONTINUOUS, XL_THIN = 8, 9, 1, 2
FLAG_COLOR = 0x0000C0          # RGB C00000 → COM BGR
CHECK_FALSE_COLOR = 0xFF0000   # RGB 0000FF(파랑) → COM BGR
XL_CELL_VALUE, XL_EQUAL = 1, 3
_ADDR_MAX = 240                # Range("A1,B2,...") 주소 문자열 한도(255) 여유
# 한국어 Excel 일부 설치본은 .NumberFormat 도 로컬 토큰만 받는다([Red]·General 거부, 실측) → 로컬 토큰으로 폴백.
_LOCAL_TOKENS = (("[Red]", "[빨강]"), ("General", "G/표준"))
_STYLE_ORDER = ("title", "section", "bold", "hdr", "italic", "unit", "flag", "date", "check", "pct", "input", "wrap")
XL_VTOP = -4160
INPUT_FILL_COLOR = 0xCCFFFF    # RGB FFFFCC(연노랑) → COM BGR = 13434879 (ColorIndex 19)
XL_EDGE_LEFT, XL_EDGE_RIGHT, XL_DOT = 7, 10, -4118
# COM 으로 'yyyy-mm-dd' 를 쓰면 한국어 Excel 이 시스템 간단한 날짜(기본 서식 14)로 바꿔 저장한다(실측:
# openpyxl 은 'mm-dd-yy' 로 읽고, 표시는 PC 의 날짜 설정을 따름). 구분자를 따옴표 리터럴로 써서 사용자 서식으로 남긴다.
COM_DATE_FMT = 'yyyy"-"mm"-"dd'
_BORDER_EDGES = {"top": (XL_EDGE_TOP,), "bottom": (XL_EDGE_BOTTOM,), "both": (XL_EDGE_TOP, XL_EDGE_BOTTOM)}


def blocks_from_columns(col_c, col_d):
    """C열 마커(§)·D열 값 → {블록 제목: [행 라벨]} (예전 모양). 규칙은 compose.existing_from_rows 와 같다:
    D 가 빈 마커는 블록을 끝내고, 마커+2행 이후에서 '(단위' 로 시작하지 않는 첫 행부터 다음 마커까지
    비어 있지 않은 D 값을 모은다(헤더 행 수가 블록마다 달라 고정 행수로 건너뛰지 않음)."""
    rows = [[None, None, c, d] for c, d in zip(col_c, col_d)]
    return {t: b["labels"] for t, b in C.existing_from_rows(rows)["blocks"].items()}


def _last_row(ws):
    u = ws.UsedRange
    return u.Row + u.Rows.Count - 1


def _last_col(ws):
    u = ws.UsedRange
    return u.Column + u.Columns.Count - 1


def _col_values(ws, col, last_row):
    if last_row < 1:
        return []
    v = ws.Range(ws.Cells(1, col), ws.Cells(last_row, col)).Value2
    if not isinstance(v, tuple):
        return [v]
    return [row[0] if isinstance(row, tuple) else row for row in v]


def sheet_rows(ws):
    """A1 부터 사용 범위 끝까지 2-D 리스트(UsedRange 가 B열부터 시작해도 열 위치가 밀리지 않게)."""
    r, c = _last_row(ws), _last_col(ws)
    v = ws.Range(ws.Cells(1, 1), ws.Cells(r, c)).Value2
    if not isinstance(v, tuple):
        return [[v]]
    return [list(row) if isinstance(row, tuple) else [row] for row in v]


def find_sheet(wb, name):
    for ws in wb.Worksheets:
        if ws.Name == name:
            return ws
    return None


def read_existing(wb):
    """구조 시트(mFS·rFS_*·rN_*)마다 {"blocks": {제목: {"start": 행, "labels": [...]}}, "periods": [...]}.
    블록 헤더 2행째(기말일 일련번호 또는 'YYYY-06-30 (반기)')에서 기간 키를 되살린다(compose.existing_from_rows)."""
    out = {}
    for ws in wb.Worksheets:
        if C.is_struct_sheet(ws.Name):
            out[ws.Name] = C.existing_from_rows(sheet_rows(ws))
    return out


_EXCEL_EPOCH = datetime(1899, 12, 30)


def _com_value(v):
    if v is None or isinstance(v, bool):
        return v
    if isinstance(v, str):
        return ("'" + v) if v else None
    if isinstance(v, date):
        # 날짜는 Excel 일련번호(double)로. pywin32 는 naive datetime 을 현지 시간대 기준으로 옮겨 써서
        # KST 에서 2025-12-31 이 2025-12-30 15:00 이 된다(실측). 서식은 _apply_style('date') 가 입힌다.
        d = v if isinstance(v, datetime) else datetime(v.year, v.month, v.day)
        return (d - _EXCEL_EPOCH).total_seconds() / 86400.0
    if isinstance(v, int):
        return float(v) if abs(v) >= 2 ** 31 else v      # Int64 → double
    return v


def set_number_format(rng, fmt):
    """영문 .NumberFormat 우선, 거부되면 한국어 로컬 토큰으로 .NumberFormatLocal."""
    try:
        rng.NumberFormat = fmt
    except Exception:     # noqa: BLE001
        loc = fmt
        for en, ko in _LOCAL_TOKENS:
            loc = loc.replace(en, ko)
        rng.NumberFormatLocal = loc


def _style_key(cell):
    if isinstance(cell.v, Percent):
        return "pct"
    f = cell.fmt
    if f in _STYLE_ORDER and f != "check":
        return f
    if f == "check" and cell.formula:
        return "check"
    return None


def style_groups(sh):
    """{스타일 키: [(r, c)]} — 같은 서식 셀을 묶어 한 번에 입힌다."""
    out = {}
    for (r, c), cell in sh.cells.items():
        k = _style_key(cell)
        if k:
            out.setdefault(k, []).append((r, c))
    return out


def border_groups(sh):
    """{"top"|"bottom"|"both": [(r, c)]} — openpyxl writer 와 같은 규칙(write_openpyxl.border_of)."""
    out = {}
    for (r, c), cell in sh.cells.items():
        side = border_of(cell)
        if side:
            out.setdefault(side, []).append((r, c))
    return out


def _apply_border(rng, side):
    for edge in _BORDER_EDGES[side]:
        b = rng.Borders(edge)
        b.LineStyle = XL_CONTINUOUS
        b.Weight = XL_THIN


def _addr_groups(coords):
    """[(r, c)] → 같은 행 연속 열을 묶은 주소들을 _ADDR_MAX 이하 문자열 묶음으로."""
    runs, by_row = [], {}
    for r, c in coords:
        by_row.setdefault(r, []).append(c)
    for r in sorted(by_row):
        cols = sorted(by_row[r])
        start = prev = cols[0]
        for c in cols[1:] + [None]:
            if c is not None and c == prev + 1:
                prev = c
                continue
            a = f"{_col_letter(start)}{r}"
            runs.append(a if start == prev else f"{a}:{_col_letter(prev)}{r}")
            if c is not None:
                start = prev = c
    return _join_addrs(runs)


def _apply_style(rng, key):
    f = rng.Font
    if key == "title":
        f.Size = 12
        f.Bold = True
    elif key == "section":
        f.Size = 11
        f.Bold = True
    elif key == "bold":
        f.Bold = True
    elif key == "hdr":
        f.Bold = True
        rng.HorizontalAlignment = XL_CENTER
        rng.VerticalAlignment = XL_CENTER
    elif key in ("italic", "unit"):
        f.Italic = True
    elif key == "flag":
        f.Italic = True
        f.Color = FLAG_COLOR
    elif key == "date":
        set_number_format(rng, COM_DATE_FMT)
        rng.HorizontalAlignment = XL_CENTER
    elif key == "check":
        set_number_format(rng, "General")
        rng.HorizontalAlignment = XL_CENTER
    elif key == "pct":
        set_number_format(rng, PCT_FMT)
    elif key == "wrap":                         # openpyxl: Alignment(wrap_text, vertical top)
        rng.WrapText = True
        rng.VerticalAlignment = XL_VTOP
    elif key == "input":                        # openpyxl: PatternFill solid FFFFCC + Side dotted 사방
        rng.Interior.Color = INPUT_FILL_COLOR
        for edge in (XL_EDGE_LEFT, XL_EDGE_TOP, XL_EDGE_BOTTOM, XL_EDGE_RIGHT):
            rng.Borders(edge).LineStyle = XL_DOT


def _join_addrs(addrs):
    """주소 목록 → 쉼표로 이은 _ADDR_MAX 이하 문자열 묶음."""
    out, cur = [], ""
    for a in addrs:
        if cur and len(cur) + 1 + len(a) > _ADDR_MAX:
            out.append(cur)
            cur = a
        else:
            cur = f"{cur},{a}" if cur else a
    if cur:
        out.append(cur)
    return out


def _apply_check_cf(ws, sh):
    """검증 셀: 값이 FALSE 면 파랑 글꼴(조건부 서식). 시트를 다시 쓰므로 기존 조건부 서식은 먼저 지운다."""
    _retry(lambda: ws.Cells.FormatConditions.Delete())
    for addr in _join_addrs(check_ranges(sh)):
        def add(addr=addr):
            fc = ws.Range(addr).FormatConditions.Add(XL_CELL_VALUE, XL_EQUAL, "=FALSE")
            fc.Font.Color = CHECK_FALSE_COLOR
        _retry(add)


def _ensure_sheet(wb, name, after_name):
    """(시트, 새로 만들었나). 새 시트는 after_name 뒤(없으면 맨 끝)."""
    ws = find_sheet(wb, name)
    if ws is not None:
        return ws, False
    after = find_sheet(wb, after_name) if after_name else None
    if after is None:
        after = wb.Worksheets(wb.Worksheets.Count)
    import pythoncom
    ws = wb.Worksheets.Add(pythoncom.Empty, after)          # late binding: 위치 인수(Before, After)
    ws.Name = name
    return ws, True


def _freeze_split(ref):
    col = "".join(ch for ch in ref if ch.isalpha())
    row = "".join(ch for ch in ref if ch.isdigit())
    n = 0
    for ch in col.upper():
        n = n * 26 + ord(ch) - 64
    return int(row) - 1, n - 1


def _view_settings(ws, sh, is_new):
    """눈금선은 창(Window) 속성이라 활성 시트일 때만 끈다(활성 시트를 바꾸지 않음).
    새로 만든 시트는 막 활성화된 상태라 확대·틀고정도 맞춘다."""
    try:
        wb = ws.Parent
        if wb.ActiveSheet.Name != ws.Name:
            return
        win = wb.Windows(1)
        win.DisplayGridlines = False
        if is_new:
            win.Zoom = 85
            if sh.freeze:
                win.FreezePanes = False
                win.SplitRow, win.SplitColumn = _freeze_split(sh.freeze)
                win.FreezePanes = True
    except Exception:     # noqa: BLE001 — 보기 설정 실패는 데이터에 영향 없음
        pass


def _write_sheet(ws, sh, is_new):
    R, Cc = sh.max_row(), sh.max_col()
    _retry(lambda: ws.Cells.Clear())
    if R == 0:
        _view_settings(ws, sh, is_new)
        return
    vals = [[None] * Cc for _ in range(R)]
    formulas = []
    for (r, c), cell in sh.cells.items():
        if cell.formula:
            formulas.append((r, c, cell.v))
        else:
            vals[r - 1][c - 1] = _com_value(cell.v)
    styles, borders = style_groups(sh), border_groups(sh)
    rng = ws.Range(ws.Cells(1, 1), ws.Cells(R, Cc))
    _retry(lambda: setattr(rng, "Value2", vals))
    for r, c, f in formulas:
        _retry(lambda r=r, c=c, f=f: setattr(ws.Cells(r, c), "Formula", f))

    def base_fmt():
        ws.Cells.Font.Name = "맑은 고딕"
        ws.Cells.Font.Size = 10
        if Cc >= C.RAW_FIRST_VALUE_COL:
            set_number_format(ws.Range(ws.Cells(1, C.RAW_FIRST_VALUE_COL), ws.Cells(R, Cc)), NUM_FMT)
    _retry(base_fmt)
    for key in _STYLE_ORDER:
        for addr in _addr_groups(styles.get(key, [])):
            _retry(lambda addr=addr, key=key: _apply_style(ws.Range(addr), key))
    for side in ("top", "bottom", "both"):
        for addr in _addr_groups(borders.get(side, [])):
            _retry(lambda addr=addr, side=side: _apply_border(ws.Range(addr), side))
    for c, w in sh.col_widths.items():
        _retry(lambda c=c, w=w: setattr(ws.Columns(c), "ColumnWidth", w))
    _apply_check_cf(ws, sh)
    _view_settings(ws, sh, is_new)


def write_model(wb, model, only_data=True):
    """데이터 시트만 다시 쓴다(없으면 모델 순서상 앞 시트 뒤에 추가).
    반환: 모델에 없어 그대로 둔 기존 데이터 시트 이름 목록."""
    app = wb.Application
    try:
        active = wb.ActiveSheet
    except Exception:     # noqa: BLE001
        active = None
    try:
        calc0 = app.Calculation                 # 사용자 원래 계산 방식(수동이면 수동 그대로 돌려준다)
    except Exception:     # noqa: BLE001
        calc0 = XL_CALC_AUTOMATIC
    try:
        _retry(lambda: setattr(app, "Calculation", XL_CALC_MANUAL))
    except Exception:     # noqa: BLE001
        pass
    try:
        prev = "RAW >>"
        for sh in model.sheets:
            if only_data and not C.is_data_sheet(sh.name):
                prev = sh.name
                continue
            ws, is_new = _ensure_sheet(wb, sh.name, prev)
            prev = sh.name
            _write_sheet(ws, sh, is_new)
    finally:
        if active is not None:
            try:
                if wb.ActiveSheet.Name != active.Name:
                    active.Activate()
            except Exception:     # noqa: BLE001
                pass
        try:
            _retry(lambda: setattr(app, "Calculation", calc0))
        except Exception:     # noqa: BLE001
            pass
    names = {s.name for s in model.sheets}
    return [ws.Name for ws in wb.Worksheets if C.is_data_sheet(ws.Name) and ws.Name not in names]
