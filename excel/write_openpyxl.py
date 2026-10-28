"""WorkbookModel → xlsx (새 파일). xlsm 변환·VBA 는 finalize 가 COM 으로.

수식 오인 방지: Cell.formula=True 인 셀만 수식으로 저장한다. '=' 로 시작하는 그 밖의 문자열은
data_type 's' 로 강제해 텍스트로 남긴다.
"""
import openpyxl
from core.normalize import Percent
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

NUM_FMT = '#,##0;[Red](#,##0);"-"'
DATE_FMT = "yyyy-mm-dd"
PCT_FMT = "0.0%"
_NAME = "맑은 고딕"
_THIN = Side(style="thin", color="000000")
INPUT_FILL = "FFFFCC"                 # 입력 칸(fmt "input") 채움 — 연노랑, 사방 점선 테두리
_DOT = Side(style="dotted", color="000000")
CHECK_FALSE_COLOR = "0000FF"          # 검증 셀 FALSE 글꼴(파랑) — 조건부 서식, TRUE 는 검정 그대로


def _font(fmt):
    if fmt == "title":
        return Font(name=_NAME, size=12, bold=True)
    if fmt == "section":
        return Font(name=_NAME, size=11, bold=True)
    if fmt in ("bold", "hdr"):
        return Font(name=_NAME, size=10, bold=True)
    if fmt in ("italic", "unit"):
        return Font(name=_NAME, size=10, italic=True)
    if fmt == "flag":
        return Font(name=_NAME, size=10, italic=True, color="C00000")
    return Font(name=_NAME, size=10)


def border_of(cell):
    """셀의 실선 지정(top/bottom/both/None). 'hdr' 은 지정이 없으면 위·아래 둘 다(한 줄 머리글)."""
    if cell.border:
        return cell.border
    return "both" if cell.fmt == "hdr" else None


def _border(side):
    return Border(top=_THIN if side in ("top", "both") else Side(),
                  bottom=_THIN if side in ("bottom", "both") else Side())


def _is_number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _abs_ref(sheet_name, ref):
    col = "".join(ch for ch in ref if ch.isalpha())
    row = "".join(ch for ch in ref if ch.isdigit())
    return "'{}'!${}${}".format(sheet_name.replace("'", "''"), col, row)


def check_ranges(sh):
    """검증 셀(fmt "check" + 수식) → 같은 열의 연속 행 묶음 주소 ["E6:E20", "F6:F20", ...] (열·행 순)."""
    by_col = {}
    for (r, c), cell in sh.cells.items():
        if cell.fmt == "check" and cell.formula:
            by_col.setdefault(c, []).append(r)
    out = []
    for c in sorted(by_col):
        rows = sorted(by_col[c])
        start = prev = rows[0]
        for r in rows[1:] + [None]:
            if r is not None and r == prev + 1:
                prev = r
                continue
            col = get_column_letter(c)
            out.append(f"{col}{start}" if start == prev else f"{col}{start}:{col}{prev}")
            if r is not None:
                start = prev = r
    return out


def write(model, path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sh in model.sheets:
        write_sheet(wb, wb.create_sheet(sh.name), sh)
    wb.save(path)
    return path


def write_sheet(wb, ws, sh):
    """grid.Sheet 하나를 빈 openpyxl 시트에 쓴다(서식·폭·틀고정·정의된 이름 포함). refresh 닫힌 파일 경로도 사용."""
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    for (r, c), cell in sorted(sh.cells.items()):
        v = cell.v
        xc = ws.cell(row=r, column=c)
        xc.value = v
        if isinstance(v, str) and v.startswith("=") and not cell.formula:
            xc.data_type = "s"          # openpyxl 이 setter 에서 'f' 로 판정한 것을 되돌림
            xc.quotePrefix = True       # 편집 시 수식으로 재해석되지 않게
        xc.font = _font(cell.fmt)
        if cell.fmt == "check" and cell.formula:
            xc.number_format = "General"
            xc.alignment = Alignment(horizontal="center")
        elif isinstance(v, Percent):
            xc.number_format = PCT_FMT
            xc.alignment = Alignment(horizontal="right")
        elif _is_number(v) or (cell.fmt in ("num", "bold") and cell.formula):
            xc.number_format = NUM_FMT
            xc.alignment = Alignment(horizontal="right")
        if cell.fmt == "date":
            xc.number_format = DATE_FMT
            xc.alignment = Alignment(horizontal="center")
        if cell.fmt == "hdr":
            xc.alignment = Alignment(horizontal="center", vertical="center")
        if cell.fmt == "wrap":
            xc.alignment = Alignment(wrap_text=True, vertical="top")
        side = border_of(cell)
        if side:
            xc.border = _border(side)
        if cell.fmt == "input":
            xc.fill = PatternFill(fill_type="solid", start_color=INPUT_FILL, end_color=INPUT_FILL)
            xc.border = Border(left=_DOT, right=_DOT, top=_DOT, bottom=_DOT)
    for c, w in sh.col_widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    rngs = check_ranges(sh)
    if rngs:                            # FALSE → 파랑 글꼴(재계산·업데이트 뒤에도 값 따라 바뀜)
        ws.conditional_formatting.add(" ".join(rngs), CellIsRule(
            operator="equal", formula=["FALSE"], font=Font(name=_NAME, size=10, color=CHECK_FALSE_COLOR)))
    if sh.freeze:
        ws.freeze_panes = sh.freeze
    for name, ref in sh.names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=_abs_ref(sh.name, ref))
    for ref, formula1 in sh.validations:
        dv = DataValidation(type="list", formula1=formula1, allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ref)
    return ws
