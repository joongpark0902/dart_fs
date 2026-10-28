"""Excel COM 마무리: VBA 모듈·[업데이트] 버튼·재계산·검증·xlsm 저장.

항상 DispatchEx 로 별도 인스턴스를 띄우고(사용자 Excel 에 붙지 않음) finally 에서 Quit 한다.
"""
import os, shutil, time
from datetime import datetime
import pythoncom
import pywintypes
import win32com.client as win32

XL_OPENXML_MACRO = 52
XL_CELLTYPE_FORMULAS, XL_ERRORS = -4123, 16
XL_CALC_AUTOMATIC = -4105
XL_SECURITY_FORCE_DISABLE = 3      # msoAutomationSecurityForceDisable: 우리 인스턴스에서 열리는 파일의 매크로를 실행하지 않음
MODULE = "DartFs"
NO_CELLS_HRESULT = -2146827284      # 0x800A03EC: SpecialCells 결과 없음
CHECK_PREFIXES = ("=ROUND(", "=OR(ROUND(")       # 옛 `=ABS(ROUND(…))<=1`(±1 허용) 모양은 더 만들지 않아 제외


def is_check_formula(f):
    """검증 셀 판정은 수식 모양으로: CHECK_PREFIXES 로 시작하고 `)=ROUND(` 비교가 있거나 OR 형.
    차이 열 `=ROUND(L-R,0)` 과 mFS 요약 COUNTIF 는 아니다."""
    if not isinstance(f, str):
        return False
    u = f.upper()
    return u.startswith(CHECK_PREFIXES) and (")=ROUND(" in u or u.startswith("=OR(ROUND("))


def scan_checks(vals, forms, r0, c0, sheet, res):
    """값·수식 2-D(UsedRange 기준, r0·c0 = 시작 행·열)를 훑어 res 의 checks·false_cells 를 채운다.
    검증 셀 값이 True 가 아니면(FALSE·오류 코드 등) 전부 FALSE 로 세고 주소를 남긴다."""
    for r, row in enumerate(vals):
        for c, v in enumerate(row):
            if is_check_formula(forms[r][c]):
                ok = v is True
                res["checks"]["TRUE" if ok else "FALSE"] += 1
                if not ok:
                    res["false_cells"].append(f"{sheet}!{_col_letter(c0 + c)}{r0 + r}")
            elif v == "해당 없음":
                res["checks"]["해당 없음"] += 1
    return res


def _retry(fn, tries=5, delay=1.0):
    """0x800AC472(편집 중)·busy 등 일시 오류 재시도. 삽입 같은 비멱등 동작에는 쓰지 않는다."""
    for i in range(tries):
        try:
            return fn()
        except Exception:      # noqa: BLE001
            if i == tries - 1:
                raise
            time.sleep(delay)


def add_button_and_vba(wb, bas_path):
    """열린 통합문서에 VBA 모듈과 버튼을 (중복 없이) 넣는다."""
    comps = wb.VBProject.VBComponents
    for i in range(comps.Count, 0, -1):
        if comps.Item(i).Name == MODULE:
            comps.Remove(comps.Item(i))
    comps.Import(os.path.abspath(bas_path))
    ws = wb.Worksheets(1)                       # 설정 (한글 시트명 대신 인덱스)
    for i in range(ws.Buttons().Count, 0, -1):
        ws.Buttons(i).Delete()
    cell = ws.Range("E4")
    btn = ws.Buttons().Add(cell.Left, cell.Top, 110, 28)
    btn.Caption = "업데이트"
    btn.OnAction = MODULE + ".UpdateData"
    btn.Font.Name = "맑은 고딕"
    btn.Font.Size = 10
    return btn


def _grid(x):
    """COM Range 값: 단일 셀은 스칼라, 여러 셀은 2-D 튜플 → 항상 2-D 리스트."""
    if isinstance(x, tuple):
        return [list(r) if isinstance(r, tuple) else [r] for r in x]
    return [[x]]


def _col_letter(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def verify(wb):
    """전체 시트 오류 셀 수 + `검증` 시트 검증 셀 TRUE/FALSE·'해당 없음' 수(is_check_formula·scan_checks)."""
    res = {"errors": 0, "checks": {"TRUE": 0, "FALSE": 0, "해당 없음": 0}, "false_cells": []}
    for ws in wb.Worksheets:
        try:
            res["errors"] += ws.UsedRange.SpecialCells(XL_CELLTYPE_FORMULAS, XL_ERRORS).Count
        except pywintypes.com_error as e:     # 오류 셀 없음일 때만 삼킴
            if e.hresult != NO_CELLS_HRESULT and not (e.excepinfo and len(e.excepinfo) > 5
                                                      and e.excepinfo[5] == NO_CELLS_HRESULT):
                raise
        if ws.Name != "검증":
            continue
        used = ws.UsedRange
        r0, c0 = used.Row, used.Column
        scan_checks(_grid(used.Value2), _grid(used.Formula), r0, c0, ws.Name, res)
    return res


def _run(open_path, bas_path, save):
    pythoncom.CoInitialize()
    app = None
    wb = None
    try:
        app = win32.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        try:
            app.AutomationSecurity = XL_SECURITY_FORCE_DISABLE
        except Exception:      # noqa: BLE001
            pass
        wb = None
        wb = app.Workbooks.Open(os.path.abspath(open_path))
        try:
            add_button_and_vba(wb, bas_path)
            app.Calculation = XL_CALC_AUTOMATIC
            _retry(lambda: app.CalculateFull())
            res = verify(wb)
            save(wb)
        finally:
            wb.Close(SaveChanges=False)
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:      # noqa: BLE001 — 원래 예외를 가리지 않고 CoUninitialize 까지
                pass
        app = None
        wb = None
        pythoncom.CoUninitialize()
    return res


def backup_existing(path):
    """기존 파일이 있으면 backup_<mtime>_rebuild<확장자> 로 같은 폴더에 옮기고 새 경로 반환(없으면 None)."""
    path = os.path.abspath(path)
    if not os.path.exists(path):
        return None
    d, name = os.path.split(path)
    ext = os.path.splitext(name)[1]
    stamp = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y%m%d-%H%M%S")
    dest = os.path.join(d, f"backup_{stamp}_rebuild{ext}")
    n = 2
    while os.path.exists(dest):
        dest = os.path.join(d, f"backup_{stamp}_rebuild_{n}{ext}")
        n += 1
    try:
        shutil.move(path, dest)
    except PermissionError as e:
        raise RuntimeError(f"기존 파일을 백업으로 옮기지 못했습니다(Excel 등에 열려 있음): {path} — 파일을 닫고 다시 실행하세요.") from e
    return dest


def finalize(path_xlsx, path_xlsm, bas_path):
    """xlsx → (VBA·버튼·재계산·검증) → xlsm. 기존 xlsm 은 백업 후 저장. 성공 시 xlsx 삭제.
    반환: verify 결과 + "backup"(백업 경로 또는 None)."""
    xlsm = os.path.abspath(path_xlsm)
    backup = []

    def save(wb):
        backup.append(backup_existing(xlsm))
        wb.SaveAs(xlsm, FileFormat=XL_OPENXML_MACRO)

    res = _run(path_xlsx, bas_path, save)
    res["backup"] = backup[0] if backup else None
    try:
        os.remove(path_xlsx)
    except OSError:
        pass
    return res


def finalize_inplace(path_xlsm, bas_path):
    """이미 있는 xlsm 에 VBA·버튼을 다시 넣고 재계산·검증 후 같은 파일에 저장."""
    return _run(path_xlsm, bas_path, lambda wb: wb.Save())
