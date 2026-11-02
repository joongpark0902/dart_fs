"""[업데이트] 진입점(VBA 버튼 → python refresh.py --workbook <path>).

열린 Excel 에 붙어(GetActiveObject — 사용자 Excel, DispatchEx 아님) 데이터 시트만 COM 으로 다시 쓴다.
파일이 Excel 에 안 열려 있으면 우리 전용 Excel 인스턴스(DispatchEx)로 열어 같은 핵심(_refresh_workbook)으로
데이터 시트만 다시 쓰고 버튼·VBA 복원·재계산·검증 후 저장한다(사용자 시트의 차트·이미지·도형 보존).
Excel COM 을 못 쓸 때만 openpyxl 경로(run_closed_openpyxl)로 대신한다(차트·이미지 손실 가능 — 경고 로그).
다른 Excel 인스턴스가 파일을 잡고 있으면(~$ 잠금 파일) 닫힌 경로를 거부한다.
어느 경로든 먼저 backup_<mtime>_update.xlsm 백업을 만든다.
사용자 Excel 은 절대 Quit 하지 않는다. 열린 경로의 저장은 사용자 몫(Ctrl+S).
업데이트는 기준일을 오늘로, 끝 연도를 max(저장된 끝 연도, 올해)로 올려 새 공시를 받는다(시작 연도는 유지).
표시 단위: 설정 JSON unit_display 0(자동)이고 cfg_unit 셀이 지난 판정값(unit_judged) 그대로면 다시 판정해 cfg_unit 을
새로 쓰고, 셀을 바꿨거나 직접 선택(0 아님)이면 셀 값을 그대로 쓴다.
버튼으로 띄운 실행은 창이 보이지 않으므로 모든 로그·오류 추적을 통합문서 폴더의 dart_fs_update.log 에 덧붙인다.
"""
import argparse, datetime as dt, json, os, shutil, sys, traceback

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from core import config as CFG, mapping as M, pipeline as P      # noqa: E402
from core.fetch import DartClient                                  # noqa: E402
from excel import compose as C, finalize as FZ, write_com as WC  # noqa: E402
from excel.write_com import blocks_from_columns                   # noqa: E402,F401
from excel.build import BAS, REFRESH                              # noqa: E402

MAPPING_USER_DIR = os.path.join(HERE, "mapping", "user")
LOG_NAME = "dart_fs_update.log"
XL_SECURITY_FORCE_DISABLE = 3          # msoAutomationSecurityForceDisable: 우리 인스턴스에서 매크로 실행 안 함
_TRUE_STR = ("TRUE", "Y", "예", "1", "O")


def backup(path):
    """같은 폴더에 backup_<mtime YYYYMMDD-HHMMSS>_update<ext> 복사(이미 있으면 _2, _3 …). 경로 반환."""
    path = os.path.abspath(path)
    d, name = os.path.split(path)
    ext = os.path.splitext(name)[1] or ".xlsm"
    stamp = dt.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y%m%d-%H%M%S")
    dst = os.path.join(d, f"backup_{stamp}_update{ext}")
    n = 2
    while os.path.exists(dst):
        dst = os.path.join(d, f"backup_{stamp}_update_{n}{ext}")
        n += 1
    shutil.copy2(path, dst)
    return dst


def _included(v):
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip().upper() in _TRUE_STR


def parse_toc_rows(rows):
    """목차 시트 행(헤더 포함) → {정규화 제목: 포함 여부}. 포함=D열, 정규화 제목=H열."""
    out = {}
    for r in (rows or [])[1:]:
        if not r or len(r) < 8 or r[7] in (None, ""):
            continue
        out[str(r[7])] = _included(r[3])
    return out


def _attach(path):
    """실행 중인 Excel(GetActiveObject)에서 path 통합문서를 찾는다. (app, wb) 또는 (None, None)."""
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    try:
        app = win32.GetActiveObject("Excel.Application")
    except Exception:    # noqa: BLE001 — Excel 미실행
        return None, None
    target = os.path.normcase(os.path.abspath(path))
    for wb in app.Workbooks:
        try:
            if os.path.normcase(wb.FullName) == target:
                return app, wb
        except Exception:    # noqa: BLE001
            continue
    return None, None


def _setting(rows, label):
    """설정 시트 행에서 B열 라벨의 C열 값."""
    for row in rows or []:
        if row and len(row) > 2 and row[1] == label:
            return row[2]
    return None


def _as_unit(v):
    """셀 값 → 양의 정수 나눗수 또는 None."""
    if v in (None, "") or isinstance(v, bool):
        return None
    try:
        u = int(float(v))
    except (TypeError, ValueError):
        return None
    return u if u > 0 else None


def resolve_unit(cfg_dict, cell):
    """업데이트 때 쓸 표시 단위: 0 = 자동(다시 판정), 그 밖 = 직접 선택한 나눗수.
    자동(JSON unit_display 0)은 셀이 비었거나 지난 판정값(unit_judged, 없으면 아무 값)과 같으면 자동 유지,
    다르면 사용자가 엑셀에서 바꾼 것 → 그 값. 직접 선택은 셀 값(못 읽으면 JSON 값)."""
    mode = _as_unit((cfg_dict or {}).get("unit_display")) or 0
    u = _as_unit(cell)
    if mode == 0:
        judged = _as_unit((cfg_dict or {}).get("unit_judged"))
        return 0 if u is None or judged is None or u == judged else u
    return u or mode


def _collect_for(reader, log):
    """reader(sheet_name) -> 2-D 행 리스트 또는 None. 설정·목차·매핑표로 설정을 복원하고 collect.
    반환 (cfg, col, conf)."""
    settings = reader("설정")
    cfg_json = _setting(settings, "설정 JSON")
    if not cfg_json:
        raise RuntimeError("설정 시트의 '설정 JSON' 을 찾지 못함")
    cfg_dict = json.loads(cfg_json)
    src = (cfg_dict or {}).get("source") or "dart"
    if src != "dart":                     # 예전 '파일로 생성' 산출물(source=file) — 기능 보류로 업데이트 불가
        raise RuntimeError(f"이 통합문서는 DART 가 아닌 입력(source={src})으로 만들어져 업데이트할 수 없습니다 — "
                           "회사 제시 엑셀 입력 기능은 보류되었습니다")
    cfg = P.BuildConfig.from_dict(cfg_dict)
    today = dt.date.today()               # 새 공시를 받도록 기준일·끝 연도를 올린다(시작 연도 유지)
    cfg.as_of = today
    cfg.end_year = max(cfg.end_year, today.year)
    log(f"업데이트 기준일 {cfg.as_of.isoformat()}, 연도 {cfg.start_year}~{cfg.end_year}")
    cfg.unit_display = resolve_unit(cfg_dict, _setting(settings, "mFS 표시 단위"))
    log("표시 단위: 자동(공시 단위 다시 판정)" if cfg.unit_display == 0 else f"표시 단위: 직접 선택 {cfg.unit_display}")
    inc = parse_toc_rows(reader("목차"))
    if inc:
        cfg.note_include = inc
    rows = reader("매핑표")
    if rows and len(rows) > 1 and any(r and r[0] not in (None, "") for r in rows[1:]):
        M.save_from_rows(rows, MAPPING_USER_DIR)
        mapping = M.load(MAPPING_USER_DIR)
        log(f"매핑표 시트 → {MAPPING_USER_DIR}")
    else:
        mapping = M.load()
        log("매핑표 시트가 없거나 비어 있어 기본 매핑 사용")
    conf = CFG.load()
    if not conf.get("api_key"):
        raise RuntimeError("config.txt 에 api_key 가 없습니다.")
    client = DartClient(conf["api_key"], CFG.CACHE_DIR, log)
    return cfg, P.collect(cfg, client, mapping, log), conf


def unit_info(model):
    """compose 결과 → settings_updates 용 표시 단위 정보."""
    return {"value": model.unit_value, "auto": model.unit_auto, "text": model.unit_text}


def settings_updates(rows, cfg, unit=None):
    """설정 시트 행 → [(행, 열, 값)] (1부터): 기준일·연도 범위·설정 JSON(as_of·end_year 만 바꿈).
    unit(unit_info) 이 자동이면 'mFS 표시 단위'(cfg_unit)에 새 판정값, JSON unit_judged 갱신. '단위 판정' 문구는 늘 갱신."""
    out = []
    auto = bool(unit and unit.get("auto"))
    for i, row in enumerate(rows or []):
        if not row or len(row) < 3:
            continue
        lab = row[1]
        if lab == "기준일":
            out.append((i + 1, 3, cfg.as_of))
        elif lab == "연도 범위":
            out.append((i + 1, 3, f"{cfg.start_year}~{cfg.end_year}"))
        elif lab == "mFS 표시 단위" and auto:
            out.append((i + 1, 3, unit["value"]))
        elif lab == "단위 판정" and unit:
            out.append((i + 1, 3, unit["text"]))
        elif lab == "설정 JSON" and row[2]:
            try:
                d = json.loads(row[2])
            except (TypeError, ValueError):
                continue
            d["as_of"], d["end_year"] = cfg.as_of.isoformat(), cfg.end_year
            if auto:
                d["unit_judged"] = unit["value"]
            out.append((i + 1, 3, json.dumps(d, ensure_ascii=False)))
    return out


def _write_settings_com(wb, cfg, unit=None):
    ws = WC.find_sheet(wb, "설정")
    if ws is None:
        return
    for r, c, v in settings_updates(FZ._retry(lambda: WC.sheet_rows(ws)), cfg, unit):
        FZ._retry(lambda r=r, c=c, v=v: setattr(ws.Cells(r, c), "Value2", WC._com_value(v)))


def shift_tail(shifts):
    """블록 밀림 경고 → 상태 문구 꼬리(없으면 '')."""
    if not shifts:
        return ""
    more = f" 외 {len(shifts) - 3}건(로그 시트)" if len(shifts) > 3 else ""
    return f" | 경고 블록 밀림 {len(shifts)}건: " + "; ".join(shifts[:3]) + more


def _status_text(res, tail, shifts=()):
    return (f"업데이트 완료 {dt.datetime.now():%Y-%m-%d %H:%M} | 오류 {res['errors']} | "
            f"검증 TRUE {res['checks']['TRUE']} FALSE {res['checks']['FALSE']}{shift_tail(shifts)}{tail}")


def set_status_com(wb, text):
    rng = wb.Names("cfg_status").RefersToRange
    FZ._retry(lambda: setattr(rng, "Value2", "'" + text))


def _com_reader(wb):
    def reader(name):
        ws = WC.find_sheet(wb, name)
        return FZ._retry(lambda: WC.sheet_rows(ws)) if ws is not None else None
    return reader


def _refresh_workbook(app, wb, log, save):
    """COM 통합문서에 데이터 시트만 다시 쓰고 재계산·검증 → 상태 셀 기록(검증 뒤).
    save=True(닫힌 파일, 우리 인스턴스): 버튼·VBA 복원 후 저장. save=False(사용자 Excel): 저장 안 함."""
    try:
        calc0 = app.Calculation                # 사용자 원래 계산 방식 — 끝나면 그대로 되돌린다
    except Exception:    # noqa: BLE001
        calc0 = WC.XL_CALC_AUTOMATIC
    try:
        existing = FZ._retry(lambda: WC.read_existing(wb))
        cfg, col, conf = _collect_for(_com_reader(wb), log)
        model = C.compose(col, existing, conf["python"], REFRESH)
        for w in model.shift_warnings:
            log(w)
        log("Excel 에 쓰는 중(COM)…")
        stale = WC.write_model(wb, model, only_data=True)
        for name in stale:
            log(f"이번 결과에 없는 데이터 시트 '{name}' 는 그대로 둠")
        _write_settings_com(wb, cfg, unit_info(model))
        if model.unit_auto:
            log(f"표시 단위 {model.unit_text}")
        if save:
            log("버튼·VBA 복원…")
            FZ.add_button_and_vba(wb, BAS)
        FZ._retry(app.CalculateFull)
        res = FZ.verify(wb)
        set_status_com(wb, _status_text(res, "" if save else " (저장은 Ctrl+S)", model.shift_warnings))
        if save:
            FZ._retry(wb.Save)
        return res
    finally:
        try:
            if app.Calculation != calc0:
                FZ._retry(lambda: setattr(app, "Calculation", calc0))
        except Exception:    # noqa: BLE001
            pass


def run_open(app, wb, log):
    """열린 통합문서(COM)에 데이터 시트만 다시 쓰고 재계산·검증. 저장은 하지 않는다."""
    return _refresh_workbook(app, wb, log, save=False)


def _existing_openpyxl(wb):
    return {ws.title: C.existing_from_rows([list(r) for r in ws.iter_rows(values_only=True)])
            for ws in wb.worksheets if C.is_struct_sheet(ws.title)}


def _status_cell_openpyxl(wb):
    try:
        dn = wb.defined_names["cfg_status"]
    except KeyError:
        return None
    for title, ref in dn.destinations:
        if title in wb.sheetnames:
            return wb[title][ref.replace("$", "")]
    return None


COM_FALLBACK_WARN = "Excel COM 사용 불가 → openpyxl 경로(사용자 시트의 차트·이미지가 사라질 수 있음)"
LOCKED_MSG = "파일이 다른 Excel 인스턴스에 열려 있습니다. 그 Excel 안의 [업데이트] 버튼을 쓰거나 파일을 닫으세요."


def lock_path(path):
    d, name = os.path.split(os.path.abspath(path))
    return os.path.join(d, "~$" + name)


def run_closed(path, log):
    """Excel 에 안 열린 파일: 우리 DispatchEx 인스턴스로 열어 _refresh_workbook(save=True) → Quit.
    COM 을 못 쓰면 run_closed_openpyxl 로 대신(경고 로그)."""
    path = os.path.abspath(path)
    if os.path.exists(lock_path(path)):
        raise RuntimeError(LOCKED_MSG)
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        log(COM_FALLBACK_WARN)
        return run_closed_openpyxl(path, log)
    pythoncom.CoInitialize()
    app = wb = None
    try:
        try:
            app = win32.DispatchEx("Excel.Application")
        except Exception:    # noqa: BLE001 — Excel 미설치·COM 서버 실패
            app = None
        if app is None:                        # DispatchEx 실패일 때만 openpyxl 대체 경로
            log(COM_FALLBACK_WARN)
            return run_closed_openpyxl(path, log)
        app.Visible = False
        app.DisplayAlerts = False
        try:
            app.AutomationSecurity = XL_SECURITY_FORCE_DISABLE
        except Exception:    # noqa: BLE001
            pass
        log("파일을 별도 Excel 인스턴스로 여는 중(COM)…")
        wb = app.Workbooks.Open(path)
        return _refresh_workbook(app, wb, log, save=True)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:    # noqa: BLE001
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:    # noqa: BLE001 — 원래 예외를 가리지 않고 CoUninitialize 까지
                pass
        app = wb = None
        pythoncom.CoUninitialize()


OPENPYXL_NOTE = ("Excel COM 없이 openpyxl 로 저장함 — [업데이트] 버튼·VBA 복원, 재계산·검증은 하지 않았습니다"
                 "(버튼이 사라졌을 수 있음). 파일을 Excel 에서 열어 확인하세요.")


def run_closed_openpyxl(path, log):
    """COM 불가(DispatchEx 실패·pywin32 없음) 시 대체 경로: openpyxl(keep_vba) 로 데이터 시트만 교체·저장.
    Excel 이 없으므로 finalize_inplace(버튼·VBA 복원·재계산·검증)는 부르지 않는다 — 같은 이유로 실패하기 때문.
    주의: openpyxl 은 사용자 시트의 차트·이미지·도형·피벗·양식 단추를 잃을 수 있다."""
    import openpyxl
    from excel import write_openpyxl as W
    wb = openpyxl.load_workbook(path, keep_vba=True)

    def reader(name):
        return [list(r) for r in wb[name].iter_rows(values_only=True)] if name in wb.sheetnames else None
    existing = _existing_openpyxl(wb)
    cfg, col, conf = _collect_for(reader, log)
    model = C.compose(col, existing, conf["python"], REFRESH)
    for w in model.shift_warnings:
        log(w)
    log("파일에 직접 쓰는 중(openpyxl)…")
    active_title = wb.active.title if wb.active is not None else None
    prev = "RAW >>"
    for sh in model.sheets:
        if not C.is_data_sheet(sh.name):
            prev = sh.name
            continue
        if sh.name in wb.sheetnames:
            idx = wb.sheetnames.index(sh.name)
            wb.remove(wb[sh.name])
        else:
            idx = wb.sheetnames.index(prev) + 1 if prev in wb.sheetnames else len(wb.sheetnames)
        W.write_sheet(wb, wb.create_sheet(sh.name, idx), sh)
        prev = sh.name
    names = {s.name for s in model.sheets}
    for title in wb.sheetnames:
        if C.is_data_sheet(title) and title not in names:
            log(f"이번 결과에 없는 데이터 시트 '{title}' 는 그대로 둠")
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = False
    if active_title in wb.sheetnames:
        wb.active = wb.sheetnames.index(active_title)
    wb.active.sheet_view.tabSelected = True
    if "설정" in wb.sheetnames:
        ws = wb["설정"]
        for r, c, v in settings_updates(reader("설정"), cfg, unit_info(model)):
            ws.cell(row=r, column=c).value = v
    cell = _status_cell_openpyxl(wb)
    if cell is not None:
        cell.value = (f"업데이트 완료 {dt.datetime.now():%Y-%m-%d %H:%M} (openpyxl 기록 — 버튼·VBA 복원·재계산·검증 안 함)"
                      f"{shift_tail(model.shift_warnings)}")
    wb.save(path)
    log(OPENPYXL_NOTE)
    return {"errors": None, "checks": {"TRUE": None, "FALSE": None, "해당 없음": None}, "false_cells": [],
            "verified": False}


def main(argv=None):
    ap = argparse.ArgumentParser(description="dart_fs [업데이트]")
    ap.add_argument("--workbook", required=True)
    a = ap.parse_args(argv)
    path = os.path.abspath(a.workbook)
    log_path = os.path.join(os.path.dirname(path), LOG_NAME)
    try:
        fh = open(log_path, "a", encoding="utf-8")
        fh.write(f"\n===== {dt.datetime.now():%Y-%m-%d %H:%M:%S} dart_fs 업데이트: {path} =====\n")
        fh.flush()
    except OSError:
        fh = None

    def log(m):
        line = f"{dt.datetime.now():%H:%M:%S} {m}"
        print(line, flush=True)
        if fh is not None:
            try:
                fh.write(line + "\n")
                fh.flush()
            except OSError:
                pass

    app = wb = None
    bk = None
    try:
        bk = backup(path)
        log(f"백업: {bk}")
        app, wb = _attach(path)
        if wb is not None:
            res = run_open(app, wb, log)
        else:
            log("Excel 에 열려 있지 않음 → 파일 직접 갱신")
            res = run_closed(path, log)
        if res.get("verified") is False:
            log("완료(재계산·검증 안 함 — Excel 에서 열어 확인)")
        else:
            log(f"완료 | 오류 {res['errors']} | 검증 TRUE {res['checks']['TRUE']} FALSE {res['checks']['FALSE']}")
        if res["checks"]["FALSE"]:
            log(f"검증 FALSE 셀: {', '.join(res.get('false_cells', []))}")
        return 0
    except Exception as e:    # noqa: BLE001
        tb = traceback.format_exc()
        print(tb, flush=True)
        log(f"업데이트 실패: {e}")
        if fh is not None:
            try:
                fh.write(tb)
            except OSError:
                pass
        if wb is not None:
            try:
                set_status_com(wb, f"업데이트 실패: {e} — 저장하지 말고 닫으세요(백업: {bk}, 로그: {log_path})")
            except Exception:    # noqa: BLE001
                pass
        return 1
    finally:
        if fh is not None:
            fh.close()


if __name__ == "__main__":
    sys.exit(main())
