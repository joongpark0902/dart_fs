"""생성 진입점: 설정 → collect → compose → xlsx → finalize(xlsm)."""
import os, re
from core import config as CFG, mapping as M, pipeline as P
from excel import compose as C, write_openpyxl as W

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BAS = os.path.join(HERE, "excel", "vba", "DartFs.bas")
REFRESH = os.path.join(HERE, "excel", "refresh.py")


def safe(name):
    return re.sub(r'[\\/:*?"<>|\s]+', "_", name or "").strip("_")[:60] or "회사"


def build(cfg_dict, log, client=None, out_dir=None):
    conf = CFG.load()
    if client is None:
        if not conf["api_key"]:
            raise RuntimeError("config.txt 에 api_key 가 없습니다.")
        from core.fetch import DartClient
        client = DartClient(conf["api_key"], CFG.CACHE_DIR, log)
    from excel import finalize as FZ           # pywin32 는 마무리 단계에서만 필요
    cfg = P.BuildConfig.from_dict(cfg_dict)
    company_dir = safe(cfg.corp_name).strip(". ") or "회사"
    out_dir = out_dir or os.path.join(CFG.OUTPUT_DIR, company_dir)
    base = os.path.join(out_dir, f"{safe(cfg.corp_name)}_{cfg.as_of.isoformat()}")
    xlsx, xlsm = base + ".xlsx", base + ".xlsm"
    lock = os.path.join(out_dir, "~$" + os.path.basename(xlsm))
    if os.path.exists(lock):                   # 수집(수 분) 전에 막는다
        raise RuntimeError(f"출력 파일이 Excel에 열려 있습니다: {xlsm} — 닫고 다시 실행하세요(잠금 파일 {lock}).")
    mapping = M.load()
    col = P.collect(cfg, client, mapping, log)
    model = C.compose(col, None, conf["python"], REFRESH)
    log(f"표시 단위 {model.unit_text}")
    os.makedirs(out_dir, exist_ok=True)
    log("엑셀 쓰기(openpyxl)…")
    W.write(model, xlsx)
    log("Excel COM 마무리(VBA·버튼·재계산·검증)…")
    res = FZ.finalize(xlsx, xlsm, BAS)
    log(f"완료: {xlsm} | 오류 셀 {res['errors']} | 검증 TRUE {res['checks']['TRUE']} FALSE {res['checks']['FALSE']}")
    if res["checks"]["FALSE"]:
        log(f"검증 FALSE 셀: {', '.join(res['false_cells'])}")
    return {"path": xlsm, "verify": res, "logs": col.logs}
