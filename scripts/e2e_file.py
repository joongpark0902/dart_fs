"""파일 입력 경로를 실제 Excel(COM finalize)까지 돌려 산출물을 검증한다.
가상 '샘플전자' 엑셀(재무제표 2024·2025 + 주석 2025)을 만들어 build(source="file") 에 넣는다.
사용: python scripts/e2e_file.py [--out tests/_out/e2e_file]"""
import argparse, os, sys, json, time
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)
from excel.build import build
from tests.fixtures import xlsx_fixture as F


def make_inputs(d):
    os.makedirs(d, exist_ok=True)
    f = lambda p, end: {"path": p, "period_end": end, "kind": "annual", "fs_div": "OFS", "unit": None}
    return [f(F.build_statements(os.path.join(d, "statements_2024.xlsx"), year=2024), "2024-12-31"),
            f(F.build_statements(os.path.join(d, "statements_2025.xlsx")), "2025-12-31"),
            f(F.build_notes(os.path.join(d, "notes_2025.xlsx")), "2025-12-31")]


def make_cfg(files):
    return {"source": "file", "corp_code": "FILE:샘플전자", "corp_name": "샘플전자", "as_of": "2026-10-27",
            "start_year": 2024, "end_year": 2025, "include_half": False, "fs_divs": ["OFS"], "unit_display": 0,
            "note_include": None, "files": files, "layout": {}}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(HERE, "tests", "_out", "e2e_file"))
    a = ap.parse_args()
    out_dir = os.path.abspath(a.out)
    files = make_inputs(os.path.join(out_dir, "inputs"))
    print("== 샘플전자 (파일 입력)", flush=True)
    t0 = time.time()
    res = build(make_cfg(files), lambda m: print("  ", m, flush=True), out_dir=out_dir)
    sec = round(time.time() - t0, 1)
    v = res["verify"]
    print(f"  {sec}s | 오류 셀 {v['errors']} | TRUE {v['checks']['TRUE']} FALSE {v['checks']['FALSE']} "
          f"해당없음 {v['checks'].get('해당 없음')} | FALSE 셀 {v['false_cells']}")
    print("  산출물:", res["path"])
    with open(os.path.join(out_dir, "e2e_file.json"), "w", encoding="utf-8") as fp:
        json.dump({"path": res["path"], "seconds": sec, "verify": v, "logs": res["logs"]},
                  fp, ensure_ascii=False, indent=1, default=str)
    return 0 if v["errors"] == 0 and v["checks"]["FALSE"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
