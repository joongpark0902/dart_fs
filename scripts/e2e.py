"""실제 DART·Excel 로 끝까지 돌려 산출물을 검증한다.
사용: python scripts/e2e.py [--only ob|hugel] [--out tests/_out/e2e]"""
import argparse, os, sys, json, time
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, HERE)
from excel.build import build
import openpyxl

CASES = {
    "ob": {"corp_code": "00254911", "corp_name": "오비맥주", "as_of": "2026-10-26", "start_year": 2023, "end_year": 2025,
           "include_half": False, "fs_divs": ["OFS"], "unit_display": 0, "note_include": {"recommended": True}},
    "hugel": {"corp_code": "00888347", "corp_name": "휴젤", "as_of": "2026-10-26", "start_year": 2022, "end_year": 2026,
              "include_half": True, "fs_divs": ["CFS", "OFS"], "unit_display": 1000000, "note_include": None},
}


def check_workbook(path):
    wb = openpyxl.load_workbook(path, keep_vba=True)
    names = wb.sheetnames
    assert names[:5] == ["설정", "목차", "mFS", "검증", "RAW >>"], names
    assert any(n.startswith("rFS_") for n in names) and names[-3:] == ["데이터", "매핑표", "로그"], names
    assert wb.vba_archive is not None, "VBA 없음"
    print("  시트:", ", ".join(names))
    return names


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only")
    ap.add_argument("--out", default=os.path.join(HERE, "tests", "_out", "e2e"))
    a = ap.parse_args()
    out_dir = os.path.abspath(a.out)
    os.makedirs(out_dir, exist_ok=True)
    for k, cfg in CASES.items():
        if a.only and a.only != k:
            continue
        print(f"== {k}", flush=True)
        t0 = time.time()
        res = build(cfg, lambda m: print("  ", m, flush=True), out_dir=out_dir)
        sec = round(time.time() - t0, 1)
        names = check_workbook(res["path"])
        v = res["verify"]
        print(f"  {sec}s | 오류 셀 {v['errors']} | TRUE {v['checks']['TRUE']} FALSE {v['checks']['FALSE']} "
              f"해당없음 {v['checks'].get('해당 없음')} | FALSE 셀 {v['false_cells']}")
        with open(os.path.join(out_dir, f"e2e_{k}.json"), "w", encoding="utf-8") as f:
            json.dump({"path": res["path"], "seconds": sec, "sheets": names, "verify": v, "logs": res["logs"]},
                      f, ensure_ascii=False, indent=1, default=str)


if __name__ == "__main__":
    main()
