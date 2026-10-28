import io
import os
import zipfile
import xml.etree.ElementTree as ET

import requests


# ── 내부 상수 ────────────────────────────────────────────────────────────────
_CORP_CODE_URL  = "https://opendart.fss.or.kr/api/corpCode.xml"
_LIST_URL       = "https://opendart.fss.or.kr/api/list.json"


# ── 1. 회사 목록 로드 ─────────────────────────────────────────────────────────
def load_corp_list(api_key, cache_path="CORPCODE.xml", log_fn=None):
    """
    DART 전체 회사 목록을 반환한다.
    cache_path 파일이 없으면 다운로드 후 저장, 있으면 캐시 사용.
    반환값: [{"corp_code": "...", "corp_name": "..."}, ...]
    """
    def log(msg):
        if log_fn:
            log_fn(msg)

    if not os.path.exists(cache_path):
        log("회사 목록 다운로드 중...")
        resp = requests.get(_CORP_CODE_URL, params={"crtfc_key": api_key})
        if not resp.content.startswith(b"PK"):
            raise RuntimeError(f"corpCode 다운로드 실패: {resp.text[:200]}")
        with zipfile.ZipFile(io.BytesIO(resp.content)) as z:
            z.extractall(os.path.dirname(cache_path) or ".")
        log(f"저장 완료: {cache_path}")
    else:
        log(f"캐시 사용: {cache_path}")

    tree = ET.parse(cache_path)
    corps = [
        {"corp_code": item.findtext("corp_code", ""),
         "corp_name": item.findtext("corp_name", ""),
         "stock_code": (item.findtext("stock_code", "") or "").strip()}
        for item in tree.getroot().findall("list")
    ]
    log(f"회사 목록 로드 완료: {len(corps):,}건")
    return corps


# ── 2. 회사 검색 ──────────────────────────────────────────────────────────────
_KO_TO_EN = {
    "에스케이": "SK",
    "엘지":     "LG",
    "씨제이":   "CJ",
    "케이티":   "KT",
    "지에스":   "GS",
    "에이치디": "HD",
    "디엘":     "DL",
}

def search_company(corp_list, keyword):
    """
    keyword가 corp_name에 포함된 항목을 반환한다 (부분일치, 대소문자 무시).
    한글 표기(에스케이 → SK 등) 자동 변환 후 원래 키워드와 합산, 중복 제거.
    반환값: [{"corp_code": "...", "corp_name": "..."}, ...]
    """
    def _match(kw, corp_name_upper):
        return kw.upper() in corp_name_upper

    # 한글→영문 변환 키워드 생성
    converted = keyword
    for ko, en in _KO_TO_EN.items():
        converted = converted.replace(ko, en)

    keywords = list(dict.fromkeys([keyword, converted]))  # 원본 → 변환 순, 중복 제거

    seen = set()
    results = []
    for c in corp_list:
        name_upper = c["corp_name"].upper()
        if c["corp_code"] not in seen and any(_match(kw, name_upper) for kw in keywords):
            seen.add(c["corp_code"])
            results.append(c)
    return results


# ── 3. 공시 목록 조회 ─────────────────────────────────────────────────────────
# 보고서 유형 → DART 공시유형(pblntf_ty) 매핑
#   A = 정기공시(사업·반기·분기보고서)
#   F = 외부감사관련(감사보고서·연결감사보고서 단독공시)
#       └ 비상장 외부감사대상 법인은 사업보고서를 내지 않고 이쪽만 제출한다.
_PERIODIC_TYPES = ("사업보고서", "반기보고서", "분기보고서")
AUDIT_TYPE      = "감사보고서"     # 부분일치라 '연결감사보고서'도 함께 잡힌다


def _fetch_disclosure_pages(api_key, corp_code, bgn_de, end_de, pblntf_ty, log):
    """
    한 공시유형(pblntf_ty)의 목록을 전체 페이지 모아 반환한다.
    조회 결과가 없으면(status 013) 빈 리스트 — 오류가 아니다.
    """
    items   = []
    page_no = 1
    while True:
        resp = requests.get(_LIST_URL, params={
            "crtfc_key": api_key,
            "corp_code": corp_code,
            "bgn_de": bgn_de,
            "end_de": end_de,
            "pblntf_ty": pblntf_ty,
            "page_no": page_no,
            "page_count": 100,
        })
        data   = resp.json()
        status = data.get("status")

        if status == "013":          # 조회된 데이터 없음
            break
        if status != "000":
            raise RuntimeError(f"list.json 오류: {status} {data.get('message')}")

        items.extend(data.get("list", []))
        total_page = int(data.get("total_page") or 1)
        if page_no >= total_page:
            break
        page_no += 1

    return items


def list_disclosures(api_key, corp_code, bgn_de, end_de,
                     report_types=None, log_fn=None):
    """
    정기공시(A)와 감사보고서 단독공시(F) 목록을 조회해 반환한다.
    report_types: ["사업보고서", "반기보고서", "분기보고서", "감사보고서"] 중 원하는 것.
                  None이면 정기공시 3종을 조회한다.
                  "감사보고서"가 들어 있으면 pblntf_ty=F 도 함께 조회하므로
                  사업보고서를 내지 않는 비상장 외감법인도 받을 수 있다.
    정정본("정정"이 report_nm에 포함)은 항상 제외한다.
    반환값: [{"rcept_no": "...", "report_nm": "...", "rcept_dt": "..."}, ...]
    """
    def log(msg):
        if log_fn:
            log_fn(msg)

    wanted   = list(report_types) if report_types else list(_PERIODIC_TYPES)
    periodic = [rt for rt in wanted if rt in _PERIODIC_TYPES]
    audit    = [rt for rt in wanted if rt not in _PERIODIC_TYPES]

    raw = []
    if periodic:
        got = _fetch_disclosure_pages(api_key, corp_code, bgn_de, end_de, "A", log)
        log(f"정기공시 수신: {len(got)}건")
        raw += got
    if audit:
        got = _fetch_disclosure_pages(api_key, corp_code, bgn_de, end_de, "F", log)
        log(f"외부감사관련 공시 수신: {len(got)}건")
        raw += got

    # rcept_no 중복 제거 (A·F 양쪽에 걸리는 공시 대비)
    seen, items = set(), []
    for i in raw:
        if i["rcept_no"] not in seen:
            seen.add(i["rcept_no"])
            items.append(i)

    # 정정본 제외
    items = [i for i in items if "정정" not in i["report_nm"]]

    # 보고서 유형 필터
    items = [i for i in items if any(rt in i["report_nm"] for rt in wanted)]
    items.sort(key=lambda i: i["rcept_dt"])

    log(f"필터 후: {len(items)}건")
    if len(items) == 0:
        if audit:
            log("해당 기간에 받을 공시가 없습니다. "
                "연도 범위를 넓혀 보세요(감사보고서는 결산 다음 해 3~4월에 공시됩니다).")
        else:
            log("정기공시(사업/반기/분기보고서) 제출 이력이 없습니다. "
                "비상장 외감법인은 감사보고서만 제출하므로 "
                "'감사보고서'를 체크하고 다시 시도하세요.")
    return [
        {"rcept_no": i["rcept_no"],
         "report_nm": i["report_nm"],
         "rcept_dt": i["rcept_dt"]}
        for i in items
    ]
