"""OpenDART 호출과 원문 캐시. vendor 모듈을 그대로 쓰고 이 앱 고유 규칙(정정 포함·캐시)만 더한다."""
import datetime, json, os, re, time
import requests
from dart_client import load_corp_list, search_company          # vendor
from downloader import download_document                         # vendor

_LIST_URL = "https://opendart.fss.or.kr/api/list.json"
_FNLTT_URL = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json"
_DOC_NAME = re.compile(r"<DOCUMENT-NAME[^>]*>(.*?)</DOCUMENT-NAME>", re.I | re.S)
TIMEOUT = 30


def _get(url, params, retries=3):
    """JSON(dict)을 돌려준다. 5xx·비JSON·네트워크 오류·020(한도초과)은 재시도."""
    last = None
    tail = url.rsplit("/", 1)[-1]
    for i in range(retries):
        wait = 1 + i
        try:
            resp = requests.get(url, params=params, timeout=TIMEOUT)
            if resp.status_code >= 500:
                last = f"HTTP {resp.status_code}"
            else:
                try:
                    data = resp.json()
                except ValueError:
                    last = "JSON 아님"
                else:
                    if isinstance(data, dict) and data.get("status") == "020":
                        last, wait = "020 요청 한도 초과", 5
                    else:
                        return data
        except requests.RequestException as e:
            last = e
        if i < retries - 1:
            time.sleep(wait)
    raise RuntimeError(f"DART 요청 실패: {tail} {last}")


def classify_documents(paths):
    out = {"main": None, "audit_ofs": None, "audit_cfs": None, "all": list(paths)}
    for p in paths:
        try:
            with open(p, encoding="utf-8", errors="replace") as f:
                head = f.read(65536)
        except OSError:
            continue
        m = _DOC_NAME.search(head)
        name = re.sub(r"\s+", "", re.sub(r"<[^>]*>", "", m.group(1))) if m else ""
        if "연결" in name and ("감사보고서" in name or "검토보고서" in name):
            out["audit_cfs"] = out["audit_cfs"] or p
        elif "감사보고서" in name or "검토보고서" in name:
            out["audit_ofs"] = out["audit_ofs"] or p
        else:
            out["main"] = out["main"] or p
    return out


class DartClient:
    def __init__(self, api_key, cache_dir, log=None):
        self.api_key = api_key
        self.cache_dir = cache_dir
        self.log = log or (lambda m: None)
        self._corps = None
        self.last_warning = None      # 마지막 fnltt_all 의 접수번호 불일치 경고(파이프라인이 로그에 남김)
        os.makedirs(os.path.join(cache_dir, "api"), exist_ok=True)
        os.makedirs(os.path.join(cache_dir, "doc"), exist_ok=True)

    def corps(self):
        if self._corps is None:
            self._corps = load_corp_list(self.api_key, os.path.join(self.cache_dir, "CORPCODE.xml"), self.log)
        return self._corps

    def search(self, keyword):
        kw = (keyword or "").strip()
        if not kw:
            return []
        corps = self.corps()
        exact_code = [c for c in corps if c.get("stock_code") == kw or c["corp_code"] == kw]
        if exact_code:
            return exact_code
        hits = search_company(corps, kw)
        hits.sort(key=lambda c: (c["corp_name"] != kw, not c.get("stock_code"), len(c["corp_name"])))
        return hits[:300]

    def list_disclosures(self, corp_code, bgn_de, end_de):
        items, seen = [], set()
        for ty in ("A", "F"):
            page = 1
            while True:
                data = _get(_LIST_URL, {"crtfc_key": self.api_key, "corp_code": corp_code, "bgn_de": bgn_de, "end_de": end_de,
                                        "pblntf_ty": ty, "page_no": page, "page_count": 100})
                st = data.get("status")
                if st == "013":
                    break
                if st != "000":
                    raise RuntimeError(f"list.json 오류 {st}: {data.get('message')}")
                for it in data.get("list", []):
                    if it["rcept_no"] not in seen:
                        seen.add(it["rcept_no"])
                        items.append({"rcept_no": it["rcept_no"], "report_nm": it["report_nm"], "rcept_dt": it["rcept_dt"]})
                if page >= int(data.get("total_page") or 1):
                    break
                page += 1
        items.sort(key=lambda i: i["rcept_dt"])
        return items

    @staticmethod
    def _read_cache(path):
        """유효한 캐시 dict 또는 None(없음·손상·013 만료)."""
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, ValueError):
            return None
        if not isinstance(data, dict):
            return None
        if data.get("status") == "013":
            try:
                age = datetime.datetime.now() - datetime.datetime.fromisoformat(data["cached_at"])
            except (KeyError, TypeError, ValueError):
                return None
            if age > datetime.timedelta(days=7):
                return None
        return data

    @staticmethod
    def _rcept_of(data):
        lst = data.get("list") or []
        return (lst[0].get("rcept_no") or None) if lst and isinstance(lst[0], dict) else None

    def fnltt_all(self, corp_code, bsns_year, reprt_code, fs_div, expect_rcept_no=None, filed_on=None):
        """전체 재무제표 API. 캐시 키에 접수번호가 없으므로, expect_rcept_no(채택한 보고서 — 정정본일 수 있음)가
        주어지고 캐시된 000 응답의 접수번호가 다르면 다시 받아 캐시를 덮어쓴다. 다시 받아도 다르면 값은 돌려주되
        self.last_warning 에 경고를 남긴다(없으면 None).
        filed_on(보고서 접수일)이 주어지면 그보다 이른 날 캐시된 013 은 낡은 것으로 보고 다시 받는다."""
        self.last_warning = None
        path = os.path.join(self.cache_dir, "api", f"{corp_code}_{bsns_year}_{reprt_code}_{fs_div}.json")
        data = self._read_cache(path)
        if data is not None and expect_rcept_no and data.get("status") == "000":
            got = self._rcept_of(data)
            if got and got != expect_rcept_no:
                self.log(f"fnlttSinglAcntAll 캐시 접수번호 {got} ≠ 채택 보고서 {expect_rcept_no} → 다시 받음")
                data = None
        if data is not None and expect_rcept_no and filed_on and data.get("status") == "013":
            try:
                cached = datetime.datetime.fromisoformat(data["cached_at"]).date()
            except (KeyError, TypeError, ValueError):
                cached = None
            if cached is None or cached < filed_on:
                self.log("fnlttSinglAcntAll 013 캐시가 보고일보다 앞섬 → 다시 받음")
                data = None
        if data is None:
            data = _get(_FNLTT_URL, {"crtfc_key": self.api_key, "corp_code": corp_code, "bsns_year": str(bsns_year),
                                     "reprt_code": reprt_code, "fs_div": fs_div})
            st0 = data.get("status")
            if st0 in ("000", "013"):
                to_save = {"status": "013", "cached_at": datetime.datetime.now().isoformat()} if st0 == "013" else data
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(to_save, f, ensure_ascii=False)
        st = data.get("status")
        if st == "013":
            return None
        if st != "000":
            raise RuntimeError(f"fnlttSinglAcntAll 오류 {st}: {data.get('message')}")
        got = self._rcept_of(data)
        if expect_rcept_no and got and got != expect_rcept_no:
            self.last_warning = (f"fnlttSinglAcntAll {bsns_year} {reprt_code} {fs_div}: 응답 접수번호 {got} ≠ "
                                 f"채택 보고서 {expect_rcept_no} — API 값이 다른 접수본(정정 전후) 기준일 수 있음")
        return data.get("list", [])

    def documents(self, corp_code, rcept_no):
        d = os.path.join(self.cache_dir, "doc", corp_code, rcept_no)
        os.makedirs(d, exist_ok=True)
        res = download_document(self.api_key, rcept_no, d, log_fn=self.log)
        if res.get("status") == "실패":
            raise RuntimeError(f"원문 다운로드 실패 {rcept_no}: {res.get('error')}")
        paths = [os.path.join(d, f) for f in sorted(os.listdir(d)) if f.lower().endswith(".xml")]
        return classify_documents(paths)
