import datetime, json, os, tempfile, unittest
from unittest import mock
from core import fetch as FT

class FakeResp:
    def __init__(self, payload=None, content=b"", status_code=200, bad=False):
        self._p = payload; self.content = content; self.text = json.dumps(payload) if payload else ""
        self.status_code = status_code; self._bad = bad
    def json(self):
        if self._bad:
            raise ValueError("not json")
        return self._p

class TestClient(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.c = FT.DartClient("KEY", self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_list_disclosures_keeps_corrections_and_merges_A_F(self):
        pages = {"A": {"status": "000", "total_page": 1, "list": [{"rcept_no": "1", "report_nm": "사업보고서 (2024.12)", "rcept_dt": "20250321"},
                                                                  {"rcept_no": "2", "report_nm": "[기재정정]사업보고서 (2024.12)", "rcept_dt": "20250401"}]},
                 "F": {"status": "013"}}
        with mock.patch.object(FT.requests, "get", side_effect=lambda url, params, timeout: FakeResp(pages[params["pblntf_ty"]])):
            got = self.c.list_disclosures("00000001", "20240101", "20261231")
        self.assertEqual([g["rcept_no"] for g in got], ["1", "2"])

    def test_fnltt_all_caches_and_returns_none_on_013(self):
        calls = []
        def fake_get(url, params, timeout):
            calls.append(params)
            if params["fs_div"] == "OFS":
                return FakeResp({"status": "013", "message": "없음"})
            return FakeResp({"status": "000", "list": [{"sj_div": "BS", "account_nm": "자산총계", "ord": "1"}]})
        with mock.patch.object(FT.requests, "get", side_effect=fake_get):
            a = self.c.fnltt_all("00000001", "2025", "11011", "CFS")
            b = self.c.fnltt_all("00000001", "2025", "11011", "CFS")
            n = self.c.fnltt_all("00000001", "2025", "11011", "OFS")
        self.assertEqual(a, b); self.assertEqual(len(calls), 2); self.assertIsNone(n)
        self.assertTrue(os.path.exists(os.path.join(self.tmp.name, "api", "00000001_2025_11011_CFS.json")))

    def test_get_retries_non_json_then_ok(self):
        seq = [FakeResp(bad=True), FakeResp({"status": "000", "list": []})]
        with mock.patch.object(FT.requests, "get", side_effect=seq) as g, mock.patch.object(FT.time, "sleep"):
            got = FT._get("http://x/api.json", {})
        self.assertEqual(got["status"], "000"); self.assertEqual(g.call_count, 2)

    def test_get_raises_after_three_non_json(self):
        seq = [FakeResp(bad=True) for _ in range(3)]
        with mock.patch.object(FT.requests, "get", side_effect=seq) as g, mock.patch.object(FT.time, "sleep") as sl:
            with self.assertRaises(RuntimeError):
                FT._get("http://x/api.json", {})
        self.assertEqual(g.call_count, 3); self.assertEqual(sl.call_count, 2)

    def test_get_retries_5xx(self):
        seq = [FakeResp({"status": "000"}, status_code=503), FakeResp({"status": "000"})]
        with mock.patch.object(FT.requests, "get", side_effect=seq) as g, mock.patch.object(FT.time, "sleep"):
            FT._get("http://x/api.json", {})
        self.assertEqual(g.call_count, 2)

    def test_get_retries_rate_limit_020(self):
        seq = [FakeResp({"status": "020", "message": "limit"}), FakeResp({"status": "000", "list": []})]
        with mock.patch.object(FT.requests, "get", side_effect=seq) as g, mock.patch.object(FT.time, "sleep") as sl:
            got = FT._get("http://x/api.json", {})
        self.assertEqual(got["status"], "000"); self.assertEqual(g.call_count, 2); sl.assert_called_with(5)

    def _cache_path(self, fs="OFS"):
        return os.path.join(self.tmp.name, "api", f"00000001_2025_11011_{fs}.json")

    def test_013_cache_stale_requeries(self):
        old = (datetime.datetime.now() - datetime.timedelta(days=10)).isoformat()
        with open(self._cache_path(), "w", encoding="utf-8") as f:
            json.dump({"status": "013", "cached_at": old}, f)
        with mock.patch.object(FT.requests, "get", return_value=FakeResp({"status": "000", "list": [{"a": 1}]})) as g:
            got = self.c.fnltt_all("00000001", "2025", "11011", "OFS")
        self.assertEqual(g.call_count, 1); self.assertEqual(got, [{"a": 1}])

    def test_013_cache_fresh_not_requeried(self):
        with mock.patch.object(FT.requests, "get", return_value=FakeResp({"status": "013"})) as g:
            self.assertIsNone(self.c.fnltt_all("00000001", "2025", "11011", "OFS"))
            self.assertIsNone(self.c.fnltt_all("00000001", "2025", "11011", "OFS"))
        self.assertEqual(g.call_count, 1)
        with open(self._cache_path(), encoding="utf-8") as f:
            self.assertIn("cached_at", json.load(f))

    def test_013_cache_before_filing_date_requeried(self):
        yday = (datetime.datetime.now() - datetime.timedelta(days=1)).isoformat()
        with open(self._cache_path(), "w", encoding="utf-8") as f:
            json.dump({"status": "013", "cached_at": yday}, f)
        today = datetime.date.today()
        with mock.patch.object(FT.requests, "get", return_value=FakeResp({"status": "000", "list": [{"rcept_no": "R", "a": 1}]})) as g:
            got = self.c.fnltt_all("00000001", "2025", "11011", "OFS", expect_rcept_no="R", filed_on=today)
        self.assertEqual(g.call_count, 1); self.assertEqual(got, [{"rcept_no": "R", "a": 1}])

    def test_013_cache_on_filing_date_not_requeried(self):
        with open(self._cache_path(), "w", encoding="utf-8") as f:
            json.dump({"status": "013", "cached_at": datetime.datetime.now().isoformat()}, f)
        with mock.patch.object(FT.requests, "get") as g:
            got = self.c.fnltt_all("00000001", "2025", "11011", "OFS", expect_rcept_no="R", filed_on=datetime.date.today())
        g.assert_not_called(); self.assertIsNone(got)

    def test_corrupt_cache_refetched(self):
        with open(self._cache_path("CFS"), "w", encoding="utf-8") as f:
            f.write("{broken")
        with mock.patch.object(FT.requests, "get", return_value=FakeResp({"status": "000", "list": [{"a": 1}]})) as g:
            got = self.c.fnltt_all("00000001", "2025", "11011", "CFS")
        self.assertEqual(g.call_count, 1); self.assertEqual(got, [{"a": 1}])
        with open(self._cache_path("CFS"), encoding="utf-8") as f:
            self.assertEqual(json.load(f)["status"], "000")

    def test_cached_response_of_other_rcept_is_refetched(self):
        """캐시 키에 접수번호가 없다 — 정정본이 채택되면 캐시(원본 접수번호)를 버리고 다시 받는다."""
        with open(self._cache_path("CFS"), "w", encoding="utf-8") as f:
            json.dump({"status": "000", "list": [{"rcept_no": "OLD", "a": 1}]}, f)
        new = FakeResp({"status": "000", "list": [{"rcept_no": "NEW", "a": 2}]})
        with mock.patch.object(FT.requests, "get", return_value=new) as g:
            got = self.c.fnltt_all("00000001", "2025", "11011", "CFS", expect_rcept_no="NEW")
            again = self.c.fnltt_all("00000001", "2025", "11011", "CFS", expect_rcept_no="NEW")
        self.assertEqual(g.call_count, 1)
        self.assertEqual(got, [{"rcept_no": "NEW", "a": 2}]); self.assertEqual(again, got)
        self.assertIsNone(self.c.last_warning)
        with open(self._cache_path("CFS"), encoding="utf-8") as f:
            self.assertEqual(json.load(f)["list"][0]["rcept_no"], "NEW")

    def test_rcept_mismatch_after_refetch_warns_but_returns(self):
        with open(self._cache_path("CFS"), "w", encoding="utf-8") as f:
            json.dump({"status": "000", "list": [{"rcept_no": "OLD"}]}, f)
        with mock.patch.object(FT.requests, "get", return_value=FakeResp({"status": "000", "list": [{"rcept_no": "OLD"}]})) as g:
            got = self.c.fnltt_all("00000001", "2025", "11011", "CFS", expect_rcept_no="NEW")
        self.assertEqual(g.call_count, 1)
        self.assertEqual(got, [{"rcept_no": "OLD"}])
        self.assertIn("OLD", self.c.last_warning); self.assertIn("NEW", self.c.last_warning)
        with mock.patch.object(FT.requests, "get") as g2:     # 일치하는 호출은 경고를 지운다
            self.c.fnltt_all("00000001", "2025", "11011", "CFS", expect_rcept_no="OLD")
        g2.assert_not_called()
        self.assertIsNone(self.c.last_warning)

    def test_classify_documents(self):
        d = self.tmp.name
        for name, doc in (("a.xml", "감사보고서"), ("b.xml", "연결감사보고서"), ("c.xml", "사업보고서")):
            with open(os.path.join(d, name), "w", encoding="utf-8") as f:
                f.write(f'<DOCUMENT><DOCUMENT-NAME ACODE="1">{doc}</DOCUMENT-NAME><BODY/></DOCUMENT>')
        got = FT.classify_documents([os.path.join(d, n) for n in ("a.xml", "b.xml", "c.xml")])
        self.assertTrue(got["audit_ofs"].endswith("a.xml")); self.assertTrue(got["audit_cfs"].endswith("b.xml")); self.assertTrue(got["main"].endswith("c.xml"))

    def test_documents_uses_vendor_downloader(self):
        with mock.patch.object(FT, "download_document", return_value={"status": "성공", "files": ["x.xml"]}) as dl:
            os.makedirs(os.path.join(self.tmp.name, "doc", "00000001", "999"), exist_ok=True)
            with open(os.path.join(self.tmp.name, "doc", "00000001", "999", "x.xml"), "w", encoding="utf-8") as f:
                f.write("<DOCUMENT><DOCUMENT-NAME>감사보고서</DOCUMENT-NAME></DOCUMENT>")
            got = self.c.documents("00000001", "999")
        dl.assert_called_once()
        self.assertTrue(got["audit_ofs"].endswith("x.xml"))

if __name__ == "__main__":
    unittest.main()
