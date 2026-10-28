import json, os, unittest
from datetime import date
from core import pipeline as P, mapping as M

FX = os.path.join(os.path.dirname(__file__), "fixtures")


class FakeClient:
    """오비맥주(비상장, API 없음)와 휴젤(상장, API 있음)을 픽스처로 흉내낸다."""
    def __init__(self):
        self.calls = []

    def list_disclosures(self, corp_code, bgn, end):
        if corp_code == "OB":
            return [{"rcept_no": "OB2025", "report_nm": "감사보고서 (2025.12)", "rcept_dt": "20260410"},
                    {"rcept_no": "OB2024", "report_nm": "감사보고서 (2024.12)", "rcept_dt": "20250404"}]
        return [{"rcept_no": "HG2025", "report_nm": "사업보고서 (2025.12)", "rcept_dt": "20260320"},
                {"rcept_no": "HG2026H", "report_nm": "반기보고서 (2026.06)", "rcept_dt": "20260814"}]

    last_warning = None

    def fnltt_all(self, corp_code, year, reprt, fs_div, expect_rcept_no=None, filed_on=None):
        self.calls.append((corp_code, year, reprt, fs_div))
        if corp_code == "OB" or fs_div != "CFS":
            return None
        name = "hugel_2025_annual_cfs.json" if reprt == "11011" else "hugel_2026_half_cfs.json"
        with open(os.path.join(FX, name), encoding="utf-8") as f:
            return json.load(f)["list"]

    def documents(self, corp_code, rcept_no):
        if corp_code == "OB":
            p = os.path.join(FX, "ob_2025_audit_ofs.xml")      # 두 연도 모두 같은 파일(테스트용)
            return {"main": p, "audit_ofs": p, "audit_cfs": None, "all": [p]}
        if rcept_no == "HG2025":
            return {"main": None, "audit_ofs": os.path.join(FX, "hugel_2025_audit_ofs.xml"),
                    "audit_cfs": os.path.join(FX, "hugel_2025_audit_cfs.xml"), "all": []}
        return {"main": None, "audit_ofs": None, "audit_cfs": None, "all": []}


class TestUnlisted(unittest.TestCase):
    def setUp(self):
        self.cfg = P.BuildConfig("OB", "오비맥주", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1, None, "default")
        self.m = M.load()

    def test_toc_preview(self):
        base, toc = P.toc_preview(self.cfg, FakeClient(), self.m, lambda s: None)
        self.assertEqual(base.rcept_no, "OB2025")
        self.assertEqual(len(toc), 35)
        self.assertEqual(next(e for e in toc if e.no == "27").topic, "판관비")
        self.assertTrue(next(e for e in toc if e.no == "11").recommended)

    def test_collect_fs_from_xml_and_restated_priority(self):
        c = P.collect(self.cfg, FakeClient(), self.m, lambda s: None)
        self.assertEqual(c.periods, ["FY2024", "FY2025"])
        fy24 = [r for r in c.fs_records if r.period_key == "FY2024" and r.sj_div == "BS"]
        self.assertTrue(fy24 and all(r.rcept_no == "OB2025" and r.role == "prev" for r in fy24))
        fy25 = [r for r in c.fs_records if r.period_key == "FY2025" and r.sj_div == "BS"]
        self.assertTrue(all(r.rcept_no == "OB2025" and r.role == "cur" for r in fy25))
        self.assertTrue(all(r.source == "XML" for r in c.fs_records))
        self.assertIn(("fiscal", "FY2025"), c.units)

    def test_collect_notes_recommended_only(self):
        self.cfg.note_include = {"recommended": True}
        c = P.collect(self.cfg, FakeClient(), self.m, lambda s: None)
        topics = {r.topic for r in c.note_records}
        self.assertIn("판관비", topics); self.assertIn("유형자산", topics)
        self.assertNotIn(None, topics)
        sga = [r for r in c.note_records if r.topic == "판관비" and r.row_label == "급여"]
        self.assertEqual({r.period_key: r.value for r in sga}, {"FY2025": 130500219, "FY2024": 130368479})
        self.assertTrue(any("기준 보고서" in s for s in c.logs))

    def test_raw_tables_kept_for_every_year(self):
        self.cfg.note_include = {"recommended": True}
        c = P.collect(self.cfg, FakeClient(), self.m, lambda s: None)
        raw = [r for r in c.note_records if r.raw_cell is not None]
        self.assertEqual({r.period_key for r in raw}, {"FY2024", "FY2025"})
        keys = [(r.rcept_no, r.note_no, r.table_idx, r.raw_cell, r.row_label, r.col_label) for r in raw]
        self.assertEqual(len(keys), len(set(keys)))

    def test_note_falls_back_to_own_report(self):
        """FY2024 의 own 보고서에만 있는 주석(다음 해 보고서엔 없음)도 own 보고서에서 채운다."""
        import core.pipeline as PP
        orig = PP.NP.note_tables
        calls = {}
        def patched(nodes, entry):
            if entry.topic == "판관비" and calls.get("blank"):
                return []
            return orig(nodes, entry)
        orig_pr = PP._notes_for_report
        def wrapped(*a, **k):
            report = a[2] if len(a) > 2 else k["report"]
            calls["blank"] = report.rcept_no == "OB2025"
            return orig_pr(*a, **k)
        PP.NP.note_tables = patched
        PP._notes_for_report = wrapped
        try:
            cfg = P.BuildConfig("OB", "오비맥주", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1, {"recommended": True}, "default")
            c = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        finally:
            PP.NP.note_tables = orig
            PP._notes_for_report = orig_pr
        sga = [r for r in c.note_records if r.topic == "판관비" and r.row_label == "급여"]
        self.assertTrue(sga)
        self.assertTrue(all(r.rcept_no == "OB2024" for r in sga))


class TestListed(unittest.TestCase):
    def setUp(self):
        cfg = P.BuildConfig("HG", "휴젤", date(2026, 10, 26), 2024, 2026, True, ["CFS"], 1, {"recommended": True}, "default")
        self.c = P.collect(cfg, FakeClient(), M.load(), lambda s: None)

    def test_collect_api_and_half(self):
        c = self.c
        self.assertEqual(c.periods, ["FY2024", "FY2025", "2025H1", "2026H1"])  # selector: 반기 대칭
        self.assertTrue(any(r.source == "API" and r.period_key == "FY2024" and r.role == "prev" for r in c.fs_records))
        self.assertTrue(any(r.period_key == "2026H1" and r.sj_div == "CIS" for r in c.fs_records))
        self.assertFalse(any(r.period_key == "2026H1" for r in c.note_records))      # 반기 주석 없음(1차 범위)
        self.assertTrue(any(r.fs_div == "CFS" and r.topic == "유형자산" for r in c.note_records))
        self.assertEqual(P.fiscal_from_api([{"thstrm_nm": "제 25 기"}]), 25)

    def test_half_has_bs_and_cis_per_sj_selection(self):
        c = self.c
        h1 = {r.sj_div for r in c.fs_records if r.period_key == "2026H1"}
        self.assertIn("BS", h1); self.assertIn("CIS", h1)
        h25 = [r for r in c.fs_records if r.period_key == "2025H1" and r.sj_div in ("IS", "CIS")]
        self.assertTrue(h25 and all(r.rcept_no == "HG2026H" and r.role == "prev" for r in h25))
        # FY2025 BS: 2025 반기 보고서가 없으므로 사업보고서(HG2025, cur)에서 온다
        fy25 = [r for r in c.fs_records if r.period_key == "FY2025" and r.sj_div == "BS"]
        self.assertTrue(fy25 and all(r.rcept_no == "HG2025" and r.role == "cur" for r in fy25))

    def test_api_rcept_warning_logged_and_rcept_passed(self):
        class Warn(FakeClient):
            def fnltt_all(self, corp_code, year, reprt, fs_div, expect_rcept_no=None, filed_on=None):
                self.calls.append(expect_rcept_no)
                got = super().fnltt_all(corp_code, year, reprt, fs_div)
                self.last_warning = "응답 접수번호 X ≠ 채택 보고서 Y" if got else None
                return got
        cl = Warn()
        cfg = P.BuildConfig("HG", "휴젤", date(2026, 10, 26), 2024, 2026, True, ["CFS"], 1, {"recommended": True}, "default")
        c = P.collect(cfg, cl, M.load(), lambda s: None)
        self.assertIn("HG2025", cl.calls); self.assertIn("HG2026H", cl.calls)
        self.assertTrue(any("응답 접수번호 X" in s for s in c.logs))

    def test_fiscal_numbers_recorded(self):
        fis = {k[1]: v for k, v in self.c.units.items() if k[0] == "fiscal"}
        self.assertTrue(fis.get("FY2025"))
        self.assertEqual(fis.get("FY2024"), fis["FY2025"] - 1)


class FlakyClient(FakeClient):
    def documents(self, corp_code, rcept_no):
        if rcept_no == "OB2024":
            raise IOError("다운로드 실패")
        return super().documents(corp_code, rcept_no)


class TestRobustness(unittest.TestCase):
    def setUp(self):
        self.cfg = P.BuildConfig("OB", "오비맥주", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1, {"recommended": True}, "default")

    def test_bad_report_does_not_abort(self):
        seen = []
        c = P.collect(self.cfg, FlakyClient(), M.load(), seen.append)
        self.assertTrue(any(r.period_key == "FY2025" for r in c.fs_records))
        bad = [s for s in c.logs if "처리 실패" in s and "2024.12" in s]
        self.assertTrue(bad)
        self.assertTrue(any("처리 실패" in s for s in seen))

    def test_toc_preview_still_raises(self):
        class Boom(FakeClient):
            def documents(self, *a):
                raise IOError("x")
        with self.assertRaises(IOError):
            P.toc_preview(self.cfg, Boom(), M.load(), lambda s: None)

    def test_fs_messages_in_col_logs_and_callback(self):
        seen = []
        c = P.collect(self.cfg, FakeClient(), M.load(), seen.append)
        self.assertTrue(any("원문" in s and "행" in s for s in c.logs))
        for s in c.logs:
            if "원문" in s and "행" in s:
                self.assertIn(s, seen)

    def test_from_dict_requires_as_of(self):
        with self.assertRaises(ValueError):
            P.BuildConfig.from_dict({"corp_code": "OB", "start_year": 2024, "end_year": 2025})


class FsRec:
    def __init__(self, pk, role, label, amount):
        self.period_key, self.role, self.label_norm, self.detail, self.amount = pk, role, label, "", amount
        self.fs_div, self.sj_div = "OFS", "BS"


class TestRestatedPick(unittest.TestCase):
    def _run(self, own_val):
        by = {"OB2025": [FsRec("FY2024", "prev", "현금", 100)], "OB2024": [FsRec("FY2024", "cur", "현금", own_val)]}
        src = {"FY2024": [("OB2025", "prev"), ("OB2024", "cur")]}
        return P._pick_restated(["FY2024"], src, by, lambda r: (r.fs_div, r.sj_div),
                                lambda r: (r.label_norm, r.detail), lambda r: (r.label_norm, r.detail))

    def test_diff_recorded(self):
        chosen, diffs = self._run(90)
        self.assertEqual([r.amount for r in chosen], [100])
        self.assertEqual(len(diffs), 1)
        self.assertEqual((diffs[0]["채택"], diffs[0]["원보고서"]), (100, 90))
        self.assertIn("현금", diffs[0]["key"])

    def test_no_diff_when_equal(self):
        chosen, diffs = self._run(100)
        self.assertEqual(len(chosen), 1)
        self.assertEqual(diffs, [])


class NoteRec:
    def __init__(self, tbl, row, col, v):
        self.period_key, self.role, self.fs_div, self.note_title = "FY2023", "cur", "OFS", "유형자산"
        self.table_idx, self.row_norm, self.row_label, self.col_label, self.value = tbl, row, row, col, v


class TestNoteRestateKey(unittest.TestCase):
    def test_table_index_shift_is_not_a_diff(self):
        """다음 연도 보고서는 전기 변동표가 표3, 담보표가 표4. 원보고서는 변동표 표1, 담보표 표3."""
        a = [NoteRec(3, "토지", "장부금액", 249), NoteRec(4, "토지", "장부금액", 151)]
        b = [NoteRec(1, "토지", "장부금액", 249), NoteRec(3, "토지", "장부금액", 151)]
        for r in a:
            r.role = "prev"
        chosen, diffs = P._pick_restated(["FY2023"], {"FY2023": [("A", "prev"), ("B", "cur")]}, {"A": a, "B": b},
                                         lambda r: (r.fs_div, r.note_title), P._note_cmp, lambda r: (r.row_label, r.col_label))
        self.assertEqual([r.value for r in chosen], [249, 151])
        self.assertEqual(diffs, [])


class TestFsRestateKey(unittest.TestCase):
    def test_same_label_different_account_order(self):
        def rec(pk, role, aid, v):
            r = FsRec(pk, role, "지배기업소유주지분", v); r.account_id = aid; return r
        a = [rec("FY2022", "prev", "NI", 58), rec("FY2022", "prev", "CI", 57)]
        b = [rec("FY2022", "cur", "CI", 57), rec("FY2022", "cur", "NI", 58)]
        _, diffs = P._pick_restated(["FY2022"], {"FY2022": [("A", "prev"), ("B", "cur")]}, {"A": a, "B": b},
                                    lambda r: (r.fs_div, r.sj_div), P._fs_cmp, lambda r: (r.label_norm, r.detail))
        self.assertEqual(diffs, [])


class TestAuditorAndRaw(unittest.TestCase):
    def test_auditor_row_ob(self):
        from core import xml_doc as X
        nodes = X.stream(X.load_xml(os.path.join(FX, "ob_2025_audit_ofs.xml")))
        class R:
            period_end = date(2025, 12, 31); report_nm = "감사보고서 (2025.12)"; rcept_no = "OB2025"
        row = P._auditor_row(R, "OFS", nodes)
        self.assertIn("회계법인", row[4])
        self.assertIn("공정하게", row[5])
        self.assertTrue(row[5].startswith("우리의 의견으로는"))

    def test_raw_cells_plain_labels(self):
        import re
        from core import xml_doc as X, note_parser as NP, selector as SEL
        root = X.load_xml(os.path.join(FX, "ob_2025_audit_ofs.xml"))
        nodes = X.stream(root)
        entries = NP.note_toc(nodes)
        n = 0
        for e in entries:
            for nt in NP.note_tables(nodes, e):
                for rl, cl, v in P._raw_cells(nt):
                    n += 1
                    self.assertFalse(re.match(r"^\d{2} ", rl), rl)
                    self.assertTrue(rl)
        self.assertGreater(n, 0)


if __name__ == "__main__":
    unittest.main()
