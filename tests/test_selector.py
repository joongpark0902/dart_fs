import unittest
from datetime import date
from core import selector as S

D = [  # 비상장 F 단독공시 + 정정본
    {"rcept_no": "20240405000001", "report_nm": "감사보고서 (2023.12)", "rcept_dt": "20240405"},
    {"rcept_no": "20250404002666", "report_nm": "감사보고서 (2024.12)", "rcept_dt": "20250404"},
    {"rcept_no": "20250404002667", "report_nm": "연결감사보고서 (2024.12)", "rcept_dt": "20250404"},
    {"rcept_no": "20260410002383", "report_nm": "감사보고서 (2025.12)", "rcept_dt": "20260410"},
    {"rcept_no": "20260420000001", "report_nm": "[기재정정]감사보고서 (2025.12)", "rcept_dt": "20260420"},
    {"rcept_no": "20260410002385", "report_nm": "연결감사보고서 (2025.12)", "rcept_dt": "20260410"},
]
L = [  # 상장
    {"rcept_no": "20250321000621", "report_nm": "사업보고서 (2024.12)", "rcept_dt": "20250321"},
    {"rcept_no": "20250814002269", "report_nm": "반기보고서 (2025.06)", "rcept_dt": "20250814"},
    {"rcept_no": "20260320000349", "report_nm": "사업보고서 (2025.12)", "rcept_dt": "20260320"},
    {"rcept_no": "20260814001402", "report_nm": "반기보고서 (2026.06)", "rcept_dt": "20260814"},
    {"rcept_no": "20260515000001", "report_nm": "분기보고서 (2026.03)", "rcept_dt": "20260515"},
]

class TestParse(unittest.TestCase):
    def test_parse(self):
        self.assertEqual(S.parse_report_nm("[기재정정]감사보고서 (2025.12)"), ("감사보고서", date(2025, 12, 31), True))
        self.assertEqual(S.parse_report_nm("연결감사보고서 (2024.12)"), ("연결감사보고서", date(2024, 12, 31), False))
        self.assertEqual(S.parse_report_nm("반기보고서 (2026.06)"), ("반기보고서", date(2026, 6, 30), False))
        self.assertIsNone(S.parse_report_nm("분기보고서 (2026.03)"))
        self.assertIsNone(S.parse_report_nm("주요사항보고서(유상증자결정)"))

class TestPlanUnlisted(unittest.TestCase):
    def test_correction_wins_and_restated_sources(self):
        p = S.plan_reports(D, 2023, 2025, date(2026, 10, 26), False, {"OFS"})
        self.assertEqual([x.key for x in p.periods], ["FY2023", "FY2024", "FY2025"])
        r25 = next(r for r in p.reports.values() if r.period_end == date(2025, 12, 31) and r.rtype == "감사보고서")
        self.assertEqual(r25.rcept_no, "20260420000001"); self.assertTrue(r25.corrected)
        self.assertEqual(r25.superseded, ("20260410002383",))
        self.assertEqual(p.sources["FY2024"][0], ("20260420000001", "prev"))
        self.assertEqual(p.sources["FY2024"][1], ("20250404002666", "cur"))
        self.assertEqual(p.sources["FY2025"], [("20260420000001", "cur")])
        self.assertEqual(p.base_report.rcept_no, "20260420000001")
        self.assertNotIn("20250404002667", p.reports)      # CFS 미선택이면 연결감사보고서 제외

    def test_as_of_cuts_latest(self):
        p = S.plan_reports(D, 2023, 2025, date(2026, 3, 31), False, {"OFS"})
        self.assertEqual([x.key for x in p.periods], ["FY2023", "FY2024"])
        self.assertEqual(p.base_report.period_end, date(2024, 12, 31))
        self.assertEqual(p.sources["FY2024"], [("20250404002666", "cur")])

class TestPlanListed(unittest.TestCase):
    def test_half_periods_and_sources(self):
        p = S.plan_reports(L, 2024, 2026, date(2026, 10, 26), True, {"CFS"})
        self.assertEqual([x.key for x in p.periods], ["FY2024", "FY2025", "2024H1", "2025H1", "2026H1"])
        self.assertEqual(p.sources["2024H1"], [("20250814002269", "prev")])
        self.assertEqual(p.sources["2025H1"], [("20260814001402", "prev"), ("20250814002269", "cur")])
        self.assertEqual(p.sources["2026H1"], [("20260814001402", "cur")])
        self.assertEqual(p.sources["FY2024"][0], ("20260320000349", "prev"))
        self.assertNotIn("2026H1", p.note_sources)
        self.assertEqual(p.note_sources["FY2024"][0], ("20260320000349", "prev"))
        self.assertTrue(all(r.rtype != "분기보고서" for r in p.reports.values()))

    def test_fiscal_periods_for(self):
        p = S.plan_reports(L, 2024, 2026, date(2026, 10, 26), True, {"CFS"})
        annual = p.reports["20260320000349"]
        fp = S.fiscal_periods_for(annual, 25)
        self.assertEqual((fp["cur"].key, fp["prev"].key, fp["prev2"].key), ("FY2025", "FY2024", "FY2023"))
        self.assertEqual(fp["prev"].fiscal_no, 24)
        half = p.reports["20260814001402"]
        hp = S.fiscal_periods_for(half, 26)
        self.assertEqual((hp["cur"].key, hp["prev"].key, hp["prev_bs"].key), ("2026H1", "2025H1", "FY2025"))

class TestReview2(unittest.TestCase):
    ASOF = date(2026, 10, 26)

    def test_multi_bracket_parse(self):
        self.assertEqual(S.parse_report_nm("[기재정정][첨부추가]사업보고서 (2024.12)"), ("사업보고서", date(2024, 12, 31), True))

    def test_bad_month_none(self):
        self.assertIsNone(S.parse_report_nm("사업보고서 (2024.13)"))
        self.assertIsNone(S.parse_report_nm("사업보고서 (2024.00)"))

    def test_bad_year_none_and_logged(self):
        self.assertIsNone(S.parse_report_nm("감사보고서 (0000.12)"))      # date(0, …) ValueError 대신 None
        ds = [{"rcept_no": "X1", "report_nm": "감사보고서 (0000.12)", "rcept_dt": "20250101"},
              {"rcept_no": "X2", "report_nm": "주요사항보고서(유상증자결정)", "rcept_dt": "20250101"},
              {"rcept_no": "X3", "report_nm": "분기보고서 (2025.03)", "rcept_dt": "20250515"},
              {"rcept_no": "A1", "report_nm": "감사보고서 (2025.12)", "rcept_dt": "20260401"}]
        plan = S.plan_reports(ds, 2025, 2025, date(2026, 10, 26), False, {"OFS"})
        skipped = [m for m in plan.log if "해석하지 못해 제외" in m]
        self.assertEqual(len(skipped), 1)
        self.assertIn("X1", skipped[0])
        self.assertIn("A1", plan.reports)

    def test_non_december_annual_excluded(self):
        d = [{"rcept_no": "20250601000001", "report_nm": "사업보고서 (2025.03)", "rcept_dt": "20250601"},
             {"rcept_no": "20250321000621", "report_nm": "사업보고서 (2024.12)", "rcept_dt": "20250321"}]
        p = S.plan_reports(d, 2024, 2025, self.ASOF, False, {"CFS"})
        self.assertEqual(p.base_report.rcept_no, "20250321000621")
        self.assertNotIn("20250601000001", p.reports)
        self.assertTrue(any("2025.03" in l and "제외" in l for l in p.log))

    def test_non_june_half_excluded(self):
        d = [{"rcept_no": "20251201000001", "report_nm": "반기보고서 (2025.09)", "rcept_dt": "20251201"},
             {"rcept_no": "20250321000621", "report_nm": "사업보고서 (2024.12)", "rcept_dt": "20250321"}]
        p = S.plan_reports(d, 2024, 2025, self.ASOF, True, {"CFS"})
        self.assertNotIn("20251201000001", p.reports)
        self.assertTrue(any("2025.09" in l and "제외" in l for l in p.log))

    def test_same_period_priority_사업_before_감사(self):
        d = [{"rcept_no": "20250321000001", "report_nm": "사업보고서 (2024.12)", "rcept_dt": "20250321"},
             {"rcept_no": "20250404000002", "report_nm": "감사보고서 (2024.12)", "rcept_dt": "20250404"},
             {"rcept_no": "20260320000003", "report_nm": "사업보고서 (2025.12)", "rcept_dt": "20260320"}]
        p = S.plan_reports(d, 2024, 2025, self.ASOF, False, {"OFS"})
        self.assertEqual(p.sources["FY2024"], [("20260320000003", "prev"), ("20250321000001", "cur"), ("20250404000002", "cur")])

    def test_correction_after_as_of_ignored(self):
        d = [{"rcept_no": "20250404000001", "report_nm": "감사보고서 (2024.12)", "rcept_dt": "20250404"},
             {"rcept_no": "20260501000002", "report_nm": "[기재정정]감사보고서 (2024.12)", "rcept_dt": "20260501"}]
        p = S.plan_reports(d, 2024, 2024, date(2026, 3, 31), False, {"OFS"})
        r = p.reports["20250404000001"]
        self.assertFalse(r.corrected); self.assertEqual(r.superseded, ())
        self.assertNotIn("20260501000002", p.reports)

    def test_include_half_false(self):
        p = S.plan_reports(L, 2024, 2026, self.ASOF, False, {"CFS"})
        self.assertFalse(any(x.key.endswith("H1") for x in p.periods))
        self.assertTrue(all(r.kind != "half" for r in p.reports.values()))

    def test_pruned_report_logged(self):
        d = [{"rcept_no": "20240405000001", "report_nm": "감사보고서 (2023.12)", "rcept_dt": "20240405"},
             {"rcept_no": "20250404000002", "report_nm": "감사보고서 (2024.12)", "rcept_dt": "20250404"}]
        p = S.plan_reports(d, 2024, 2024, self.ASOF, False, {"OFS"})
        self.assertNotIn("20240405000001", p.reports)
        self.assertIn("감사보고서 (2023.12): 사용 안 함(기간 범위 밖)", p.log)

if __name__ == "__main__":
    unittest.main()
