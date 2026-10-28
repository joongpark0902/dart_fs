import json, os, unittest
import xml.etree.ElementTree as ET
from core import fs_parser as F, xml_doc as X
from core.model import Period

FX = os.path.join(os.path.dirname(__file__), "fixtures")

def load_json(name):
    with open(os.path.join(FX, name), encoding="utf-8") as f:
        return json.load(f)["list"]

class TestApiRecords(unittest.TestCase):
    def test_annual_three_periods(self):
        items = load_json("hugel_2025_annual_cfs.json")
        periods = {"cur": Period.annual(2025, 25), "prev": Period.annual(2024, 24), "prev2": Period.annual(2023, 23)}
        recs = F.api_records(items, "CFS", periods, "20260320000349")
        bs_cur = [r for r in recs if r.sj_div == "BS" and r.period_key == "FY2025"]
        self.assertEqual(len(bs_cur), 43)
        assets = next(r for r in bs_cur if r.label == "자산총계")
        self.assertEqual(assets.amount, 1076837920124)
        self.assertEqual(assets.account_id, "ifrs-full_Assets")
        prev = next(r for r in recs if r.sj_div == "BS" and r.label == "자산총계" and r.period_key == "FY2024")
        self.assertEqual(prev.amount, 939439325841); self.assertEqual(prev.role, "prev")
        self.assertTrue(any(r.period_key == "FY2023" for r in recs))
        sce = [r for r in recs if r.sj_div == "SCE" and r.period_key == "FY2025"]
        self.assertTrue(all(r.detail for r in sce))
        self.assertIn("자본금", {r.detail for r in sce})
        self.assertIn("총계", {r.detail for r in sce})
        nonstd = next(r for r in recs if r.account_id is None)
        self.assertTrue(nonstd.label)

    def test_half_uses_cumulative_for_flows(self):
        items = load_json("hugel_2026_half_cfs.json")
        periods = {"cur": Period.half(2026, 26), "prev": Period.half(2025, 25), "prev_bs": Period.annual(2025, 25)}
        recs = F.api_records(items, "CFS", periods, "20260814001402")
        op = next(r for r in recs if r.sj_div == "CIS" and r.label == "영업이익" and r.period_key == "2026H1")
        self.assertEqual(op.amount, 103655427999)         # thstrm_add_amount(누적)
        op_prev = next(r for r in recs if r.sj_div == "CIS" and r.label == "영업이익" and r.period_key == "2025H1")
        self.assertEqual(op_prev.amount, 95623891825)     # frmtrm_add_amount
        bs_prev = [r for r in recs if r.sj_div == "BS" and r.period_key == "FY2025"]
        self.assertTrue(bs_prev)                           # 반기 BS 의 전기말 = FY2025 기말
        self.assertFalse(any(r.sj_div == "BS" and r.period_key == "2025H1" for r in recs))
        assets = next(r for r in bs_prev if r.account_id == "ifrs-full_Assets")
        self.assertEqual(assets.amount, 1076837920124)     # frmtrm_amount = FY2025 기말 (연간 파일과 동일)
        self.assertEqual(assets.role, "prev")

    def test_half_is_requires_cumulative_column_but_cf_may_fall_back(self):
        periods = {"cur": Period.half(2026, 26), "prev": Period.half(2025, 25), "prev_bs": Period.annual(2025, 25)}
        items = [
            {"sj_div": "IS", "account_id": "ifrs-full_Revenue", "account_nm": "매출액", "ord": "1",
             "thstrm_amount": "300", "thstrm_add_amount": "", "frmtrm_amount": "200", "frmtrm_add_amount": ""},
            {"sj_div": "CF", "account_id": "ifrs-full_CashFlowsFromUsedInOperatingActivities", "account_nm": "영업활동현금흐름", "ord": "2",
             "thstrm_amount": "700", "frmtrm_amount": "600"},
        ]
        recs = F.api_records(items, "CFS", periods, "r1")
        self.assertEqual([r for r in recs if r.sj_div == "IS"], [])
        cf = {r.period_key: r.amount for r in recs if r.sj_div == "CF"}
        self.assertEqual(cf, {"2026H1": 700, "2025H1": 600})

    def test_missing_cur_returns_empty(self):
        items = [{"sj_div": "BS", "account_nm": "자산총계", "ord": "1", "thstrm_amount": "1"}]
        self.assertEqual(F.api_records(items, "CFS", {"prev": Period.annual(2024, 24)}, "r"), [])

    def test_sce_detail_label(self):
        self.assertEqual(F.sce_detail_label("자본 [구성요소]|지배기업의 소유주에게 귀속되는 지분 [구성요소]|자본금 [구성요소]"), "자본금")
        self.assertEqual(F.sce_detail_label("연결재무제표 [member]"), "총계")
        self.assertEqual(F.sce_detail_label("자본 [구성요소]|비지배지분 [구성요소]"), "비지배지분")
        self.assertEqual(F.sce_detail_label("-"), "")
        self.assertEqual(F.sce_detail_label("별도재무제표 [member]"), "총계")
        # 2022 사업보고서 응답: 구성요소에도 [member]
        self.assertEqual(F.sce_detail_label("자본 [member]|지배기업 소유주 귀속분|자본금 [member]"), "자본금")
        self.assertEqual(F.sce_detail_label("자본 [member]|비지배지분 [member]"), "비지배지분")


class TestXmlStatements(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = X.load_xml(os.path.join(FX, "ob_2025_audit_ofs.xml"))
        cls.periods = {"cur": Period.annual(2025, 28), "prev": Period.annual(2024, 27)}
        cls.recs, cls.info = F.xml_statements(cls.root, "OFS", cls.periods, "20260410002383")

    def test_statements_found_with_units(self):
        self.assertEqual(set(r.sj_div for r in self.recs), {"BS", "CIS", "SCE", "CF"})
        self.assertEqual(self.info["units"]["BS"], 1)
        self.assertEqual(self.info["fiscal"], {"cur": 28, "prev": 27})

    def test_bs_rows_coalesce_two_columns_per_period(self):
        bs = [r for r in self.recs if r.sj_div == "BS" and r.period_key == "FY2025"]
        cash = next(r for r in bs if r.label == "1. 현금및현금성자산")
        self.assertEqual(cash.amount, 171311183566)
        self.assertEqual(cash.label_norm, "현금및현금성자산")
        ca = next(r for r in bs if r.label == "I. 유동자산")
        self.assertEqual(ca.amount, 523624393619)
        prev_cash = next(r for r in self.recs if r.sj_div == "BS" and r.label == "1. 현금및현금성자산" and r.period_key == "FY2024")
        self.assertEqual(prev_cash.amount, 206064972985)
        self.assertEqual(prev_cash.role, "prev")
        self.assertTrue(all(r.source == "XML" for r in bs))
        heading = next(r for r in bs if r.label == "자 산")
        self.assertIsNone(heading.amount)

    def test_bs_balances(self):
        bs = {r.label_norm: r.amount for r in self.recs if r.sj_div == "BS" and r.period_key == "FY2025"}
        self.assertEqual(bs["자산총계"], bs["부채총계"] + bs["자본총계"])

    def test_order_follows_document(self):
        bs = [r for r in self.recs if r.sj_div == "BS" and r.period_key == "FY2025"]
        self.assertEqual([r.ord for r in bs], sorted(r.ord for r in bs))
        self.assertLess(bs[0].ord, next(r for r in bs if r.label_norm == "자산총계").ord)

    def test_sce_rows_have_detail_and_period_from_dates(self):
        sce = [r for r in self.recs if r.sj_div == "SCE"]
        self.assertEqual(len(sce), 48)
        self.assertTrue(all(r.detail for r in sce))
        self.assertEqual({r.period_key for r in sce}, {"FY2025", "FY2024"})
        self.assertEqual({r.detail for r in sce}, {"자본금", "기타불입항목", "이익잉여금", "총계"})   # 총 계 정규화, 주석 열 제외
        open_row = next(r for r in sce if r.period_key == "FY2024" and r.label.startswith("2024.1.1") and r.detail == "총계")
        self.assertEqual(open_row.amount, 1486812638518)
        self.assertEqual(open_row.label_norm, "기초"); self.assertEqual(open_row.role, "prev")
        none_cell = next(r for r in sce if r.label == "주식기준보상거래의 인식" and r.detail == "자본금")
        self.assertIsNone(none_cell.amount)
        neg = next(r for r in sce if r.label == "배당금지급" and r.detail == "이익잉여금")
        self.assertEqual(neg.amount, -332800000000)
        self.assertIn("기말", {r.label_norm for r in sce})

    def test_cf_has_values(self):
        cf = [r for r in self.recs if r.sj_div == "CF" and r.period_key == "FY2025" and r.amount is not None]
        self.assertGreater(len(cf), 10)
        op = next(r for r in cf if r.label_norm == "영업활동으로인한현금흐름")
        self.assertEqual(op.amount, 275285140680)
        op_prev = next(r for r in self.recs if r.sj_div == "CF" and r.period_key == "FY2024" and r.label_norm == "영업활동으로인한현금흐름")
        self.assertEqual(op_prev.amount, 430997500583)


class TestXmlWarnings(unittest.TestCase):
    XML = ("<DOCUMENT><BODY><TITLE>재무제표</TITLE>"
           "<TABLE><TR><TD>재 무 상 태 표</TD></TR><TR><TD>제 1 (당)기말 : 2025년 12월 31일 현재</TD></TR></TABLE>"
           "<TABLE><THEAD><TR><TH>과 목</TH><TH>금액</TH></TR></THEAD>"
           "<TBODY><TR><TD>자산</TD><TD>100</TD></TR></TBODY></TABLE></BODY></DOCUMENT>")

    def test_unparsed_body_warns(self):
        recs, info = F.xml_statements(ET.fromstring(self.XML), "OFS", {"cur": Period.annual(2025, 1)}, "r")
        self.assertEqual(recs, [])
        w = " | ".join(info["warnings"])
        self.assertIn("재무상태표: 본문 표를 해석하지 못함", w)
        self.assertIn("기간 열", w)
        self.assertIn("재무제표 섹션에서 표를 찾지 못함", w)


class TestOrphanBodyTables(unittest.TestCase):
    TITLE = ("<TABLE><TR><TD>재 무 상 태 표</TD></TR><TR><TD>제 1 (당)기말 : 2025년 12월 31일 현재</TD></TR></TABLE>")
    BODY = ('<TABLE BORDER="1"><THEAD><TR><TH>과 목</TH><TH>제 1 (당) 기말</TH></TR></THEAD>'
            "<TBODY><TR><TD>{0}</TD><TD>100</TD></TR></TBODY></TABLE>")

    def _run(self, body):
        xml = f"<DOCUMENT><BODY><TITLE>재무제표</TITLE>{body}</BODY></DOCUMENT>"
        return F.xml_statements(ET.fromstring(xml), "OFS", {"cur": Period.annual(2025, 1)}, "r")

    def test_body_before_any_title_warns(self):
        recs, info = self._run(self.BODY.format("고아행") + self.TITLE + self.BODY.format("자산총계"))
        self.assertEqual([r.label for r in recs], ["자산총계"])
        w = [m for m in info["warnings"] if "제목표 없는 본문표" in m]
        self.assertEqual(len(w), 1)
        self.assertIn("첫 제목표 앞", w[0]); self.assertIn("고아행", w[0])

    def test_second_body_after_records_warns_absorbed(self):
        recs, info = self._run(self.TITLE + self.BODY.format("자산총계") + self.BODY.format("흡수될행"))
        self.assertEqual([r.label for r in recs], ["자산총계", "흡수될행"])       # 동작은 그대로(흡수), 경고만
        w = [m for m in info["warnings"] if "제목표 없는 본문표" in m]
        self.assertEqual(len(w), 1)
        self.assertIn("재무상태표 뒤", w[0]); self.assertIn("이전 재무제표로 흡수될 수 있음", w[0])

    def test_normal_statement_no_warning(self):
        _, info = self._run(self.TITLE + self.BODY.format("자산총계"))
        self.assertFalse([m for m in info["warnings"] if "제목표 없는 본문표" in m])


class TestFiscalOnlyHeaders(unittest.TestCase):
    """첫 감사 연도 원문(2023): 제목표에 (당)(전) 없음, 본문 헤더 '제 26 기말 / 제 25 기말(감사받지 않은…) / 제 25 기초(…)'."""
    def test_title_without_markers_takes_first_fiscal_as_cur(self):
        g = X.Grid(body=[["재 무 상 태 표", "재 무 상 태 표"], ["제 26 기말 2023년 12월 31일 현재", "제 26 기말 2023년 12월 31일 현재"]], ncols=2)
        self.assertEqual(F._title_table(g), ("BS", 1, {"cur": 26}))

    def test_groups_derive_prev_and_opening_is_prior_year_end(self):
        hdr = ["과 목", "주석", "제 26 기말", "제 26 기말", "제 25 기말(감사받지 않은 재무제표)", "제 25 기말(감사받지 않은 재무제표)",
               "제 25 기초(감사받지 않은 재무제표)", "제 25 기초(감사받지 않은 재무제표)"]
        per = {"cur": Period.annual(2023, 26), "prev": Period.annual(2022, 25), "prev2": Period.annual(2021, 24)}
        got = [(cols, p.key, role) for cols, p, role in F._period_groups(hdr, per, {"cur": 26})]
        self.assertEqual(got, [([2, 3], "FY2023", "cur"), ([4, 5], "FY2022", "prev"), ([6, 7], "FY2021", "prev2")])


if __name__ == "__main__":
    unittest.main()
