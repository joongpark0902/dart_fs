import unittest
from core import normalize as N

class TestNormalize(unittest.TestCase):
    def test_clean_text_collapses_fullwidth_and_sentinel(self):
        self.assertEqual(N.clean_text("\u3000\u3000현금및\u2028현금성자산 "), "현금및 현금성자산")

    def test_norm_label_strips_numbering_and_note_refs(self):
        self.assertEqual(N.norm_label("1. 현금및현금성자산"), "현금및현금성자산")
        self.assertEqual(N.norm_label("I. 유동자산"), "유동자산")
        self.assertEqual(N.norm_label("Ⅱ. 비유동자산"), "비유동자산")
        self.assertEqual(N.norm_label("취득(주1)"), "취득")
        self.assertEqual(N.norm_label("취득(*)(주1)"), "취득")
        self.assertEqual(N.norm_label("기타포괄손익-공정가치 측정 금융자산"), "기타포괄손익공정가치측정금융자산")
        self.assertEqual(N.norm_label("매출채권 및 기타채권"), N.norm_label("매출채권및기타채권"))
        self.assertEqual(N.norm_label("합 계"), "합계")

    def test_norm_label_numbered_footmarks(self):
        """'기타(*)'(FY2022)와 '기타 (*1)'(FY2023~)는 같은 행 — 연도 병합 블록에서 갈리면 C-변동이 거짓 FALSE."""
        self.assertEqual(N.norm_label("기타 (*1)"), N.norm_label("기타(*)"))
        self.assertEqual(N.norm_label("총 비용( *2 )"), "총비용")
        self.assertEqual(N.norm_label("우선주(**)"), "우선주")
        self.assertEqual(N.norm_label("(주)한국"), "(주)한국")

    def test_parse_amount(self):
        self.assertEqual(N.parse_amount("171,311,183,566"), 171311183566)
        self.assertEqual(N.parse_amount("(148,045,701)"), -148045701)
        self.assertEqual(N.parse_amount("△1,234"), -1234)
        self.assertEqual(N.parse_amount("-1,234"), -1234)
        self.assertIsNone(N.parse_amount("-"))
        self.assertIsNone(N.parse_amount(""))
        self.assertIsNone(N.parse_amount("동청주세무서 외"))
        self.assertEqual(N.parse_amount("2010272.504"), 2010272.504)
        self.assertEqual(N.parse_amount("1,372,755,481 "), 1372755481)

    def test_parse_unit(self):
        self.assertEqual(N.parse_unit("(단위 : 원)"), 1)
        self.assertEqual(N.parse_unit("(단위: 천원)"), 1000)
        self.assertEqual(N.parse_unit("(당기) (단위 : 천 원)"), 1000)
        self.assertEqual(N.parse_unit("(단위: 백만원)"), 1_000_000)
        self.assertIsNone(N.parse_unit("구 분"))

    def test_split_heading_cuts_body_glued_to_title(self):
        self.assertEqual(N.split_heading("9. 재고자산당기말 및 전기말 현재 재고자산의 내역은 다음과 같습니다.")[:2], ("9", "재고자산"))
        self.assertEqual(N.split_heading("1. 회사의 개요")[:2], ("1", "회사의 개요"))
        self.assertEqual(N.split_heading("12. 리스회사가 리스이용자인 경우의 리스에 대한 정보는 다음과 같습니다.")[:2], ("12", "리스"))
        self.assertEqual(N.split_heading("18. 퇴직급여제도 (1) 확정기여형퇴직급여제도회사는")[:2], ("18", "퇴직급여제도"))
        self.assertEqual(N.split_heading("34. 보고기간 후 사건회사는 유동성 확보를 위해")[:2], ("34", "보고기간 후 사건"))
        self.assertEqual(N.split_heading("26. 비용의 성격별 분류당기 및 전기 중 발생한")[:2], ("26", "비용의 성격별 분류"))
        self.assertEqual(N.split_heading("2. 재무제표 작성기준 및 중요한 회계정책다음은 재무제표 작성에")[:2], ("2", "재무제표 작성기준 및 중요한 회계정책"))
        self.assertEqual(N.split_heading("35. 재무제표의 승인동 재무제표는 2026년 2월 11일자")[:2], ("35", "재무제표의 승인"))
        self.assertEqual(N.split_heading("27. 판매비, 관리비와 물류비")[:2], ("27", "판매비, 관리비와 물류비"))
        self.assertIsNone(N.split_heading("2.2 회계정책과 공시의 변경"))
        self.assertIsNone(N.split_heading("(1) 매출채권 및 기타채권"))
        self.assertIsNone(N.split_heading("2025년 12월 31일"))

    def test_period_role_and_fiscal_no(self):
        self.assertEqual(N.period_role("제 28(당) 기말"), "cur")
        self.assertEqual(N.period_role("제 27 (전)기"), "prev")
        self.assertEqual(N.period_role("당기말"), "cur")
        self.assertEqual(N.period_role("전기"), "prev")
        self.assertEqual(N.period_role("전전기"), "prev2")
        self.assertEqual(N.period_role("(당기)"), "cur")
        self.assertIsNone(N.period_role("구 분"))
        self.assertIsNone(N.period_role("당기손익-공정가치측정금융자산"))
        self.assertEqual(N.fiscal_no("제 28(당) 기말"), 28)
        self.assertEqual(N.fiscal_no("제 26 기 반기"), 26)
        self.assertIsNone(N.fiscal_no("당기"))

    def test_review_fixes(self):
        # 1 heading
        self.assertEqual(N.split_heading("12. 순확정급여부채 변동 내역")[:2], ("12", "순확정급여부채 변동 내역"))
        self.assertEqual(N.split_heading("35. 재무제표의 승인동 재무제표는 2026년")[:2], ("35", "재무제표의 승인"))
        # 2 unit
        self.assertEqual(N.parse_unit("(단위: 만원)"), 10000)
        self.assertEqual(N.parse_unit("(단위: 천만원)"), 10**7)
        self.assertEqual(N.parse_unit("(단위: 백억원)"), 10**10)
        self.assertEqual(N.parse_unit("(단위: 백만 원)"), 10**6)
        self.assertIsNone(N.parse_unit("(단위: 달러)"))
        # 3 period role
        self.assertEqual(N.period_role("제 28 기 (당기)"), "cur")
        self.assertEqual(N.period_role("(전)기"), "prev")
        self.assertEqual(N.period_role("(당)기"), "cur")
        for t, r in [("당분기", "cur"), ("당분기말", "cur"), ("전분기", "prev"), ("전분기말", "prev")]:
            self.assertEqual(N.period_role(t), r, t)
        # 4 note refs: 주로 시작하는 일반 괄호는 보존(괄호 포함 그대로, 공백만 제거)
        self.assertEqual(N.norm_label("주식선택권(주식결제형)"), "주식선택권(주식결제형)")
        self.assertEqual(N.norm_label("(주)한국"), "(주)한국")
        self.assertEqual(N.norm_label("취득(주석 3)"), "취득")
        # minor
        self.assertEqual(N.parse_amount(0), 0)
        self.assertEqual(N.parse_amount(12.5), 12.5)

if __name__ == "__main__":
    unittest.main()
