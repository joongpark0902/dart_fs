import os, tempfile, unittest
from core import mapping as M, note_std as S
from core.model import NoteEntry

class TestMapping(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.m = M.load()

    def test_id_wins(self):
        self.assertEqual(self.m.map_account("자산 총계(주석3)", "ifrs-full_Assets", "BS"), ("자산총계", "id"))

    def test_keyword_exact_then_startswith_then_contains(self):
        self.assertEqual(self.m.map_account("1. 현금및현금성자산", None, "BS"), ("현금및현금성자산", "kw"))
        self.assertEqual(self.m.map_account("매출채권 및 기타채권", None, "BS"), ("매출채권및기타채권", "kw"))
        self.assertEqual(self.m.map_account("수익(매출액)", None, "CIS"), ("매출액", "kw"))
        self.assertEqual(self.m.map_account("판매비, 관리비와 물류비", None, "IS"), ("판매비와관리비", "kw"))

    def test_sga_with_logistics_label(self):
        self.assertEqual(self.m.map_account("IV. 판매비와 관리비 및 물류비", None, "CIS"), ("판매비와관리비", "kw"))
        self.assertEqual(self.m.map_account("II. 매출원가", None, "CIS")[0], "매출원가")

    def test_unmapped(self):
        self.assertEqual(self.m.map_account("용기지급수수료", None, "IS"), (None, "none"))
        self.assertEqual(self.m.map_account("자산", None, "BS"), (None, "none"))   # 섹션 제목은 정확일치 아님

    def test_tie_is_unmapped(self):
        # 같은 prio 에 서로 다른 std 가 '포함' 으로만 걸리면 미매핑
        self.assertEqual(self.m.map_account("총유동자산및비유동자산", None, "BS"), (None, "none"))  # 유동자산·비유동자산 둘 다 (포함, prio1) 동률
        # 결합 라벨은 잔여 규칙으로 부분일치 불가
        self.assertEqual(self.m.map_account("단기차입금및유동성장기부채", None, "BS"), (None, "none"))

    def test_task5_keywords(self):
        """5-3: 손실·손익 표기 변형(로마숫자·번호 접두 포함)이 같은 표준계정으로. 비예시는 그대로 미매핑."""
        m = self.m
        for lab, sj, std in [("Ⅶ. 당기순손실", "IS", "당기순이익"), ("당기순이익(손실)", "CIS", "당기순이익"), ("당기순손익", "IS", "당기순이익"),
                             ("당기순손실(이익)", "IS", "당기순이익"), ("Ⅰ. 당기순손실", "IS", "당기순이익"),
                             ("Ⅴ. 영업손실", "IS", "영업이익"), ("영업이익(손실)", "IS", "영업이익"), ("영업손익", "IS", "영업이익"),
                             ("Ⅵ. 법인세비용차감전 순손실", "IS", "법인세비용차감전순이익"), ("법인세비용차감전순이익(손실)", "IS", "법인세비용차감전순이익"),
                             ("법인세비용차감전순손익", "IS", "법인세비용차감전순이익"), ("법인세차감전순이익", "IS", "법인세비용차감전순이익"),
                             ("법인세차감전순손실", "IS", "법인세비용차감전순이익"),
                             ("법인세수익", "IS", "법인세비용"), ("법인세수익(비용)", "IS", "법인세비용"), ("법인세비용(수익)", "IS", "법인세비용"),
                             ("Ⅱ. 당기기타포괄손익", "IS", "기타포괄손익"), ("기타포괄손익(손실)", "IS", "기타포괄손익"),
                             ("Ⅲ. 당기총포괄손익", "IS", "총포괄이익"), ("Ⅲ. 당기총포괄손실", "IS", "총포괄이익"), ("총포괄손실", "IS", "총포괄이익"),
                             ("총포괄손익", "IS", "총포괄이익"), ("당기총포괄이익", "IS", "총포괄이익"), ("총포괄이익(손실)", "IS", "총포괄이익"),
                             ("지배기업의 소유주지분", "BS", "지배기업소유주지분"), ("지배기업소유주지분", "BS", "지배기업소유주지분"),
                             ("지배기업의소유주에게귀속되는자본", "BS", "지배기업소유주지분"), ("지배기업소유주귀속지분", "BS", "지배기업소유주지분"),
                             ("지배주주지분", "BS", "지배기업소유주지분"),
                             ("Ⅴ. 결손금", "BS", "이익잉여금"), ("미처리결손금", "BS", "이익잉여금"), ("이익잉여금(결손금)", "BS", "이익잉여금"),
                             ("유동성장기차입금", "BS", "유동성장기부채")]:
            self.assertEqual(m.map_account(lab, None, sj), (std, "kw"), lab)
        for lab, sj in [("당기순이익에대한법인세", "IS"), ("지배기업 소유주지분", "IS"), ("비지배지분", "IS"), ("기본주당손실", "IS"),
                        ("반기기타포괄손익", "IS"), ("법인세비용차감후기타포괄손익", "IS"), ("이월결손금", "BS"), ("세무상결손금", "BS"),
                        ("미처리결손금의 처리", "BS"), ("- 지분법자본변동", "IS"), ("영업손실에 대한 조정", "IS"), ("법인세비용차감전순손실에 대한 조정", "IS")]:
            self.assertEqual(m.map_account(lab, None, sj), (None, "none"), lab)

    def test_partial_match_residual_rules(self):
        for lab, sj in [("계속영업이익", "IS"), ("상품매출", "IS"),       # 5-3: '법인세비용차감전순손익' 은 이제 키워드(매핑됨)
                        ("기타유동부채", "BS"), ("장기매출채권", "BS"), ("유동성리스부채", "BS"),
                        ("영업권이외의무형자산", "BS")]:
            self.assertEqual(self.m.map_account(lab, None, sj), (None, "none"), lab)
        self.assertEqual(self.m.map_account("법인세비용차감전순이익(손실)", None, "IS"), ("법인세비용차감전순이익", "kw"))
        self.assertEqual(self.m.map_account("자산총계", None, "BS")[0], "자산총계")
        self.assertEqual(self.m.map_account("유동자산합계", None, "BS")[0], "유동자산")

    def test_task2_new_standard_accounts(self):
        """2-1: 기타포괄손익(IS)·환율변동효과·현금및현금성자산의증가(CF)·충당부채·비유동충당부채(BS)."""
        m = self.m
        for lab, sj, std in [("X. 기타포괄손익", "CIS", "기타포괄손익"), ("기타포괄이익", "IS", "기타포괄손익"),
                             ("기타포괄손익(손실)", "IS", "기타포괄손익"), ("세후기타포괄손익", "IS", "기타포괄손익"),
                             ("Ⅳ. 현금및현금성자산에 대한 환율변동효과", "CF", "환율변동효과"),
                             ("VI. 외화표시현금및현금성자산의환율변동효과", "CF", "환율변동효과"),
                             ("환율변동효과", "CF", "환율변동효과"), ("현금및현금성자산의 환율변동효과", "CF", "환율변동효과"),
                             ("V. 현금및현금성자산의 증가(I+II+III+Ⅳ)", "CF", "현금및현금성자산의증가"),
                             ("현금및현금성자산의 순증가(감소)", "CF", "현금및현금성자산의증가"),
                             ("현금및현금성자산의 증가(감소)", "CF", "현금및현금성자산의증가"),
                             ("현금및현금성자산의감소", "CF", "현금및현금성자산의증가"),
                             ("현금의 증가(감소)", "CF", "현금및현금성자산의증가"),
                             ("비유동충당부채", "BS", "비유동충당부채"), ("장기충당부채", "BS", "비유동충당부채"),
                             ("충당부채", "BS", "충당부채")]:
            self.assertEqual(m.map_account(lab, None, sj), (std, "kw"), lab)
        self.assertEqual(m.map_account("x", "ifrs-full_OtherComprehensiveIncome", "IS"), ("기타포괄손익", "id"))
        self.assertEqual(m.map_account("x", "ifrs-full_EffectOfExchangeRateChangesOnCashAndCashEquivalents", "CF"), ("환율변동효과", "id"))
        self.assertEqual(m.map_account("x", "ifrs-full_IncreaseDecreaseInCashAndCashEquivalents", "CF"), ("현금및현금성자산의증가", "id"))
        self.assertEqual(m.map_account("x", "ifrs-full_Provisions", "BS"), ("충당부채", "id"))
        self.assertEqual(m.map_account("x", "ifrs-full_NoncurrentProvisions", "BS"), ("비유동충당부채", "id"))
        self.assertEqual(m.std_order[("IS", "기타포괄손익")], 190)
        self.assertEqual(m.std_order[("CF", "환율변동효과")], 70)
        self.assertEqual(m.std_order[("CF", "현금및현금성자산의증가")], 75)
        self.assertEqual(m.std_order[("BS", "비유동충당부채")], 325)
        self.assertEqual(m.std_order[("BS", "충당부채")], 246)

    def test_task2_no_new_mismaps(self):
        m = self.m
        for lab, sj in [("기타유동부채", "BS"), ("기타포괄손익누계액", "BS"), ("충당부채", "IS"),
                        ("환율변동효과 반영전 현금및현금성자산의 순증가(감소)", "CF"),
                        ("기타포괄손익-공정가치측정금융자산 평가손익", "IS"), ("반기기타포괄손익", "IS"),
                        ("법인세비용차감후기타포괄손익", "IS"), ("복구충당부채", "BS"), ("기타충당부채", "BS")]:
            self.assertEqual(m.map_account(lab, None, sj), (None, "none"), lab)
        self.assertEqual(m.map_account("유동충당부채", None, "BS"), ("유동충당부채", "kw"))

    def test_save_from_rows_numeric_cells(self):
        rows = [["종류", "sj", "키", "표준계정/주제", "우선순위/순서", "시트", "추천"],
                ["표준코드", "BS", "x_id", "테스트", 41.0, "", ""],
                ["키워드", "BS", " 테스트 ", " 테스트 ", " 1 ", "", ""],
                ["키워드", "BS", "테스트2", "테스트", None, "", ""]]
        with tempfile.TemporaryDirectory() as d:
            M.save_from_rows(rows, d)
            m2 = M.load(d)
            self.assertEqual(m2.std_order[("BS", "테스트")], 41)
            self.assertEqual([k[3] for k in m2.keywords], [1, 9])
            self.assertEqual(m2.keywords[0][:2], ("BS", "테스트"))

    def test_rows_roundtrip(self):
        rows = self.m.rows_for_sheet()
        self.assertEqual(rows[0], ["종류", "sj", "키", "표준계정/주제", "우선순위/순서", "시트", "추천"])
        with tempfile.TemporaryDirectory() as d:
            M.save_from_rows(rows, d)
            m2 = M.load(d)
            self.assertEqual(m2.rows_for_sheet(), self.m.rows_for_sheet())
            self.assertEqual(len(m2.keywords), len(self.m.keywords))
            self.assertEqual(len(m2.topics), len(self.m.topics))
            self.assertEqual(m2.std_by_id, self.m.std_by_id)

class TestTopics(unittest.TestCase):
    def test_detect(self):
        m = M.load()
        self.assertEqual(S.detect_topic("판매비관리비와물류비", m), ("판관비", True))
        self.assertEqual(S.detect_topic("비용의성격별분류", m), ("성격별비용", True))
        self.assertEqual(S.detect_topic("범주별금융손익", m), ("금융손익", True))
        self.assertEqual(S.detect_topic("유형자산", m), ("유형자산", True))
        self.assertEqual(S.detect_topic("차입금및사채", m), ("차입금", True))
        self.assertEqual(S.detect_topic("고객과의계약에서생기는수익", m), ("매출구분", True))
        for t, exp in [("매출채권및기타채권", (None, False)), ("기타수익및기타비용", (None, False)),
                       ("이자수익", (None, False)), ("매출원가", (None, False)), ("수익", ("매출구분", True)),
                       ("금융리스크관리", (None, False)), ("리스", ("리스", True)), ("리스부채", ("리스", True)),
                       ("자본", ("자본", True))]:
            self.assertEqual(S.detect_topic(t, m), exp, t)
        self.assertEqual(S.detect_topic("회사의개요", m), (None, False))
        self.assertEqual(S.detect_topic("보고기간후사건", m), (None, False))

    def test_assign(self):
        m = M.load()
        es = [NoteEntry("11", "유형자산", "유형자산", None, False, 0, 1), NoteEntry("1", "회사의 개요", "회사의개요", None, False, 0, 1)]
        S.assign_topics(es, m)
        self.assertEqual((es[0].topic, es[0].recommended, es[0].include), ("유형자산", True, True))
        self.assertEqual((es[1].topic, es[1].recommended, es[1].include), (None, False, True))

if __name__ == "__main__":
    unittest.main()
