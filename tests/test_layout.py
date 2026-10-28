import unittest
from core import layout as L
from core.model import FsRecord, NoteRecord, Block, Row

def fs(pk, label, amt, sj="BS", ord_=0, unit=1, detail=None):
    from core import normalize as N
    return FsRecord(pk, "OFS", sj, ord_, label, N.norm_label(label), amt, unit, detail, None, "r", "XML", "cur")

def nr(pk, title, cap, row, col, val, unit=1000, no="11", idx=0):
    from core import normalize as N
    return NoteRecord(pk, "OFS", no, title, "유형자산", idx, cap, row, N.norm_label(row), col, val, unit, "r")

class TestJoinRows(unittest.TestCase):
    def test_latest_order_and_insert_old_only_rows(self):
        groups = {
            "FY2024": [("1. 현금", "현금", 1), ("2. 단기금융상품", "단기금융상품", 2), ("3. 재고자산", "재고자산", 3)],
            "FY2025": [("1. 현금", "현금", 10), ("2. 재고자산", "재고자산", 30), ("3. 기타", "기타", 40)],
        }
        rows = L.join_rows(groups, ["FY2024", "FY2025"])
        self.assertEqual([r.label for r in rows], ["1. 현금", "2. 단기금융상품", "2. 재고자산", "3. 기타"])
        self.assertEqual(rows[1].values, {"FY2024": 2})
        self.assertEqual(rows[2].values, {"FY2024": 3, "FY2025": 30})

    def test_renamed_account_stays_separate(self):
        groups = {"FY2024": [("당기손익인식금융자산", "당기손익인식금융자산", 1)],
                  "FY2025": [("당기손익-공정가치측정금융자산", "당기손익공정가치측정금융자산", 2)]}
        rows = L.join_rows(groups, ["FY2024", "FY2025"])
        self.assertEqual(len(rows), 2)

    def test_duplicate_norm_in_same_period_kept_in_order(self):
        groups = {"FY2025": [("기타", "기타", 1), ("기타", "기타", 2)]}
        rows = L.join_rows(groups, ["FY2025"])
        self.assertEqual([r.values["FY2025"] for r in rows], [1, 2])

    def test_account_ids_follow_latest_order_exactly(self):
        groups = {"FY2024": [("리스부채", "리스부채", 1, "B"), ("리스부채", "리스부채", 2, "A")],
                  "FY2025": [("리스부채", "리스부채", 10, "A"), ("리스부채", "리스부채", 20, "B")]}
        rows = L.join_rows(groups, ["FY2024", "FY2025"])
        self.assertEqual([r.account_id for r in rows], ["A", "B"])

    def test_join_rows_accepts_three_tuples(self):
        rows = L.join_rows({"FY2025": [("x", "x", 1)]}, ["FY2025"])
        self.assertIsNone(rows[0].account_id)

class TestBlocks(unittest.TestCase):
    def test_fs_blocks_bs_and_sce_by_detail(self):
        recs = [fs("FY2025", "자산총계", 100, ord_=1), fs("FY2024", "자산총계", 90, ord_=1),
                fs("FY2025", "기초", 5, "SCE", 1, detail="자본금"), fs("FY2025", "기초", 7, "SCE", 1, detail="총계")]
        blocks = L.fs_blocks(recs, ["FY2024", "FY2025"])
        self.assertEqual([(b.sj_div, b.col_label) for b in blocks], [("BS", None), ("SCE", "자본금"), ("SCE", "총계")])
        self.assertEqual(blocks[0].rows[0].values, {"FY2024": 90, "FY2025": 100})
        self.assertEqual(blocks[0].periods, ["FY2024", "FY2025"])

    def test_note_blocks_split_by_col_label(self):
        recs = [nr("FY2025", "유형자산", "(2) 변동내역 (당기)", "토지", "기초", 1), nr("FY2025", "유형자산", "(2) 변동내역 (당기)", "토지", "취득(주1)", 2),
                nr("FY2024", "유형자산", "(2) 변동내역 (전기)", "토지", "기초", 3), nr("FY2024", "유형자산", "(2) 변동내역 (전기)", "토지", "취득(*)(주1)", 4)]
        blocks = L.note_blocks(recs, ["FY2024", "FY2025"])
        self.assertEqual(len(blocks), 2)
        acq = next(b for b in blocks if b.col_label == "취득(주1)")
        self.assertEqual(acq.rows[0].values, {"FY2024": 4, "FY2025": 2})     # 열 라벨도 norm 으로 묶임
        self.assertEqual(acq.title, "(2) 변동내역 (당기)")

    def test_caption_sig(self):
        self.assertEqual(L.caption_sig("(당기) (단위 : 천원)"), "")
        self.assertEqual(L.caption_sig("(2) 변동내역 (전기) (단위: 원)"), L.caption_sig("(2) 변동내역 (당기)"))

    def test_note_blocks_merge_cur_prev_tables_across_table_idx(self):
        cap = "(2) 당기 및 전기 중 유형자산 장부금액의 변동내역은 다음과 같습니다. ({}) (단위 : 천원)"
        recs = [nr("FY2025", "유형자산", cap.format("당기"), "토지", "기초", 1, idx=1),
                nr("FY2024", "유형자산", cap.format("전기"), "토지", "기초", 3, idx=2)]
        blocks = L.note_blocks(recs, ["FY2024", "FY2025"])
        self.assertEqual(len(blocks), 1)
        self.assertEqual(blocks[0].col_label, "기초")
        self.assertEqual(blocks[0].rows[0].values, {"FY2024": 3, "FY2025": 1})

    def test_note_blocks_identical_captions_same_period_stay_separate(self):
        recs = [nr("FY2025", "유형자산", "(2) 내역 (당기)", "토지", "기초", 1, idx=1),
                nr("FY2025", "유형자산", "(2) 내역 (당기)", "토지", "기초", 2, idx=2),
                nr("FY2024", "유형자산", "(2) 내역 (전기)", "토지", "기초", 7, idx=1),
                nr("FY2024", "유형자산", "(2) 내역 (전기)", "토지", "기초", 8, idx=2)]
        blocks = L.note_blocks(recs, ["FY2024", "FY2025"])
        self.assertEqual(len(blocks), 2)
        self.assertEqual(sorted(tuple(sorted(b.rows[0].values.items())) for b in blocks),
                         [(("FY2024", 7), ("FY2025", 1)), (("FY2024", 8), ("FY2025", 2))])

    def test_unify_units_converts_to_latest(self):
        b = Block("t", 1000, ["FY2024", "FY2025"], [Row("급여", "급여", {"FY2024": 5000, "FY2025": 7})],
                  meta={"units": {"FY2024": 1, "FY2025": 1000}})
        warns = L.unify_units(b)
        self.assertEqual(b.rows[0].values, {"FY2024": 5, "FY2025": 7})
        self.assertTrue(warns)

    def test_unify_units_idempotent_and_records_warning(self):
        b = Block("t", 1000, ["FY2024", "FY2025"], [Row("급여", "급여", {"FY2024": 5000, "FY2025": 7})],
                  meta={"units": {"FY2024": 1, "FY2025": 1000}})
        L.unify_units(b)
        L.unify_units(b)
        self.assertEqual(b.rows[0].values, {"FY2024": 5, "FY2025": 7})
        self.assertEqual(len(b.meta["warnings"]), 1)

    def test_fs_blocks_keeps_unit_warnings(self):
        recs = [fs("FY2024", "자산총계", 90000, unit=1), fs("FY2025", "자산총계", 100, unit=1000)]
        b = L.fs_blocks(recs, ["FY2024", "FY2025"])[0]
        self.assertTrue(b.meta["warnings"])
        self.assertEqual(b.meta["units"], {"FY2024": 1000, "FY2025": 1000})
        self.assertEqual(b.rows[0].values, {"FY2024": 90, "FY2025": 100})

    def test_caption_sig_drops_fiscal_markers(self):
        """상장사 원문: 변동표 소제목이 '① 제 25(당) 기' / '② 제 24(전) 기'. 연도마다 회차가 달라 블록이 해마다 갈리던 문제."""
        self.assertEqual(L.caption_sig("② 제 24(전) 기"), L.caption_sig("② 제 22(전) 기"))
        self.assertEqual(L.caption_sig("① 제 25(당) 기"), L.caption_sig("② 제 24(전) 기"))
        self.assertEqual(L.caption_sig("① 당기 중 사용권자산의 변동내역"), L.caption_sig("② 전기 중 사용권자산의 변동내역"))
        self.assertNotEqual(L.caption_sig("제1종 우선주"), "")

    def test_caption_sig_keeps_subheading_number(self):
        self.assertEqual(L.caption_sig("(1) 장부금액 (당기)"), "항1장부금액")
        self.assertNotEqual(L.caption_sig("(1) 장부금액 (당기)"), L.caption_sig("(2) 장부금액 (당기)"))
        self.assertEqual(L.caption_sig("1) 장부금액 (당기)"), "항1장부금액")

    def test_note_blocks_numbered_subtables_align_when_table_inserted(self):
        recs = [nr("FY2025", "유형자산", "(1) 장부금액 (당기)", "토지", "기초", 1, idx=1),
                nr("FY2025", "유형자산", "(2) 장부금액 (당기)", "토지", "기초", 2, idx=2),
                nr("FY2024", "유형자산", "(2) 장부금액 (전기)", "토지", "기초", 8, idx=1)]
        blocks = L.note_blocks(recs, ["FY2024", "FY2025"])
        b2 = next(b for b in blocks if b.rows[0].values.get("FY2025") == 2)
        self.assertEqual(b2.rows[0].values, {"FY2024": 8, "FY2025": 2})

    def test_note_blocks_warns_when_same_sig_table_count_differs(self):
        recs = [nr("FY2025", "유형자산", "내역 (당기)", "토지", "기초", 1, idx=1),
                nr("FY2025", "유형자산", "내역 (당기)", "토지", "기초", 2, idx=2),
                nr("FY2024", "유형자산", "내역 (전기)", "토지", "기초", 7, idx=1)]
        blocks = L.note_blocks(recs, ["FY2024", "FY2025"])
        msg = "주석 11 유형자산: 같은 캡션 표 수가 기간마다 다름 — 블록 정렬 확인 필요"
        self.assertEqual(len(blocks), 2)
        for b in blocks:
            self.assertIn(msg, b.meta["warnings"])

    def test_note_blocks_no_warning_when_counts_equal(self):
        recs = [nr("FY2025", "유형자산", "내역 (당기)", "토지", "기초", 1),
                nr("FY2024", "유형자산", "내역 (전기)", "토지", "기초", 7)]
        for b in L.note_blocks(recs, ["FY2024", "FY2025"]):
            self.assertFalse(b.meta.get("warnings"))

    def test_note_blocks_latest_by_period_index(self):
        recs = [nr("FY2025", "유형자산", "내역 (당기)", "토지", "기초", 1),
                nr("FY2024", "유형자산", "내역 (전기)", "토지", "기초", 7)]
        b = L.note_blocks(recs, ["FY2025", "FY2024"])[0]     # 기간 목록 마지막이 최신
        self.assertEqual(b.title, "내역 (전기)")

    def test_apply_existing_order_keeps_positions_and_appends_new(self):
        b = Block("t", 1, ["FY2025"], [Row("B", "B", {"FY2025": 2}), Row("C", "C", {"FY2025": 3}), Row("A", "A", {"FY2025": 1})])
        b2 = L.apply_existing_order(b, ["A", "X", "B"])
        self.assertEqual([r.label for r in b2.rows], ["A", "X", "B", "C"])
        self.assertEqual(b2.rows[1].values, {})
        self.assertIn("공시에서 사라짐", b2.rows[1].flags)

    def test_apply_existing_order_sce_opening_closing_rows(self):
        """원문 자본변동표 행은 norm 이 '기초'·'기말'. 같은 라벨(재실행)·다음 해 라벨 모두 기존 행에 붙어야 한다."""
        def blk(open_lab, close_lab):
            return Block("자본변동표 — 총계", 1, ["FY2025"], [Row(open_lab, "기초", {"FY2025": 1}), Row("배당금지급", "배당금지급", {"FY2025": 2}),
                                                     Row(close_lab, "기말", {"FY2025": 3})], sj_div="SCE")
        ex = ["2025.1.1(당기초)", "배당금지급", "2025.12.31(당기말)"]
        for b in (blk("2025.1.1(당기초)", "2025.12.31(당기말)"), blk("2026.1.1(당기초)", "2026.12.31(당기말)")):
            b2 = L.apply_existing_order(b, ex)
            self.assertEqual([r.values.get("FY2025") for r in b2.rows], [1, 2, 3])
            self.assertFalse(any(r.flags for r in b2.rows))

if __name__ == "__main__":
    unittest.main()
