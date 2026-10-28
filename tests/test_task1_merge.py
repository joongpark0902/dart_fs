"""파싱·병합 결함 4건: 각주 표식 헤더·캡션 꼬리 기간·캡션 다른 블록 행 겹침 병합·SCE detail 정규화."""
import copy, dataclasses, unittest
import xml.etree.ElementTree as ET
from datetime import date
from core import note_parser as NP, xml_doc as X, layout as L, normalize as N
from core.model import Period, NoteRecord, FsRecord

PER = {"cur": Period.annual(2025, 28), "prev": Period.annual(2024, 27)}


class TestFootnoteHeaders(unittest.TestCase):
    def test_period_of_strips_footnote_marks(self):
        self.assertIs(NP._period_of("당기(*)", PER), PER["cur"])
        self.assertIs(NP._period_of("전기(주1)", PER), PER["prev"])
        self.assertIs(NP._period_of("전기 ( ** )", PER), PER["prev"])
        self.assertIs(NP._period_of("당기(*1)", PER), PER["cur"])
        self.assertIs(NP._period_of("전기( *2 )", PER), PER["prev"])
        self.assertIs(NP._period_of("2025.12.31(*)", PER), PER["cur"])
        self.assertIs(NP._period_of("2024.12.31(*1)", PER), PER["prev"])

    def test_period_of_existing_behavior(self):
        self.assertIs(NP._period_of("당기", PER), PER["cur"])
        self.assertIs(NP._period_of("제 28 기 (당기)", PER), PER["cur"])
        self.assertIs(NP._period_of("2024년 12월 31일", PER), PER["prev"])
        self.assertIsNone(NP._period_of("급여", PER))

    def test_table_with_footnoted_period_headers_is_two_col(self):
        def tr(cells, tag="TD"):
            return "<TR>" + "".join(f"<{tag}>{c}</{tag}>" for c in cells) + "</TR>"
        xml = ("<BODY><TITLE>주석</TITLE><P>1. A</P><TABLE>" + tr(["구 분", "당기(*)", "전기(*)"], "TH")
               + tr(["급여", "1,000", "900"]) + tr(["복리후생비", "200", "180"]) + "</TABLE></BODY>")
        nodes = X.stream(ET.fromstring(xml))
        e = NP.note_toc(nodes)[0]
        recs, kind = NP.table_records(NP.note_tables(nodes, e)[0], e, PER, "OFS", "x")
        self.assertEqual(kind, "two_col")
        self.assertEqual({(r.row_label, r.period_key): r.value for r in recs},
                         {("급여", "FY2025"): 1000, ("급여", "FY2024"): 900,
                          ("복리후생비", "FY2025"): 200, ("복리후생비", "FY2024"): 180})


class TestCaptionTail(unittest.TestCase):
    def test_caption_tail_cases(self):
        cases = [("1) 당기 (단위: 천원)", "cur"), ("2) 전기 (단위: 천원)", "prev"), ("① 당기", "cur"),
                 ("(1) 당기", "cur"),
                 ("당기와 전기 중 연결회사의 유형자산 변동내역은 다음과 같습니다. (단위: 천원)", None)]
        for cap, role in cases:
            with self.subTest(cap=cap):
                p = NP._caption_period(cap, PER)
                self.assertIs(p, PER[role] if role else None)

    def test_caption_tail_after_sentence_with_subno(self):
        cap = "(1) 당기와 전기 중 유형자산의 변동내역은 다음과 같습니다.1) 당기 (단위: 천원)"
        self.assertIs(NP._caption_period(cap, PER), PER["cur"])

    def test_caption_tail_rejects_two_period_prose(self):
        for cap in ("유형자산의 변동내역 당기와 전기 (단위: 천원)", "당기 및 전기 (단위: 백만원)", "(1) 당기와 전기",
                    "당기와 전기 중 발생한 손익은 전기", "기타포괄손익-공정가치측정 금융자산 중 당기",
                    "당기 또는 전기"):
            with self.subTest(cap=cap):
                self.assertIsNone(NP._caption_period(cap, PER))

    def test_caption_tail_existing_rules_first(self):
        self.assertIs(NP._caption_period("(2) 변동내역 (전기)", PER), PER["prev"])


def nr(pk, cap, row, val, col=None, topic="판관비", idx=0, no="25", title="판매비와관리비"):
    return NoteRecord(pk, "CFS", no, title, topic, idx, cap, row, N.norm_label(row), col, val, 1000, "r")


PERIODS4 = ["FY2022", "FY2023", "FY2024", "FY2025"]
CAP_OLD = "(단위: 천원)"
CAP_NEW = "당기와 전기 중 연결회사의 판매비와관리비의 내역은 다음과 같습니다. (단위: 천원)"
COMMON = [f"항목{i:02d}" for i in range(15)]


def _sets(old_extra, new_extra, old_periods=("FY2022", "FY2023"), new_periods=("FY2024", "FY2025"), topic="판관비",
          common=COMMON):
    recs = []
    for pk in old_periods:
        for i, lab in enumerate(list(common) + old_extra):
            recs.append(nr(pk, CAP_OLD, lab, 100 + i, topic=topic))
    for pk in new_periods:
        for i, lab in enumerate(list(common) + new_extra):
            recs.append(nr(pk, CAP_NEW, lab, 200 + i, topic=topic))
    return recs


def _shape(blocks):
    return [(x.title, x.col_label, [(r.label, dict(r.values)) for r in x.rows]) for x in blocks]


class TestOverlapMerge(unittest.TestCase):
    def test_merge_when_rows_overlap_and_periods_disjoint(self):
        recs = _sets(["옛항목A", "옛항목B"], ["새항목A", "새항목B"])     # 17개 중 15개 동일
        blocks = L.note_blocks(recs, PERIODS4)
        self.assertEqual(len(blocks), 1)
        b = blocks[0]
        self.assertEqual(b.title, CAP_NEW)
        self.assertEqual(sorted({p for r in b.rows for p in r.values}), PERIODS4)
        labels = [r.label for r in b.rows]
        self.assertEqual(len(labels), 19)
        self.assertEqual(set(labels), set(COMMON) | {"옛항목A", "옛항목B", "새항목A", "새항목B"})
        self.assertEqual(labels[-2:], ["새항목A", "새항목B"])
        self.assertEqual(b.rows[0].values, {"FY2022": 100, "FY2023": 100, "FY2024": 200, "FY2025": 200})
        w = [m for m in b.meta.get("warnings", []) if "캡션이 달라도 행이 겹쳐 병합" in m]
        self.assertEqual(len(w), 1)
        self.assertIn("주석 25 판매비와관리비", w[0])
        self.assertIn("FY2022", w[0]); self.assertIn("FY2025", w[0])

    def test_low_overlap_stays_separate(self):
        old = [f"옛{i}" for i in range(7)]
        new = [f"새{i}" for i in range(7)]
        blocks = L.note_blocks(_sets(old, new, common=COMMON[:3]), PERIODS4)   # 3/10 = 0.3
        self.assertEqual(len(blocks), 2)

    def test_overlapping_periods_stay_separate(self):
        recs = _sets(["옛항목A"], ["새항목A"], old_periods=("FY2023", "FY2024"), new_periods=("FY2024", "FY2025"))
        self.assertEqual(len(L.note_blocks(recs, PERIODS4)), 2)

    def test_topic_none_never_merged(self):
        recs = _sets(["옛항목A"], ["새항목A"], topic=None)
        self.assertEqual(len(L.note_blocks(recs, PERIODS4)), 2)

    def test_subheading_rows_not_counted(self):
        # 숫자 없는 소제목 행만 같으면 겹침으로 보지 않는다
        heads = ["소제목1", "소제목2", "소제목3"]
        recs = []
        for pk in ("FY2022", "FY2023"):
            recs += [nr(pk, CAP_OLD, h, None) for h in heads] + [nr(pk, CAP_OLD, "옛값", 1)]
        for pk in ("FY2024", "FY2025"):
            recs += [nr(pk, CAP_NEW, h, None) for h in heads] + [nr(pk, CAP_NEW, "새값", 2)]
        self.assertEqual(len(L.note_blocks(recs, PERIODS4)), 2)

    def test_merge_into_best_overlapping_candidate(self):
        """후보가 둘이면 겹침 비율이 더 높은 블록에 붙인다(첫 후보 아님). A·B 는 서로 겹치지 않아 따로 남는다."""
        a_rows = ["x1", "x2", "x3", "a1", "a2", "a3"]
        b_rows = ["y1", "y2", "y3", "y4", "b1", "b2"]
        c_rows = ["x1", "x2", "x3", "y1", "y2", "y3", "y4"]          # A 와 3/6 = 0.5, B 와 4/6 ≈ 0.67
        recs = ([nr("FY2022", "가 표 (단위: 천원)", r, 1) for r in a_rows]
                + [nr("FY2023", "나 표 (단위: 천원)", r, 2) for r in b_rows]
                + [nr("FY2024", "다 표 (단위: 천원)", r, 3) for r in c_rows])
        blocks = L.note_blocks(recs, PERIODS4)
        self.assertEqual(len(blocks), 2)
        a, bc = blocks
        self.assertEqual({p for r in a.rows for p in r.values}, {"FY2022"})
        self.assertEqual({p for r in bc.rows for p in r.values}, {"FY2023", "FY2024"})
        self.assertIn("y1", [r.label for r in bc.rows])

    def test_idempotent(self):
        recs = _sets(["옛항목A", "옛항목B"], ["새항목A", "새항목B"])
        a = L.note_blocks(recs, PERIODS4)
        self.assertEqual(_shape(a), _shape(L.note_blocks(recs, PERIODS4)))
        self.assertEqual(len(a), 1)
        self.assertEqual(len([m for m in a[0].meta.get("warnings", []) if "병합" in m]), 1)


def fs(pk, label, amt, detail, ord_=1):
    return FsRecord(pk, "CFS", "SCE", ord_, label, N.norm_label(label), amt, 1, detail, None, "r", "XML", "cur")


class TestSceDetailKey(unittest.TestCase):
    def test_detail_spacing_variants_one_block(self):
        recs = [fs("FY2024", "기초", 10, "기타포괄손익 누계액"), fs("FY2025", "기초", 12, "기타포괄손익누계액"),
                fs("FY2024", "기초", 1, "주식발행 초과금"), fs("FY2025", "기초", 2, "주식발행초과금")]
        blocks = L.fs_blocks(recs, ["FY2024", "FY2025"])
        self.assertEqual(len(blocks), 2)
        oci = next(b for b in blocks if "기타포괄" in b.col_label)
        self.assertEqual(oci.col_label, "기타포괄손익누계액")
        self.assertEqual(f"{oci.title} — {oci.col_label}", "자본변동표 — 기타포괄손익누계액")
        self.assertEqual(oci.rows[0].values, {"FY2024": 10, "FY2025": 12})

    def test_data_sheet_finds_section_by_normalized_detail(self):
        from core import pipeline as P, mapping as M
        from excel import compose as C
        from tests.test_pipeline import FakeClient
        cfg = P.BuildConfig("OB", "오비맥주", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1,
                            {"recommended": True}, "default")
        col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        c2 = copy.copy(col)
        c2.fs_records = [dataclasses.replace(r, detail="기타 불입항목")
                         if r.sj_div == "SCE" and r.period_key == "FY2024" and r.detail == "기타불입항목" else r
                         for r in col.fs_records]
        wb = C.compose(c2, None, "p", "s")
        titles = [r[C.RAW_LABEL_COL - 1] for r in wb.get("rFS_별도").values_matrix()]
        self.assertEqual([t for t in titles if str(t).startswith("자본변동표 — 기타")], ["자본변동표 — 기타불입항목"])
        vals = wb.get("데이터").values_matrix()
        blk = {r[1] for r in vals[1:] if r[0] == "rFS_별도" and r[2] in ("기타 불입항목", "기타불입항목")}
        self.assertEqual(len(blk), 1)
        self.assertIn("기타불입항목", next(iter(blk)))


class TestMergedNoteDataSheet(unittest.TestCase):
    def test_merged_note_block_keeps_section_for_old_caption_rows(self):
        from core import pipeline as P, mapping as M
        from excel import compose as C
        from tests.test_pipeline import FakeClient
        cfg = P.BuildConfig("OB", "오비맥주", date(2026, 10, 26), 2024, 2025, False, ["OFS"], 1,
                            {"recommended": True}, "default")
        col = P.collect(cfg, FakeClient(), M.load(), lambda s: None)
        c2 = copy.copy(col)
        cap = "당기와 전기 중 판매비와관리비의 내역은 다음과 같습니다. (단위: 천원)"
        c2.note_records = [dataclasses.replace(r, caption=cap) if r.topic == "판관비" and r.period_key == "FY2025" else r
                           for r in col.note_records]
        wb = C.compose(c2, None, "p", "s")
        titles = [r[C.RAW_LABEL_COL - 1] for r in wb.get("rN_판관비").values_matrix()]
        secs = [t for t in titles if str(t).startswith("주석 ")]
        self.assertEqual(len(secs), 1)
        self.assertIn(cap, secs[0])
        vals = wb.get("데이터").values_matrix()
        blk = {(r[4], r[1]) for r in vals[1:] if r[0] == "rN_판관비"}
        self.assertEqual({b for _, b in blk}, {secs[0]})
        self.assertEqual({p for p, _ in blk}, {"FY2024", "FY2025"})


if __name__ == "__main__":
    unittest.main()
