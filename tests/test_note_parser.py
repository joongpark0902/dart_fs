import os, unittest
import xml.etree.ElementTree as ET
from core import note_parser as NP, xml_doc as X, normalize as N
from core.model import Period

FX = os.path.join(os.path.dirname(__file__), "fixtures")
PER = {"cur": Period.annual(2025, 28), "prev": Period.annual(2024, 27)}

class TestNoteToc(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ob = X.stream(X.load_xml(os.path.join(FX, "ob_2025_audit_ofs.xml")))
        cls.hg = X.stream(X.load_xml(os.path.join(FX, "hugel_2025_audit_cfs.xml")))

    def test_ob_toc_has_35_top_level_notes(self):
        toc = NP.note_toc(self.ob)
        self.assertEqual([e.no for e in toc], [str(i) for i in range(1, 36)])
        titles = {e.no: e.title for e in toc}
        self.assertEqual(titles["9"], "재고자산")
        self.assertEqual(titles["27"], "판매비, 관리비와 물류비")
        self.assertEqual(titles["11"], "유형자산")
        self.assertEqual(titles["26"], "비용의 성격별 분류")
        self.assertTrue(all(e.end > e.start for e in toc))
        self.assertEqual(sum(e.table_count for e in toc), 103)   # 본문표 103 (캡션표 제외)

    def test_hugel_toc(self):
        toc = NP.note_toc(self.hg)
        self.assertEqual(toc[0].title, "일반적 사항")
        self.assertIn("유형자산", [e.title for e in toc])
        self.assertGreaterEqual(len(toc), 28)

    def test_toc_tolerates_one_skipped_number(self):
        xml = ("<BODY><TITLE>주석</TITLE><P>1. A</P><P>3. C</P><P>7. 본문 속 번호</P><P>4. D</P></BODY>")
        toc = NP.note_toc(X.stream(ET.fromstring(xml)))
        self.assertEqual([e.no for e in toc], ["1", "3", "4"])

class TestOutsidePeriodColumns(unittest.TestCase):
    @staticmethod
    def _recs(header3):
        def tr(cells, tag="TD"):
            return "<TR>" + "".join(f"<{tag}>{c}</{tag}>" for c in cells) + "</TR>"
        xml = ("<BODY><TITLE>주석</TITLE><P>1. A</P><TABLE>"
               + tr(["구 분", "당기", "전기", header3], "TH")
               + tr(["차입금", "1,000", "900", "1,234"]) + tr(["사채", "2,000", "1,800", "5,678"])
               + "</TABLE></BODY>")
        nodes = X.stream(ET.fromstring(xml))
        e = NP.note_toc(nodes)[0]
        nt = NP.note_tables(nodes, e)[0]
        return NP.table_records(nt, e, PER, "OFS", "x")

    def test_unrecognized_numeric_column_skipped(self):
        recs, _ = self._recs("기타")
        self.assertTrue(recs)
        self.assertFalse([r for r in recs if r.col_label == "기타"])

    def test_info_column_kept_as_cur(self):
        recs, _ = self._recs("이자율")
        rate = [r for r in recs if r.col_label == "이자율"]
        self.assertTrue(rate and all(r.period_key == "FY2025" for r in rate))

    def test_missing_cur_period(self):
        nodes = X.stream(ET.fromstring("<BODY><TITLE>주석</TITLE><P>1. A</P><TABLE><TR><TH>구분</TH><TH>당기</TH></TR><TR><TD>a</TD><TD>1</TD></TR></TABLE></BODY>"))
        e = NP.note_toc(nodes)[0]
        nt = NP.note_tables(nodes, e)[0]
        self.assertEqual(NP.table_records(nt, e, {"prev": PER["prev"]}, "OFS", "x"), ([], "other"))

class TestTableRecords(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.nodes = X.stream(X.load_xml(os.path.join(FX, "ob_2025_audit_ofs.xml")))
        cls.toc = {e.no: e for e in NP.note_toc(cls.nodes)}

    def _records(self, no, idx):
        e = self.toc[no]
        nt = NP.note_tables(self.nodes, e)[idx]
        return NP.table_records(nt, e, PER, "OFS", "20260410002383")

    def test_two_col_sga(self):
        recs, kind = self._records("27", 0)
        self.assertEqual(kind, "two_col")
        pay = [r for r in recs if r.row_label == "급여"]
        self.assertEqual({r.period_key: r.value for r in pay}, {"FY2025": 130500219, "FY2024": 130368479})
        self.assertTrue(all(r.col_label is None and r.unit == 1000 for r in recs))
        total = next(r for r in recs if r.row_norm == "합계" and r.period_key == "FY2025")
        self.assertEqual(total.value, 729713929)

    def test_grouped_ppe_composition(self):
        recs, kind = self._records("11", 0)
        self.assertEqual(kind, "grouped")
        land = {(r.period_key, r.col_label): r.value for r in recs if r.row_label == "토지"}
        self.assertEqual(land[("FY2025", "취득원가")], 257585832)
        self.assertEqual(land[("FY2024", "합 계")], 257475143)
        bld = next(r for r in recs if r.row_label == "건물" and r.period_key == "FY2025" and r.col_label == "감가상각누계액")
        self.assertEqual(bld.value, -148045701)

    def test_rollforward_ppe_uses_caption_period(self):
        tables = NP.note_tables(self.nodes, self.toc["11"])
        cur = next(t for t in tables if "(당기)" in t.caption and "기초" in t.grid.header[0])
        prev = next(t for t in tables if "(전기)" in t.caption and "기초" in t.grid.header[0])
        rc, kc = NP.table_records(cur, self.toc["11"], PER, "OFS", "x")
        rp, kp = NP.table_records(prev, self.toc["11"], PER, "OFS", "x")
        self.assertEqual((kc, kp), ("rollforward", "rollforward"))
        self.assertTrue(all(r.period_key == "FY2025" for r in rc))
        self.assertTrue(all(r.period_key == "FY2024" for r in rp))
        acq = next(r for r in rc if r.row_label == "기계장치" and r.col_label == "취득(주1)")
        self.assertEqual(acq.value, 14647100)
        self.assertEqual(cur.caption.count("변동내역"), 1)

    def test_footnote_table_attached_to_previous_not_next_caption(self):
        from core import layout as L
        tables = NP.note_tables(self.nodes, self.toc["11"])
        cur = next(t for t in tables if "(당기)" in t.caption and "기초" in t.grid.header[0])
        prev = next(t for t in tables if "(전기)" in t.caption and "기초" in t.grid.header[0])
        self.assertIn("(주1)", cur.footnote)
        self.assertEqual(L.caption_sig(cur.caption), L.caption_sig(prev.caption))
        e = self.toc["11"]
        recs = NP.table_records(cur, e, PER, "OFS", "x")[0] + NP.table_records(prev, e, PER, "OFS", "x")[0]
        blocks = L.note_blocks(recs, ["FY2024", "FY2025"])
        hit = [b for b in blocks if b.col_label == "기초" and any(
            r.label == "토지" and set(r.values) >= {"FY2024", "FY2025"} for r in b.rows)]
        self.assertTrue(hit)

    def test_text_columns_kept_as_text(self):
        tables = NP.note_tables(self.nodes, self.toc["11"])
        t = next(t for t in tables if "담보" in t.caption)
        recs, kind = NP.table_records(t, self.toc["11"], PER, "OFS", "x")
        self.assertEqual(kind, "grouped")
        who = [r for r in recs if r.col_label == "담보제공처"]
        self.assertTrue(who and all(isinstance(r.value, str) for r in who))

    def test_other_when_no_period(self):
        # 2. 중요한 회계정책의 표들은 기간 표시가 없으므로 전부 other (1번 주석에는 표가 없음)
        e = self.toc["2"]
        kinds = [NP.table_records(t, e, PER, "OFS", "x")[1] for t in NP.note_tables(self.nodes, e)]
        self.assertTrue(kinds)
        self.assertTrue(all(k == "other" for k in kinds))

class TestUnitHeaderRow(unittest.TestCase):
    """휴젤 원문: 헤더 첫 행이 전 열 병합된 '(단위: 천원)'. 열 라벨로 쓰면 판관비 합계 블록이 '— (단위: 천원)' 로 갈린다."""
    @classmethod
    def setUpClass(cls):
        cls.nodes = X.stream(X.load_xml(os.path.join(FX, "hugel_2025_audit_cfs.xml")))
        cls.toc = {e.title.replace(" ", ""): e for e in NP.note_toc(cls.nodes)}
        cls.per = {"cur": Period.annual(2025, 25), "prev": Period.annual(2024, 24)}

    def _recs(self, title, idx):
        e = self.toc[title]
        return NP.table_records(NP.note_tables(self.nodes, e)[idx], e, self.per, "CFS", "x")

    def test_sga_two_col_without_unit_label(self):
        recs, kind = self._recs("판매비와관리비", 0)
        self.assertEqual(kind, "two_col")
        self.assertTrue(recs and all(r.col_label is None and r.unit == 1000 for r in recs))
        self.assertTrue(any(r.row_norm == "합계" and r.period_key == "FY2025" for r in recs))

    def test_rollforward_columns_are_movement_labels(self):
        recs, kind = self._recs("유형자산", 1)
        cols = {r.col_label for r in recs}
        self.assertIn("기 말", cols)
        self.assertFalse(any("단위" in (c or "") for c in cols))


class TestFootnoteJudgement(unittest.TestCase):
    @staticmethod
    def _tables(text):
        xml = ("<BODY><TITLE>주석</TITLE><P>1. A</P>"
               "<TABLE><TR><TH>구분</TH><TH>당기</TH></TR><TR><TD>a</TD><TD>1</TD></TR></TABLE>"
               f"<TABLE><TR><TD>{text}</TD></TR></TABLE>"
               "<TABLE><TR><TH>구분</TH><TH>당기</TH></TR><TR><TD>b</TD><TD>2</TD></TR></TABLE></BODY>")
        nodes = X.stream(ET.fromstring(xml))
        return NP.note_tables(nodes, NP.note_toc(nodes)[0])

    def _assert_footnote(self, text):
        t = self._tables(text)
        self.assertEqual(len(t), 2)
        self.assertIn(text.split()[0], t[0].footnote)
        self.assertNotIn(text.split()[0], t[1].caption)

    def _assert_caption(self, text):
        t = self._tables(text)
        self.assertEqual(len(t), 2)
        self.assertEqual(t[0].footnote, "")
        self.assertIn(text.split()[0], t[1].caption)

    def test_footnotes(self):
        self._assert_footnote("(주1) 당기 중 취득")
        self._assert_footnote("(*) 제주소주 지분")
        self._assert_footnote("주1) 당기 중 처분")

    def test_captions(self):
        self._assert_caption("(주석 12 참조) (단위: 천원)")
        self._assert_caption("(주)OO 관련 담보내역")


def _tr(cells, tag="TD"):
    return "<TR>" + "".join(f"<{tag}>{c}</{tag}>" for c in cells) + "</TR>"


class TestColKind(unittest.TestCase):
    """5-1(a): 대시류 값은 분모에서 뺀다 — 값이 대부분 '-' 인 숫자 열(유형자산 변동표 '토지')이 라벨 열로 붙지 않도록."""
    @staticmethod
    def _kind(vals):
        from types import SimpleNamespace as NS
        return NP._col_kind(NS(body=[[v] for v in vals]), 0)

    def test_dashes_excluded_from_denominator(self):
        self.assertEqual(self._kind(["14,705,207", "-", "-", "(169,881)", "-", "-"]), "num")
        self.assertEqual(self._kind(["원화", "원화", "외화"]), "text")
        self.assertEqual(self._kind(["-", "-"]), "num")
        self.assertEqual(self._kind(["－", "—", "―", " "]), "num")
        self.assertEqual(self._kind(["", ""]), "empty")
        self.assertEqual(self._kind(["원화", "-", "-", "-", "-"]), "text")       # 대시를 빼면 남는 값이 전부 글자

    def test_dash_heavy_numeric_column_is_not_a_label_column(self):
        xml = ("<BODY><TITLE>주석</TITLE><P>1. 유형자산</P><P>1) 당기 (단위: 천원)</P><TABLE>"
               + _tr(["구분", "토지", "건물", "합계"], "TH")
               + _tr(["기초 순장부금액", "14,705,207", "1,000", "14,706,207"])
               + _tr(["취득", "-", "200", "200"]) + _tr(["감가상각비", "-", "(50)", "(50)"])
               + _tr(["처분", "-", "-", "-"]) + _tr(["기타", "-", "-", "-"])
               + _tr(["기말 순장부금액", "14,705,207", "1,150", "14,706,357"]) + "</TABLE></BODY>")
        nodes = X.stream(ET.fromstring(xml))
        e = NP.note_toc(nodes)[0]
        recs, kind = NP.table_records(NP.note_tables(nodes, e)[0], e, PER, "OFS", "x")
        self.assertEqual(kind, "rollforward")
        self.assertEqual({r.row_label for r in recs}, {"기초 순장부금액", "취득", "감가상각비", "처분", "기타", "기말 순장부금액"})
        land = {r.row_label: r.value for r in recs if r.col_label == "토지"}
        self.assertEqual(land["기초 순장부금액"], 14705207)
        self.assertIsNone(land["취득"])


class TestFootnoteParagraph(unittest.TestCase):
    """5-1(b): 앞 표의 각주 문단 '(*) …' 은 다음 표의 캡션에 붙지 않는다(각주표와 같은 취급)."""
    @staticmethod
    def _tables(p_text):
        xml = ("<BODY><TITLE>주석</TITLE><P>1. 유형자산</P><P>(2) 변동내역은 다음과 같습니다.</P>"
               "<TABLE><TR><TD>1) 당기 (단위: 천원)</TD></TR></TABLE>"
               "<TABLE><TR><TH>구분</TH><TH>기초</TH><TH>기말</TH></TR><TR><TD>토지</TD><TD>1</TD><TD>2</TD></TR></TABLE>"
               f"<P>{p_text}</P>"
               "<TABLE><TR><TD>2) 전기 (단위: 천원)</TD></TR></TABLE>"
               "<TABLE><TR><TH>구분</TH><TH>기초</TH><TH>기말</TH></TR><TR><TD>토지</TD><TD>3</TD><TD>4</TD></TR></TABLE></BODY>")
        nodes = X.stream(ET.fromstring(xml))
        return NP.note_tables(nodes, NP.note_toc(nodes)[0])

    def test_footnote_paragraph_dropped_from_next_caption(self):
        for p in ("(*) 손상차손누계액을 합산한 금액입니다.", "(*1) 손상차손누계액 포함", "(주1) 당기 중 취득", "주1) 당기 중 처분"):
            with self.subTest(p=p):
                t = self._tables(p)
                self.assertEqual(len(t), 2)
                self.assertEqual(t[1].caption, "(2) 변동내역은 다음과 같습니다. 2) 전기 (단위: 천원)")
                self.assertIn(p.split()[0], t[0].footnote)
                self.assertNotIn("손상차손", t[1].caption)

    def test_footnote_paragraph_only_before_caption(self):
        """소제목 P 가 없이 각주 P 만 앞에 있으면 캡션은 캡션표 글자만."""
        xml = ("<BODY><TITLE>주석</TITLE><P>1. 유형자산</P>"
               "<TABLE><TR><TH>구분</TH><TH>당기</TH></TR><TR><TD>a</TD><TD>1</TD></TR></TABLE>"
               "<P>(*) 손상차손누계액을 합산한 금액입니다.</P>"
               "<TABLE><TR><TD>2) 전기 (단위: 천원)</TD></TR></TABLE>"
               "<TABLE><TR><TH>구분</TH><TH>기초</TH></TR><TR><TD>b</TD><TD>2</TD></TR></TABLE></BODY>")
        nodes = X.stream(ET.fromstring(xml))
        t = NP.note_tables(nodes, NP.note_toc(nodes)[0])
        self.assertEqual(t[1].caption, "2) 전기 (단위: 천원)")
        self.assertEqual(t[0].footnote, "(*) 손상차손누계액을 합산한 금액입니다.")

    @staticmethod
    def _period_after(p_text):
        """P 뒤에 기간 표시 없는 변동표(기초·기말 열) 하나: 그 표 레코드의 기간 키 집합과 kind."""
        xml = ("<BODY><TITLE>주석</TITLE><P>1. 유형자산</P><P>(1) 당기와 전기 중 변동내역은 다음과 같습니다.</P>"
               "<TABLE><TR><TD>1) 당기 (단위: 천원)</TD></TR></TABLE>"
               "<TABLE><TR><TH>구분</TH><TH>기초</TH><TH>기말</TH></TR><TR><TD>토지</TD><TD>1</TD><TD>2</TD></TR></TABLE>"
               f"<P>{p_text}</P>"
               "<TABLE><TR><TH>구분</TH><TH>기초</TH><TH>기말</TH></TR><TR><TD>토지</TD><TD>3</TD><TD>4</TD></TR></TABLE></BODY>")
        nodes = X.stream(ET.fromstring(xml))
        e = NP.note_toc(nodes)[0]
        t = NP.note_tables(nodes, e)
        recs, kind = NP.table_records(t[1], e, PER, "OFS", "x")
        return t, {r.period_key for r in recs}, kind

    def test_footnote_and_heading_in_one_paragraph(self):
        """원문은 각주와 다음 소제목이 한 P 에 붙기도 한다: '(*) …입니다.2) 전기' → 각주는 앞 표, '2) 전기' 는 다음 표의 소제목."""
        t, pks, kind = self._period_after("(*) 손상차손누계액을 합산한 금액입니다.2) 전기")
        self.assertEqual((kind, pks), ("rollforward", {"FY2024"}))
        self.assertEqual(t[0].footnote, "(*) 손상차손누계액을 합산한 금액입니다.")
        self.assertEqual(t[1].caption, "2) 전기")
        t, pks, kind = self._period_after("(*) 손상차손누계액을 합산한 금액입니다.(2) 당기와 전기 중 연결회사의 무형자산 변동내역은 다음과 같습니다.1) 당기")
        self.assertEqual((kind, pks), ("rollforward", {"FY2025"}))
        self.assertEqual(t[0].footnote, "(*) 손상차손누계액을 합산한 금액입니다.")
        self.assertEqual(t[1].caption, "(2) 당기와 전기 중 연결회사의 무형자산 변동내역은 다음과 같습니다.1) 당기")
        # 각주 두 문장 뒤 소제목, 그리고 각주만 있는 P(소제목 유지 → 앞 소제목 + 다음 캡션표)
        t, pks, kind = self._period_after("(*1) 손상차손누계액을 합산한 금액입니다.(*2) 정부보조금을 차감한 금액입니다.2) 전기")
        self.assertEqual((kind, pks), ("rollforward", {"FY2024"}))
        self.assertEqual(t[0].footnote, "(*1) 손상차손누계액을 합산한 금액입니다. (*2) 정부보조금을 차감한 금액입니다.")
        t, pks, kind = self._period_after("(*) 손상차손누계액을 합산한 금액입니다.")
        self.assertEqual(t[1].caption, "(1) 당기와 전기 중 변동내역은 다음과 같습니다.")   # 소제목 유지, 캡션표는 앞 표가 소비(기간 표시 없음)
        self.assertEqual((kind, pks), ("other", set()))

    def test_paragraph_merely_containing_footnote_token_is_kept(self):
        xml = ("<BODY><TITLE>주석</TITLE><P>1. 유형자산</P>"
               "<TABLE><TR><TH>구분</TH><TH>당기</TH></TR><TR><TD>a</TD><TD>1</TD></TR></TABLE>"
               "<P>(2) 담보(*) 제공 내역은 다음과 같습니다.</P>"
               "<TABLE><TR><TH>구분</TH><TH>당기</TH></TR><TR><TD>b</TD><TD>2</TD></TR></TABLE></BODY>")
        nodes = X.stream(ET.fromstring(xml))
        t = NP.note_tables(nodes, NP.note_toc(nodes)[0])
        self.assertEqual(t[1].caption, "(2) 담보(*) 제공 내역은 다음과 같습니다.")
        self.assertEqual(t[0].footnote, "")


class TestColLabelUnit(unittest.TestCase):
    """5-1(c): 캡션에 단위가 없고 열 라벨이 '자본금(천원)' 처럼 단위를 품으면 그 열의 단위 = 그 단위, 라벨은 단위를 뗀다."""
    @staticmethod
    def _recs(caption):
        xml = ("<BODY><TITLE>주석</TITLE><P>1. 자본금</P>" + f"<P>{caption}</P>" + "<TABLE>"
               + _tr(["구분", "당기", "당기", "전기", "전기"], "TH")
               + _tr(["구분", "발행주식수(주)", "자본금(천원)", "발행주식수(주)", "자본금(천원)"], "TH")
               + _tr(["보통주", "400,000", "200,000", "400,000", "200,000"])
               + _tr(["우선주", "322,172", "161,086", "322,172", "161,086"]) + "</TABLE></BODY>")
        nodes = X.stream(ET.fromstring(xml))
        e = NP.note_toc(nodes)[0]
        return NP.table_records(NP.note_tables(nodes, e)[0], e, PER, "OFS", "x")

    def test_unit_from_col_label_when_caption_has_none(self):
        recs, kind = self._recs("(1) 자본금의 내역은 다음과 같습니다.")
        self.assertEqual(kind, "grouped")
        cap = [r for r in recs if r.col_label == "자본금(천원)"]              # 표시·배치 키는 원문 라벨 그대로, 단위만 천원
        self.assertEqual(len(cap), 4)
        self.assertTrue(all(r.unit == 1000 for r in cap))
        self.assertEqual(N.split_col_unit("자본금(천원)"), ("자본금", 1000))   # 비교용 라벨은 단위를 뗀 것
        self.assertEqual(N.split_col_unit("1주의금액(원)"), ("1주의금액", 1))
        self.assertEqual(N.split_col_unit("발행주식수(주)"), ("발행주식수(주)", None))
        shares = [r for r in recs if r.col_label == "발행주식수(주)"]           # '(주)' 는 단위가 아니다 → 라벨·단위 그대로
        self.assertEqual(len(shares), 4)
        self.assertTrue(all(r.unit == 1 for r in shares))
        self.assertEqual(next(r.value for r in cap if r.row_label == "우선주" and r.period_key == "FY2025"), 161086)

    def test_caption_unit_wins(self):
        recs, _ = self._recs("(1) 자본금의 내역은 다음과 같습니다. (단위: 원)")
        self.assertTrue(recs and all(r.unit == 1 for r in recs))
        self.assertTrue([r for r in recs if r.col_label == "자본금(천원)"])


if __name__ == "__main__":
    unittest.main()
