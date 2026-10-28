import os, unittest
from core import xml_doc as X

FX = os.path.join(os.path.dirname(__file__), "fixtures")
OB = os.path.join(FX, "ob_2025_audit_ofs.xml")

class TestXmlDoc(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = X.load_xml(OB)
        cls.nodes = X.stream(cls.root)

    def test_stream_has_titles_paragraphs_tables_in_order(self):
        kinds = [n.kind for n in self.nodes]
        self.assertIn("title", kinds); self.assertIn("p", kinds); self.assertIn("table", kinds)
        self.assertEqual(sum(1 for k in kinds if k == "table"), 230)
        titles = [X.text(n.elem) for n in self.nodes if n.kind == "title"]
        self.assertIn("주석", titles)
        self.assertTrue(any("재 무 제 표" in t for t in titles))

    def test_nodes_carry_current_title(self):
        notes = [n for n in self.nodes if n.title == "주석"]
        self.assertGreater(len(notes), 100)

    def test_table_grid_expands_spans(self):
        # 유형자산 구성내역: 구 분(rowspan2) | 당기말(colspan4) | 전기말(colspan4)
        for n in self.nodes:
            if n.kind == "table":
                g = X.table_grid(n.elem)
                if g.header and g.header[0][:2] == ["구 분", "당기말"] and g.ncols == 9:
                    self.assertEqual(g.header[0], ["구 분"] + ["당기말"] * 4 + ["전기말"] * 4)
                    self.assertEqual(g.header[1][0], "구 분")
                    self.assertEqual(g.header[1][1:5], ["취득원가", "감가상각누계액", "손상차손누계액", "합 계"])
                    self.assertEqual(g.body[0][0], "토지")
                    self.assertEqual(g.body[0][1], "257,585,832")
                    return
        self.fail("유형자산 구성내역 표를 못 찾음")

    def test_caption_table_detection(self):
        caps = [X.table_grid(n.elem) for n in self.nodes if n.kind == "table"]
        caps = [g for g in caps if X.is_caption_table(g)]
        self.assertTrue(any("(단위" in g.text() for g in caps))
        self.assertTrue(any(g.text().startswith("(당기)") for g in caps))
        body = [X.table_grid(n.elem) for n in self.nodes if n.kind == "table" and X.table_grid(n.elem).header]
        self.assertFalse(any(X.is_caption_table(g) for g in body if g.ncols >= 3 and len(g.body) > 3))

    def test_table_grid_ignores_nested_table_rows(self):
        import xml.etree.ElementTree as ET
        t = ET.fromstring("<TABLE><TBODY>"
            "<TR><TD>a</TD><TD>b</TD></TR>"
            "<TR><TD>c</TD><TD><TABLE><TR><TD>x</TD></TR><TR><TD>y</TD></TR></TABLE></TD></TR>"
            "</TBODY></TABLE>")
        g = X.table_grid(t)
        self.assertEqual(len(g.rows()), 2)
        self.assertEqual(g.body[0], ["a", "b"])
        self.assertEqual(g.body[1], ["c", "xy"])

    def _cap(self, xml):
        import xml.etree.ElementTree as ET
        return X.is_caption_table(X.table_grid(ET.fromstring(xml)))

    def test_caption_negative_and_positive_small_tables(self):
        td = lambda *c: "<TR>" + "".join("<TD>%s</TD>" % x for x in c) + "</TR>"
        self.assertFalse(self._cap("<TABLE>" + td("구분", "당기", "전기") + td("급여", "1,000", "900") + "</TABLE>"))
        self.assertFalse(self._cap("<TABLE>" + td("단위당 원가", "1,000") + "</TABLE>"))
        self.assertTrue(self._cap("<TABLE>" + td("당기말") + "</TABLE>"))
        self.assertTrue(self._cap("<TABLE>" + td("(단위 : 천원)") + "</TABLE>"))
        self.assertTrue(self._cap("<TABLE>" + td("(당기) (단위 : 천원)") + "</TABLE>"))

    def test_section_slicing(self):
        sec = X.section(self.nodes,
                        lambda n: n.kind == "title" and X.compact_title(n) == "주석",
                        lambda n: n.kind == "title")
        self.assertGreater(len(sec), 200)
        self.assertEqual(sec[0].kind, "title")
        self.assertTrue(all(n.kind != "title" for n in sec[1:]))

if __name__ == "__main__":
    unittest.main()
