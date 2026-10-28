"""index_v2.html '파일로 생성' 탭의 순수 함수(layout 수정·연도 범위)를 node 로 돌려 본다. node 가 없으면 건너뜀."""
import json
import os
import re
import shutil
import subprocess
import unittest

HTML = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "web", "index_v2.html")
NODE = shutil.which("node")
FUNCS = ("setLayout", "autoOf", "layoutKind", "layoutHeader", "layoutLabelCols", "layoutPeriod", "layoutUnit", "fPeriodRange",
         "fRemapAfterDelete")


def _extract(src, name):
    m = re.search(r"^function %s\(" % name, src, re.M)
    assert m, name
    i = src.index("{", m.end())
    depth = 0
    for j in range(i, len(src)):
        depth += {"{": 1, "}": -1}.get(src[j], 0)
        if depth == 0:
            return src[m.start():j + 1]
    raise AssertionError(name)


def _run(body):
    with open(HTML, encoding="utf-8") as fh:
        src = fh.read()
    pk = re.search(r"^const PK_RE = .*$", src, re.M).group(0)
    js = "\n".join([pk, "const FI = {files: [], scan: [], layout: {}, auto: {}, notes: {}};"]
                   + [_extract(src, n) for n in FUNCS] + [body])
    out = subprocess.run([NODE, "-e", js], capture_output=True, text=True, encoding="utf-8", timeout=30)
    if out.returncode:
        raise AssertionError(out.stderr)
    return json.loads(out.stdout)


@unittest.skipUnless(NODE, "node 없음")
class TestLayoutOverrides(unittest.TestCase):
    T = "const t = {key: 'a.xlsx|BS|0', kind: 'BS', header_row: 5, auto_periods: {'3': 'FY2025', '4': 'FY2024'}};"

    def test_reselecting_auto_value_removes_override(self):
        r = _run(self.T + """
          autoOf(t);
          const s = [];
          layoutKind(t, 'OTHER'); s.push(JSON.parse(JSON.stringify(FI.layout)));
          layoutKind(t, 'BS'); s.push(JSON.parse(JSON.stringify(FI.layout)));
          layoutHeader(t, 7); layoutHeader(t, 5); s.push(JSON.parse(JSON.stringify(FI.layout)));
          layoutPeriod(t, '4', ''); s.push(JSON.parse(JSON.stringify(FI.layout)));
          layoutPeriod(t, '4', 'FY2024'); s.push(JSON.parse(JSON.stringify(FI.layout)));
          layoutUnit(t, '1000'); layoutUnit(t, '0'); s.push(JSON.parse(JSON.stringify(FI.layout)));
          console.log(JSON.stringify(s));""")
        self.assertEqual(r[0], {"a.xlsx|BS|0": {"kind": "OTHER"}})
        self.assertEqual(r[1:], [{}, {}, {"a.xlsx|BS|0": {"periods": {"4": None}}}, {}, {}])

    def test_auto_value_kept_while_overridden(self):
        # 수정 후 다시 그린 표(t2)는 수정값을 보여 준다 — 자동값은 처음 값으로 남아야 한다
        r = _run(self.T + """
          autoOf(t);
          layoutPeriod(t, '3', 'FY2024'); layoutKind(t, 'IS');
          const t2 = {key: t.key, kind: 'IS', header_row: 5, auto_periods: {'3': 'FY2024', '4': 'FY2024'}};
          autoOf(t2);
          layoutPeriod(t2, '4', 'FY2023'); layoutPeriod(t2, '3', 'FY2025'); layoutKind(t2, 'BS');
          console.log(JSON.stringify(FI.layout));""")
        self.assertEqual(r, {"a.xlsx|BS|0": {"periods": {"4": "FY2023"}}})

    def test_label_cols_override_and_auto_reselect(self):
        r = _run("""
          const t = {key: 'a.xlsx|BS|0', kind: 'BS', header_row: 5, label_cols: [1, 2], auto_periods: {}};
          autoOf(t);
          const s = [];
          layoutLabelCols(t, [2]); s.push(JSON.parse(JSON.stringify(FI.layout)));
          layoutLabelCols(t, [1, 2]); s.push(JSON.parse(JSON.stringify(FI.layout)));
          console.log(JSON.stringify(s));""")
        self.assertEqual(r, [{"a.xlsx|BS|0": {"label_cols": [2]}}, {}])

    def test_rescan_after_header_row_and_label_cols_edit(self):
        with open(HTML, encoding="utf-8") as fh:
            src = fh.read()
        self.assertRegex(src, r"hdrIn\.onchange = \(\) => \{ layoutHeader\(t, \+hdrIn\.value\); fScanAll\(\); \}")
        self.assertIn("layoutLabelCols(t, cols); fScanAll();", src)


@unittest.skipUnless(NODE, "node 없음")
class TestDeleteRemap(unittest.TestCase):
    def test_surviving_duplicate_basenames_keys_move(self):
        r = _run("""
          FI.files = [{path: 'C:/a/x.xlsx', basename: 'x.xlsx'}, {path: 'C:/b/x.xlsx', basename: 'x.xlsx'},
                      {path: 'C:/c/x.xlsx', basename: 'x.xlsx'}, {path: 'C:/d/y.xlsx', basename: 'y.xlsx'}];
          FI.scan = [{path: 'C:/a/x.xlsx', key_base: 'x.xlsx'}, {path: 'C:/b/x.xlsx', key_base: 'x.xlsx#2'},
                     {path: 'C:/c/x.xlsx', key_base: 'x.xlsx#3'}, {path: 'C:/d/y.xlsx', key_base: 'y.xlsx'}];
          FI.layout = {'x.xlsx|BS|0': {kind: 'OTHER'}, 'x.xlsx#2|BS|0': {unit: 1000}, 'x.xlsx#3|IS|1': {periods: {'3': null}}, 'y.xlsx|BS|0': {kind: 'IS'}};
          FI.auto = {'x.xlsx#2|BS|0': {kind: 'BS'}};
          fRemapAfterDelete('C:/a/x.xlsx');
          console.log(JSON.stringify([FI.layout, FI.auto]));""")
        self.assertEqual(r[0], {"x.xlsx|BS|0": {"unit": 1000}, "x.xlsx#2|IS|1": {"periods": {"3": None}}, "y.xlsx|BS|0": {"kind": "IS"}})
        self.assertEqual(r[1], {"x.xlsx|BS|0": {"kind": "BS"}})


@unittest.skipUnless(NODE, "node 없음")
class TestPeriodRange(unittest.TestCase):
    def test_skips_hidden_sheets_and_half_notes(self):
        r = _run("""
          FI.files = [{path: 'A', kind: 'annual'}, {path: 'H', kind: 'half'}];
          const tb = (kind, ap) => ({key: 'k' + Math.random(), kind, auto_periods: ap});
          FI.scan = [
            {path: 'A', scan: {sheets: [
              {name: 'BS', hidden: false, tables: [tb('BS', {'3': 'FY2025', '4': 'FY2024'})]},
              {name: 'old', hidden: true, tables: [tb('BS', {'3': 'FY2019'})]}]}},
            {path: 'H', scan: {sheets: [
              {name: 'BS', hidden: false, tables: [tb('BS', {'3': '2026H1'})]},
              {name: 'note', hidden: false, tables: [tb('NOTE', {'3': 'FY2020'})]}]}}];
          console.log(JSON.stringify(fPeriodRange([{period_end: '2025-12-31', kind: 'annual'}, {period_end: '2026-06-30', kind: 'half'}])));""")
        self.assertEqual(r, {"start_year": 2024, "end_year": 2026, "include_half": True})


if __name__ == "__main__":
    unittest.main()
