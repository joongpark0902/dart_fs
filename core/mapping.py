"""표준계정·키워드·주석 주제 CSV. 엑셀 매핑표 시트와 왕복한다."""
import csv, os, re
from dataclasses import dataclass, field
from core import normalize as N

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_DIR = os.path.join(HERE, "mapping")
SHEET_HEADER = ["종류", "sj", "키", "표준계정/주제", "우선순위/순서", "시트", "추천"]


_RESID = re.compile(r"^(합계|계|소계|총계|\(.*\))?$")
_PREFIX = re.compile(r"^(총|순)?$")


def _int_str(v, default):
    t = ("" if v is None else str(v)).strip()
    try:
        return str(int(float(t)))
    except ValueError:
        return default


@dataclass
class Topic:
    topic: str
    sheet: str
    recommended: bool
    keywords: list


@dataclass
class Mapping:
    std_by_id: dict = field(default_factory=dict)      # (sj, account_id) -> std
    std_order: dict = field(default_factory=dict)      # (sj, std) -> order
    keywords: list = field(default_factory=list)       # (sj, std, kw_norm, prio)
    topics: list = field(default_factory=list)

    def map_account(self, label, account_id, sj):
        sj = "IS" if sj == "CIS" else sj
        if account_id and (sj, account_id) in self.std_by_id:
            return self.std_by_id[(sj, account_id)], "id"
        key = N.norm_label(label)
        if not key:
            return None, "none"
        best = {}
        for ksj, std, kw, prio in self.keywords:
            if ksj != sj:
                continue
            if key == kw:
                rank = (0, prio)
            elif key.startswith(kw) and _RESID.match(key[len(kw):]):
                rank = (1, prio)
            elif self._contains(key, kw):
                rank = (2, prio)
            else:
                continue
            if std not in best or rank < best[std]:
                best[std] = rank
        if not best:
            return None, "none"
        ranked = sorted(best.items(), key=lambda kv: kv[1])
        if len(ranked) > 1 and ranked[0][1] == ranked[1][1]:
            return None, "none"
        return ranked[0][0], "kw"

    @staticmethod
    def _contains(key, kw):
        i = key.find(kw)
        while i >= 0:
            if _PREFIX.match(key[:i]) and _RESID.match(key[i + len(kw):]):
                return True
            i = key.find(kw, i + 1)
        return False

    def rows_for_sheet(self):
        rows = [list(SHEET_HEADER)]
        for (sj, acc), std in self.std_by_id.items():
            rows.append(["표준코드", sj, acc, std, self.std_order.get((sj, std), ""), "", ""])
        for sj, std, kw, prio in self.keywords:
            rows.append(["키워드", sj, kw, std, prio, "", ""])
        for t in self.topics:
            rows.append(["주석주제", "", ";".join(t.keywords), t.topic, "", t.sheet, "Y" if t.recommended else "N"])
        return rows


def _read(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def load(mapping_dir=None):
    d = mapping_dir or DEFAULT_DIR
    m = Mapping()
    for r in _read(os.path.join(d, "standard_accounts.csv")):
        m.std_by_id[(r["sj"], r["account_id"])] = r["std"]
        m.std_order[(r["sj"], r["std"])] = int(r["order"] or 0)
    for r in _read(os.path.join(d, "keywords.csv")):
        m.keywords.append((r["sj"], r["std"], N.norm_label(r["keyword"]), int(r["prio"] or 9)))
    for r in _read(os.path.join(d, "note_topics.csv")):
        m.topics.append(Topic(r["topic"], r["sheet"], r["recommended"].strip().upper() == "Y",
                              [N.norm_label(k) for k in r["keywords"].split(";") if k.strip()]))
    return m


def save_from_rows(rows, mapping_dir):
    """엑셀 매핑표 시트(rows_for_sheet 형식) → CSV 3개. 헤더 행은 건너뛴다."""
    os.makedirs(mapping_dir, exist_ok=True)
    std, kw, tp = [], [], []
    for r in rows[1:]:
        r = [("" if v is None else str(v).strip()) for v in r] + [""] * 7
        kind = r[0]
        if kind == "표준코드":
            std.append({"sj": r[1], "account_id": r[2], "std": r[3], "order": _int_str(r[4], "0")})
        elif kind == "키워드":
            kw.append({"sj": r[1], "std": r[3], "keyword": r[2], "prio": _int_str(r[4], "9")})
        elif kind == "주석주제":
            tp.append({"topic": r[3], "sheet": r[5], "recommended": r[6] or "N", "keywords": r[2]})
    for name, data, fields in (("standard_accounts.csv", std, ["sj", "account_id", "std", "order"]),
                               ("keywords.csv", kw, ["sj", "std", "keyword", "prio"]),
                               ("note_topics.csv", tp, ["topic", "sheet", "recommended", "keywords"])):
        with open(os.path.join(mapping_dir, name), "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=fields)
            w.writeheader()
            w.writerows(data)
