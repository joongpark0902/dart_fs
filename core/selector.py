"""공시 목록 → 쓸 보고서와 기간별 출처 우선순위(재작성 기준)."""
import calendar, re
from dataclasses import dataclass, field
from datetime import date
from core.model import Period, ReportRef

_NM = re.compile(r"^(?:\[[^\]]*\]\s*)*(사업보고서|반기보고서|분기보고서|연결감사보고서|감사보고서|반기연결검토보고서|반기검토보고서)\s*\((\d{4})\.(\d{2})\)")
_KIND = {"사업보고서": "annual", "감사보고서": "annual", "연결감사보고서": "annual",
         "반기보고서": "half", "반기검토보고서": "half", "반기연결검토보고서": "half"}
_FSDIV = {"사업보고서": ("CFS", "OFS"), "반기보고서": ("CFS", "OFS"), "감사보고서": ("OFS",), "연결감사보고서": ("CFS",),
          "반기검토보고서": ("OFS",), "반기연결검토보고서": ("CFS",)}


@dataclass
class Plan:
    periods: list = field(default_factory=list)
    reports: dict = field(default_factory=dict)
    sources: dict = field(default_factory=dict)
    note_sources: dict = field(default_factory=dict)
    base_report: object = None
    log: list = field(default_factory=list)


def parse_report_nm(nm):
    m = _NM.match(nm or "")
    if not m or m.group(1) == "분기보고서":
        return None
    y, mo = int(m.group(2)), int(m.group(3))
    if not 1 <= mo <= 12:
        return None
    try:
        end = date(y, mo, calendar.monthrange(y, mo)[1])
    except ValueError:              # 연도 0000 등 date 범위 밖
        return None
    return m.group(1), end, "정정" in nm


_TARGET_WORD = re.compile(r"사업보고서|반기보고서|감사보고서|검토보고서")


def _to_date(s):
    return date(int(s[:4]), int(s[4:6]), int(s[6:8]))


def plan_reports(disclosures, start_year, end_year, as_of, include_half, fs_divs):
    plan = Plan()
    fs_divs = set(fs_divs)
    cands = {}
    for d in disclosures:
        parsed = parse_report_nm(d["report_nm"])
        if not parsed:
            nm = d.get("report_nm") or ""
            if _TARGET_WORD.search(nm) and "분기보고서" not in nm:
                plan.log.append(f"{nm} ({d.get('rcept_no')}): 보고서명(기간)을 해석하지 못해 제외")
            continue
        rtype, pend, corrected = parsed
        if pend > as_of or _to_date(d["rcept_dt"]) > as_of or not (start_year - 1 <= pend.year <= end_year + 1):
            continue
        if not fs_divs & set(_FSDIV[rtype]):
            continue
        if _KIND[rtype] == "half" and not include_half:
            continue
        if pend.month != (12 if _KIND[rtype] == "annual" else 6):
            plan.log.append(f"{d['report_nm']}: 비12월 결산(또는 비6월 반기) 보고서 — 1차 범위 밖, 제외")
            continue
        ref = ReportRef(d["rcept_no"], d["report_nm"], _to_date(d["rcept_dt"]), _KIND[rtype], rtype, pend, _FSDIV[rtype], corrected)
        cands.setdefault((rtype, pend), []).append(ref)
    for key, refs in cands.items():
        refs.sort(key=lambda r: (r.rcept_dt, r.rcept_no))
        chosen = refs[-1]
        chosen.superseded = tuple(r.rcept_no for r in refs[:-1])
        if chosen.superseded:
            plan.log.append(f"{chosen.report_nm}: 정정본 {chosen.rcept_no} 채택, 이전 {', '.join(chosen.superseded)}")
        plan.reports[chosen.rcept_no] = chosen

    def reports_ending(y, m):
        return sorted([r for r in plan.reports.values() if r.period_end.year == y and r.period_end.month == m],
                      key=lambda r: (r.rtype in ("사업보고서", "반기보고서"), r.rcept_dt), reverse=True)

    def _sources_for(y, m):
        return ([(r.rcept_no, "prev") for r in reports_ending(y + 1, m)]
                + [(r.rcept_no, "cur") for r in reports_ending(y, m)])

    annual_years = [y for y in range(start_year, end_year + 1) if date(y, 12, 31) <= as_of]
    for y in annual_years:
        src = _sources_for(y, 12)
        if src:
            plan.periods.append(Period.annual(y))
            plan.sources[f"FY{y}"] = src
            plan.note_sources[f"FY{y}"] = list(src)
        else:
            plan.log.append(f"FY{y}: 쓸 보고서 없음")
    if include_half:
        for y in range(start_year, end_year + 1):
            if date(y, 6, 30) > as_of:
                continue
            src = _sources_for(y, 6)
            if src:  # 연간과 동일하게 출처가 하나라도 있으면 열로 만든다(다음 반기보고서 전기 열 우선)
                plan.periods.append(Period.half(y))
                plan.sources[f"{y}H1"] = src
    annual_reports = [r for r in plan.reports.values() if r.kind == "annual"]
    if annual_reports:
        plan.base_report = max(annual_reports, key=lambda r: (r.period_end, r.rtype in ("사업보고서",), r.rcept_dt))
    needed = {rc for lst in plan.sources.values() for rc, _ in lst}
    for rc in list(plan.reports):
        if rc not in needed and (plan.base_report is None or rc != plan.base_report.rcept_no):
            plan.log.append(f"{plan.reports[rc].report_nm}: 사용 안 함(기간 범위 밖)")
            del plan.reports[rc]
    return plan


def fiscal_periods_for(report, fiscal_cur):
    y = report.period_end.year
    fc = fiscal_cur
    if report.kind == "annual":
        return {"cur": Period.annual(y, fc), "prev": Period.annual(y - 1, fc - 1 if fc else None),
                "prev2": Period.annual(y - 2, fc - 2 if fc else None)}
    return {"cur": Period.half(y, fc), "prev": Period.half(y - 1, fc - 1 if fc else None),
            "prev_bs": Period.annual(y - 1, fc - 1 if fc else None)}
