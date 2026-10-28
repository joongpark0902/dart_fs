"""가상 회사 '샘플전자'의 엑셀 재무제표 픽스처(테스트 실행 때 tests/_out 에 생성, git 에는 이 코드만).
숫자는 지어낸 값이며 맞물린다: 자산총계 = 부채총계 + 자본총계, 당기순이익 → 이익잉여금, 현금 증감 → 기말 현금."""
import os
from openpyxl import Workbook

M = 1_000_000
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(HERE, "_out", "xlsx_input")

# (과목, 당기, 전기) — 단위 원
BS_ROWS = [
    ("자산", None, None),
    ("Ⅰ. 유동자산", 600 * M, 500 * M),
    ("현금및현금성자산", 200 * M, 150 * M),
    ("매출채권", 250 * M, 220 * M),
    ("재고자산", 150 * M, 130 * M),
    ("Ⅱ. 비유동자산", 900 * M, 840 * M),
    ("유형자산", 800 * M, 750 * M),
    ("무형자산", 100 * M, 90 * M),
    ("자산총계", 1500 * M, 1340 * M),
    ("부채", None, None),
    ("Ⅰ. 유동부채", 400 * M, 400 * M),
    ("매입채무", 300 * M, 280 * M),
    ("단기차입금", 100 * M, 120 * M),
    ("Ⅱ. 비유동부채", 200 * M, 240 * M),
    ("장기차입금", 200 * M, 240 * M),
    ("부채총계", 600 * M, 640 * M),
    ("자본", None, None),
    ("Ⅰ. 자본금", 500 * M, 500 * M),
    ("Ⅱ. 이익잉여금", 400 * M, 200 * M),
    ("자본총계", 900 * M, 700 * M),
    ("부채와자본총계", 1500 * M, 1340 * M),
]
BS_NOTES = {"현금및현금성자산": 4, "매출채권": 5, "재고자산": 6, "유형자산": 12}

IS_ROWS = [
    ("Ⅰ. 매출액", 2000 * M, 1800 * M),
    ("Ⅱ. 매출원가", 1400 * M, 1300 * M),
    ("Ⅲ. 매출총이익", 600 * M, 500 * M),
    ("Ⅳ. 판매비와관리비", 330 * M, 330 * M),
    ("Ⅴ. 영업이익", 270 * M, 170 * M),
    ("Ⅵ. 금융수익", 10 * M, 5 * M),
    ("Ⅶ. 금융비용", 20 * M, 25 * M),
    ("Ⅷ. 법인세비용차감전순이익", 260 * M, 150 * M),
    ("Ⅸ. 법인세비용", 60 * M, 30 * M),
    ("Ⅹ. 당기순이익", 200 * M, 120 * M),
]

CIS_ROWS = [
    ("당기순이익", 200 * M, 120 * M),
    ("기타포괄손익", "-", "-"),
    ("재평가잉여금", "-", "-"),
    ("총포괄이익", 200 * M, 120 * M),
]

# (구분 A열, 세부 B열, 당기, 전기) — 2단 라벨
CF_ROWS = [
    ("Ⅰ. 영업활동현금흐름", None, 300 * M, 250 * M),
    (None, "1. 당기순이익", 200 * M, 120 * M),
    (None, "2. 감가상각비", 80 * M, 70 * M),
    (None, "3. 운전자본의 변동", 20 * M, 60 * M),
    ("Ⅱ. 투자활동현금흐름", None, -150 * M, -200 * M),
    (None, "1. 유형자산의 취득", -130 * M, -180 * M),
    (None, "2. 무형자산의 취득", -20 * M, -20 * M),
    ("Ⅲ. 재무활동현금흐름", None, -100 * M, 20 * M),
    (None, "1. 차입금의 증감", -100 * M, 20 * M),
    ("Ⅳ. 현금의 증가", None, 50 * M, 70 * M),
    ("Ⅴ. 기초의 현금", None, 150 * M, 80 * M),
    ("Ⅵ. 기말의 현금", None, 200 * M, 150 * M),
]

# (구분, 자본금, 이익잉여금, 총계)
SCE_ROWS = [
    ("2024.01.01(전기초)", 500 * M, 80 * M, 580 * M),
    ("당기순이익", "-", 120 * M, 120 * M),
    ("2024.12.31(전기말)", 500 * M, 200 * M, 700 * M),
    ("2025.01.01(당기초)", 500 * M, 200 * M, 700 * M),
    ("당기순이익", "-", 200 * M, 200 * M),
    ("2025.12.31(당기말)", 500 * M, 400 * M, 900 * M),
]

# 2024 파일(전년도 결산)용 자본변동표: 2025 파일의 전기 구간과 맞물린다(2024.12.31 자본총계 700M, 전기초 이익잉여금 80M).
SCE_ROWS_2024 = [
    ("2023.01.01(전기초)", 500 * M, 20 * M, 520 * M),
    ("당기순이익", "-", 60 * M, 60 * M),
    ("2023.12.31(전기말)", 500 * M, 80 * M, 580 * M),
    ("2024.01.01(당기초)", 500 * M, 80 * M, 580 * M),
    ("당기순이익", "-", 120 * M, 120 * M),
    ("2024.12.31(당기말)", 500 * M, 200 * M, 700 * M),
]

# 2024 파일의 당기 재무상태표에서 2025 파일의 전기 열과 다른 값(재작성 테스트용). 유동자산 소계는 그대로(221+129=350).
BS_2024_AS_FILED = {"매출채권": 221 * M, "재고자산": 129 * M}

SGA_ROWS = [
    ("급여", 150 * M, 140 * M),
    ("퇴직급여", 20 * M, 20 * M),
    ("감가상각비", 60 * M, 70 * M),
    ("지급수수료", 100 * M, 100 * M),
    ("합계", 330 * M, 330 * M),
]

# (구분, 토지, 건물, 기계장치, 합계)
PPE_ROWS = [
    ("기초", 300 * M, 300 * M, 150 * M, 750 * M),
    ("취득", "-", 50 * M, 80 * M, 130 * M),
    ("감가상각비", "-", -30 * M, -50 * M, -80 * M),
    ("기말", 300 * M, 320 * M, 180 * M, 800 * M),
]

# 2024 파일용 유형자산 변동표: 기말 750M = 2025 파일의 기초
PPE_ROWS_2024 = [
    ("기초", 300 * M, 280 * M, 120 * M, 700 * M),
    ("취득", "-", 40 * M, 60 * M, 100 * M),
    ("감가상각비", "-", -20 * M, -30 * M, -50 * M),
    ("기말", 300 * M, 300 * M, 150 * M, 750 * M),
]

# (계정코드, 계정과목, 차변, 대변)
TB_ROWS = [
    (101, "현금및현금성자산", 200 * M, None),
    (108, "매출채권", 250 * M, None),
    (146, "재고자산", 150 * M, None),
    (201, "유형자산", 800 * M, None),
    (231, "무형자산", 100 * M, None),
    (251, "매입채무", None, 300 * M),
    (260, "단기차입금", None, 100 * M),
    (293, "장기차입금", None, 200 * M),
    (331, "자본금", None, 500 * M),
    (375, "이익잉여금", None, 200 * M),
    (401, "매출", None, 2000 * M),
    (451, "매출원가", 1400 * M, None),
    (801, "판매관리비", 330 * M, None),
    (901, "금융수익", None, 10 * M),
    (931, "금융비용", 20 * M, None),
    (998, "법인세비용", 60 * M, None),
    (None, "합계", 3310 * M, 3310 * M),
]


def _out(name):
    os.makedirs(OUT_DIR, exist_ok=True)
    return os.path.join(OUT_DIR, name)


def _head(ws, name, sub, unit_col):
    ws["A1"] = name
    ws["A2"] = sub
    ws["A3"] = "샘플전자"
    ws.cell(3, unit_col, "(단위: 원)")


def _two(cur, prev, year):
    """연도별 (당기, 전기) 값. 2025 파일은 (당기, 전기), 2024 파일은 (전기, 전기) — 전전기 값은 만들지 않는다."""
    return (cur, prev) if year >= 2025 else (prev, prev)


def build_statements(path=None, year=2025):
    """① 시트별 BS/IS/CF/SCE. 머리글 4행 '제 12(당) 기'·'제 11(전) 기' + 5행 '2025.12.31'·'2024.12.31'.
    year=2024 면 전년도 결산 파일(제 11(당) 기·제 10(전) 기, 당기 = 2025 파일의 전기 값, 단 BS_2024_AS_FILED 는 다르게)."""
    path = path or _out(f"sample_statements{'' if year == 2025 else f'_{year}'}.xlsx")
    fiscal = 12 - (2025 - year)
    hdr1 = [f"제 {fiscal}(당) 기", f"제 {fiscal - 1}(전) 기"]
    hdr2 = [f"{year}.12.31", f"{year - 1}.12.31"]
    wb = Workbook()
    ws = wb.active
    ws.title = "재무상태표"
    _head(ws, "재무상태표", f"제 {fiscal} 기 {year}년 12월 31일 현재", 4)
    ws.append(["과목", "주석"] + hdr1)
    ws.append([None, None] + hdr2)
    for lab, cur, prev in BS_ROWS:
        c, p = _two(cur, prev, year)
        if year == 2024 and lab in BS_2024_AS_FILED:
            c = BS_2024_AS_FILED[lab]
        ws.append([lab, BS_NOTES.get(lab), c, p])

    ws = wb.create_sheet("손익계산서")
    _head(ws, "손익계산서", f"제 {fiscal} 기 {year}년 1월 1일부터 {year}년 12월 31일까지", 3)
    ws.append(["과목"] + hdr1)
    ws.append([None] + hdr2)
    for lab, cur, prev in IS_ROWS:
        ws.append([lab, *_two(cur, prev, year)])

    ws = wb.create_sheet("현금흐름표")
    _head(ws, "현금흐름표", f"제 {fiscal} 기 {year}년 1월 1일부터 {year}년 12월 31일까지", 4)
    ws.append(["과목", None] + hdr1)
    ws.append([None, None] + hdr2)
    for a, b, cur, prev in CF_ROWS:
        ws.append([a, b, *_two(cur, prev, year)])

    ws = wb.create_sheet("자본변동표")
    _head(ws, "자본변동표", f"제 {fiscal} 기 {year}년 1월 1일부터 {year}년 12월 31일까지", 4)
    ws.append(["구분", "자본금", "이익잉여금", "총계"])
    for row in (SCE_ROWS if year >= 2025 else SCE_ROWS_2024):
        ws.append(list(row))
    wb.save(path)
    return path


def build_is_cis(path=None):
    """② 한 시트에 IS(당기·전기 각각 금액·구성비 2열)와 CIS 두 표."""
    path = path or _out("sample_is_cis.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "손익계산서"
    ws["A1"] = "손익계산서"
    ws["A2"] = "(단위: 원)"
    ws.append(["과목", "당기", None, "전기", None])            # 3행
    ws.append([None, "금액", "구성비", "금액", "구성비"])        # 4행
    for lab, cur, prev in IS_ROWS:                               # 5~14행
        ws.append([lab, cur, round(cur / IS_ROWS[0][1], 4), prev, round(prev / IS_ROWS[0][2], 4)])
    ws.append([])                                                # 15행
    ws.append([])                                                # 16행
    ws.append(["포괄손익계산서"])                                # 17행
    ws.append(["(단위: 원)"])                                    # 18행
    ws.append(["과목", "당기", "전기"])                          # 19행
    for row in CIS_ROWS:                                         # 20~23행
        ws.append(list(row))
    wb.save(path)
    return path


def build_notes(path=None, year=2025):
    """③ 주석 시트: '27. 판매비와관리비'(당기/전기) 표와 '12. 유형자산' 변동표(열=자산 종류). year=2024 면 전년도 파일."""
    path = path or _out(f"sample_notes{'' if year == 2025 else f'_{year}'}.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "주석"
    ws["A1"] = "27. 판매비와관리비"
    ws["A2"] = "(단위: 원)"
    ws.append(["구분", "당기", "전기"])                          # 3행
    for lab, cur, prev in SGA_ROWS:                              # 4~8행
        ws.append([lab, *_two(cur, prev, year)])
    ws.append([])                                                # 9행
    ws.append([])                                                # 10행
    ws.append(["12. 유형자산"])                                   # 11행
    ws.append(["(1) 당기 중 유형자산의 변동내역은 다음과 같습니다."])  # 12행
    ws.append(["(단위: 원)"])                                    # 13행
    ws.append(["구분", "토지", "건물", "기계장치", "합계"])         # 14행
    for row in (PPE_ROWS if year >= 2025 else PPE_ROWS_2024):   # 15~18행
        ws.append(list(row))
    wb.save(path)
    return path


def build_trial_balance(path=None):
    """④ 시산표(계정코드·계정과목·차변·대변), 단위 문구 없음 → OTHER, unit None."""
    path = path or _out("sample_tb.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "시산표"
    ws["A1"] = "합계잔액시산표"
    ws.append([])                                                # 2행
    ws.append(["계정코드", "계정과목", "차변", "대변"])            # 3행
    for row in TB_ROWS:                                          # 4~20행
        ws.append(list(row))
    wb.save(path)
    return path
