# -*- coding: utf-8 -*-
"""프로젝트 트랙커 엑셀 파일 생성 스크립트.

시트 구성:
  1. 대시보드   - 프로젝트 현황 요약 (자동 집계)
  2. 프로젝트   - 프로젝트 목록 (드롭다운, 조건부 서식, 자동 계산)
  3. 작업       - 프로젝트별 세부 작업 목록
  4. 간트차트   - 주 단위 타임라인 (조건부 서식 기반 자동 차트)
  5. 설정       - 드롭다운 목록 및 팀원 관리

사용법: python3 build_tracker.py
"""
import datetime as dt

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "맑은 고딕"

# ---------- 색상 팔레트 ----------
C_HEADER_BG = "1F3864"   # 진한 남색
C_HEADER_FG = "FFFFFF"
C_TITLE = "1F3864"
C_ACCENT = "2E75B6"
C_LIGHT = "D6E4F0"
C_DONE = "C6EFCE"        # 완료 - 초록
C_DONE_FG = "006100"
C_PROGRESS = "FFEB9C"    # 진행중 - 노랑
C_PROGRESS_FG = "9C6500"
C_DELAY = "FFC7CE"       # 지연 - 빨강
C_DELAY_FG = "9C0006"
C_HOLD = "D9D9D9"        # 보류 - 회색
C_HOLD_FG = "404040"
C_PLAN = "DDEBF7"        # 계획 - 하늘
C_PLAN_FG = "1F4E79"

THIN = Side(style="thin", color="B4C6E7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

STATUSES = ["계획", "진행중", "완료", "보류", "지연"]
PRIORITIES = ["높음", "중간", "낮음"]
CATEGORIES = ["개발", "디자인", "마케팅", "운영", "기획", "기타"]
MEMBERS = ["김철수", "이영희", "박민수", "정수진", "최지훈"]

TODAY = dt.date(2026, 7, 10)  # 샘플 데이터 기준일 (수식은 TODAY() 사용)


def header_style(cell):
    cell.font = Font(name=FONT, bold=True, color=C_HEADER_FG, size=10)
    cell.fill = PatternFill("solid", fgColor=C_HEADER_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER


def body_style(cell, horizontal="center"):
    cell.font = Font(name=FONT, size=10)
    cell.alignment = Alignment(horizontal=horizontal, vertical="center")
    cell.border = BORDER


wb = Workbook()

# =====================================================================
# 1. 설정 시트 (먼저 만들어야 드롭다운 참조 가능)
# =====================================================================
ws_set = wb.active
ws_set.title = "설정"
ws_set.sheet_view.showGridLines = False

ws_set["B2"] = "⚙️ 설정 - 드롭다운 목록 관리"
ws_set["B2"].font = Font(name=FONT, bold=True, size=14, color=C_TITLE)
ws_set["B3"] = "아래 목록을 수정하면 각 시트의 드롭다운 선택지가 바뀝니다. (항목 추가 시 이름 관리자에서 범위 확장 필요)"
ws_set["B3"].font = Font(name=FONT, size=9, color="808080")

set_cols = [("B", "상태", STATUSES), ("C", "우선순위", PRIORITIES),
            ("D", "분류", CATEGORIES), ("E", "팀원", MEMBERS)]
for col, title, items in set_cols:
    c = ws_set[f"{col}5"]
    c.value = title
    header_style(c)
    for i, item in enumerate(items):
        cell = ws_set[f"{col}{6 + i}"]
        cell.value = item
        body_style(cell)
    ws_set.column_dimensions[col].width = 14

# 상태별 색상 안내
ws_set["G5"] = "상태별 색상"
header_style(ws_set["G5"])
ws_set.column_dimensions["G"].width = 14
status_colors = [("계획", C_PLAN, C_PLAN_FG), ("진행중", C_PROGRESS, C_PROGRESS_FG),
                 ("완료", C_DONE, C_DONE_FG), ("보류", C_HOLD, C_HOLD_FG),
                 ("지연", C_DELAY, C_DELAY_FG)]
for i, (name, bg, fg) in enumerate(status_colors):
    cell = ws_set[f"G{6 + i}"]
    cell.value = name
    cell.font = Font(name=FONT, size=10, color=fg, bold=True)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = BORDER

# 이름 정의 (드롭다운 참조용)
from openpyxl.workbook.defined_name import DefinedName
wb.defined_names.add(DefinedName("상태목록", attr_text=f"설정!$B$6:$B${5 + len(STATUSES)}"))
wb.defined_names.add(DefinedName("우선순위목록", attr_text=f"설정!$C$6:$C${5 + len(PRIORITIES)}"))
wb.defined_names.add(DefinedName("분류목록", attr_text=f"설정!$D$6:$D${5 + len(CATEGORIES)}"))
wb.defined_names.add(DefinedName("팀원목록", attr_text=f"설정!$E$6:$E${5 + len(MEMBERS)}"))
wb.defined_names.add(DefinedName("프로젝트ID목록", attr_text="프로젝트!$B$6:$B$105"))

# =====================================================================
# 2. 프로젝트 시트
# =====================================================================
ws_p = wb.create_sheet("프로젝트")
ws_p.sheet_view.showGridLines = False

PROJ_HEADERS = ["프로젝트 ID", "프로젝트명", "담당자", "분류", "우선순위", "상태",
                "시작일", "마감일", "진행률", "남은 일수", "기간(일)", "비고"]
PROJ_WIDTHS = [12, 28, 10, 10, 10, 10, 12, 12, 10, 10, 9, 30]

ws_p["B2"] = "📋 프로젝트 목록"
ws_p["B2"].font = Font(name=FONT, bold=True, size=14, color=C_TITLE)
ws_p["B3"] = "행을 추가하면 수식(남은 일수·기간)을 아래로 복사하세요. 상태/우선순위 등은 드롭다운으로 선택합니다."
ws_p["B3"].font = Font(name=FONT, size=9, color="808080")

HR = 5           # 헤더 행
FIRST = HR + 1   # 데이터 시작 행
MAX_ROWS = 100   # 수식/서식 적용 행 수

for i, (h, w) in enumerate(zip(PROJ_HEADERS, PROJ_WIDTHS)):
    col = get_column_letter(2 + i)  # B부터
    cell = ws_p[f"{col}{HR}"]
    cell.value = h
    header_style(cell)
    ws_p.column_dimensions[col].width = w
ws_p.column_dimensions["A"].width = 2
ws_p.row_dimensions[HR].height = 28

# 샘플 데이터
sample_projects = [
    ("PRJ-001", "신규 홈페이지 리뉴얼", "김철수", "개발", "높음", "진행중",
     dt.date(2026, 6, 1), dt.date(2026, 8, 14), 0.55, "메인 페이지 시안 확정"),
    ("PRJ-002", "모바일 앱 v2.0 출시", "이영희", "개발", "높음", "진행중",
     dt.date(2026, 5, 18), dt.date(2026, 7, 31), 0.72, "QA 진행 중"),
    ("PRJ-003", "여름 프로모션 캠페인", "박민수", "마케팅", "중간", "완료",
     dt.date(2026, 6, 1), dt.date(2026, 6, 30), 1.0, ""),
    ("PRJ-004", "사내 ERP 도입 검토", "정수진", "운영", "낮음", "보류",
     dt.date(2026, 7, 1), dt.date(2026, 9, 30), 0.1, "예산 승인 대기"),
    ("PRJ-005", "브랜드 리디자인", "최지훈", "디자인", "중간", "지연",
     dt.date(2026, 5, 1), dt.date(2026, 7, 5), 0.8, "외주 일정 지연"),
    ("PRJ-006", "고객 만족도 조사", "이영희", "기획", "중간", "계획",
     dt.date(2026, 7, 21), dt.date(2026, 8, 8), 0.0, "설문 문항 설계 예정"),
]

for r_off, (pid, name, owner, cat, pri, status, start, end, prog, note) in enumerate(sample_projects):
    r = FIRST + r_off
    values = [pid, name, owner, cat, pri, status, start, end, prog, None, None, note]
    for c_off, v in enumerate(values):
        col = get_column_letter(2 + c_off)
        cell = ws_p[f"{col}{r}"]
        if v is not None:
            cell.value = v

# 수식·서식 (샘플 + 빈 행 포함 MAX_ROWS까지)
for r in range(FIRST, FIRST + MAX_ROWS):
    for c_off in range(len(PROJ_HEADERS)):
        col = get_column_letter(2 + c_off)
        cell = ws_p[f"{col}{r}"]
        body_style(cell, horizontal="left" if col in ("C", "M") else "center")
    ws_p[f"H{r}"].number_format = "yyyy-mm-dd"
    ws_p[f"I{r}"].number_format = "yyyy-mm-dd"
    ws_p[f"J{r}"].number_format = "0%"
    # 남은 일수: 완료면 "-", 마감 지났으면 음수
    ws_p[f"K{r}"] = (f'=IF($B{r}="","",IF($G{r}="완료","-",I{r}-TODAY()))')
    # 기간(일)
    ws_p[f"L{r}"] = f'=IF($B{r}="","",I{r}-H{r}+1)'

# 드롭다운 (데이터 유효성 검사)
rng = f"{FIRST}:{FIRST + MAX_ROWS - 1}"
dv_status = DataValidation(type="list", formula1="=상태목록", allow_blank=True)
dv_pri = DataValidation(type="list", formula1="=우선순위목록", allow_blank=True)
dv_cat = DataValidation(type="list", formula1="=분류목록", allow_blank=True)
dv_member = DataValidation(type="list", formula1="=팀원목록", allow_blank=True)
ws_p.add_data_validation(dv_status); dv_status.add(f"G{FIRST}:G{FIRST + MAX_ROWS - 1}")
ws_p.add_data_validation(dv_pri); dv_pri.add(f"F{FIRST}:F{FIRST + MAX_ROWS - 1}")
ws_p.add_data_validation(dv_cat); dv_cat.add(f"E{FIRST}:E{FIRST + MAX_ROWS - 1}")
ws_p.add_data_validation(dv_member); dv_member.add(f"D{FIRST}:D{FIRST + MAX_ROWS - 1}")

# 조건부 서식: 상태 색상
status_rng = f"G{FIRST}:G{FIRST + MAX_ROWS - 1}"
for name, bg, fg in status_colors:
    ws_p.conditional_formatting.add(
        status_rng,
        CellIsRule(operator="equal", formula=[f'"{name}"'],
                   fill=PatternFill("solid", fgColor=bg),
                   font=Font(name=FONT, color=fg, bold=True)))

# 우선순위 '높음' 강조
ws_p.conditional_formatting.add(
    f"F{FIRST}:F{FIRST + MAX_ROWS - 1}",
    CellIsRule(operator="equal", formula=['"높음"'],
               font=Font(name=FONT, color="C00000", bold=True)))

# 진행률 데이터 막대
ws_p.conditional_formatting.add(
    f"J{FIRST}:J{FIRST + MAX_ROWS - 1}",
    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                color=C_ACCENT, showValue=True))

# 마감 초과(미완료) 행: 남은 일수 빨간색
ws_p.conditional_formatting.add(
    f"K{FIRST}:K{FIRST + MAX_ROWS - 1}",
    FormulaRule(formula=[f'AND($B{FIRST}<>"",$G{FIRST}<>"완료",ISNUMBER($K{FIRST}),$K{FIRST}<0)'],
                fill=PatternFill("solid", fgColor=C_DELAY),
                font=Font(name=FONT, color=C_DELAY_FG, bold=True)))
# 마감 임박(7일 이내) 주황색
ws_p.conditional_formatting.add(
    f"K{FIRST}:K{FIRST + MAX_ROWS - 1}",
    FormulaRule(formula=[f'AND($B{FIRST}<>"",$G{FIRST}<>"완료",ISNUMBER($K{FIRST}),$K{FIRST}>=0,$K{FIRST}<=7)'],
                fill=PatternFill("solid", fgColor="FFE4C4"),
                font=Font(name=FONT, color="C55A11", bold=True)))

ws_p.freeze_panes = f"D{FIRST}"
ws_p.auto_filter.ref = f"B{HR}:M{FIRST + MAX_ROWS - 1}"

# =====================================================================
# 3. 작업 시트
# =====================================================================
ws_t = wb.create_sheet("작업")
ws_t.sheet_view.showGridLines = False

TASK_HEADERS = ["작업 ID", "프로젝트 ID", "작업명", "담당자", "우선순위", "상태",
                "시작일", "마감일", "진행률", "남은 일수", "비고"]
TASK_WIDTHS = [10, 12, 32, 10, 10, 10, 12, 12, 10, 10, 28]

ws_t["B2"] = "✅ 작업(태스크) 목록"
ws_t["B2"].font = Font(name=FONT, bold=True, size=14, color=C_TITLE)
ws_t["B3"] = "프로젝트 ID를 입력해 프로젝트와 연결하세요. 프로젝트 시트의 진행률은 여기 작업 진행률을 참고해 직접 갱신합니다."
ws_t["B3"].font = Font(name=FONT, size=9, color="808080")

for i, (h, w) in enumerate(zip(TASK_HEADERS, TASK_WIDTHS)):
    col = get_column_letter(2 + i)
    cell = ws_t[f"{col}{HR}"]
    cell.value = h
    header_style(cell)
    ws_t.column_dimensions[col].width = w
ws_t.column_dimensions["A"].width = 2
ws_t.row_dimensions[HR].height = 28

sample_tasks = [
    ("T-001", "PRJ-001", "와이어프레임 작성", "김철수", "높음", "완료",
     dt.date(2026, 6, 1), dt.date(2026, 6, 12), 1.0, ""),
    ("T-002", "PRJ-001", "메인 페이지 퍼블리싱", "김철수", "높음", "진행중",
     dt.date(2026, 6, 15), dt.date(2026, 7, 18), 0.6, ""),
    ("T-003", "PRJ-001", "백엔드 API 연동", "박민수", "중간", "진행중",
     dt.date(2026, 7, 1), dt.date(2026, 8, 7), 0.3, ""),
    ("T-004", "PRJ-002", "베타 테스트 피드백 반영", "이영희", "높음", "진행중",
     dt.date(2026, 6, 22), dt.date(2026, 7, 17), 0.8, ""),
    ("T-005", "PRJ-002", "스토어 심사 제출", "이영희", "높음", "계획",
     dt.date(2026, 7, 20), dt.date(2026, 7, 27), 0.0, "심사 기간 고려"),
    ("T-006", "PRJ-005", "로고 시안 확정", "최지훈", "중간", "지연",
     dt.date(2026, 6, 1), dt.date(2026, 7, 3), 0.9, "외주사 회신 대기"),
]

for r_off, (tid, pid, name, owner, pri, status, start, end, prog, note) in enumerate(sample_tasks):
    r = FIRST + r_off
    values = [tid, pid, name, owner, pri, status, start, end, prog, None, note]
    for c_off, v in enumerate(values):
        col = get_column_letter(2 + c_off)
        if v is not None:
            ws_t[f"{col}{r}"] = v

for r in range(FIRST, FIRST + MAX_ROWS):
    for c_off in range(len(TASK_HEADERS)):
        col = get_column_letter(2 + c_off)
        cell = ws_t[f"{col}{r}"]
        body_style(cell, horizontal="left" if col in ("D", "L") else "center")
    ws_t[f"H{r}"].number_format = "yyyy-mm-dd"
    ws_t[f"I{r}"].number_format = "yyyy-mm-dd"
    ws_t[f"J{r}"].number_format = "0%"
    ws_t[f"K{r}"] = f'=IF($B{r}="","",IF($G{r}="완료","-",I{r}-TODAY()))'
    # 헬퍼 키: "프로젝트ID|해당 프로젝트에서의 몇 번째 작업인지"
    # (프로젝트 상세 시트가 MATCH로 이 키를 찾아 작업을 나열한다)
    ws_t[f"N{r}"] = f'=IF($B{r}="","",$C{r}&"|"&COUNTIF($C$6:$C{r},$C{r}))'
ws_t.column_dimensions["N"].hidden = True

dv_status2 = DataValidation(type="list", formula1="=상태목록", allow_blank=True)
dv_pri2 = DataValidation(type="list", formula1="=우선순위목록", allow_blank=True)
dv_member2 = DataValidation(type="list", formula1="=팀원목록", allow_blank=True)
ws_t.add_data_validation(dv_status2); dv_status2.add(f"G{FIRST}:G{FIRST + MAX_ROWS - 1}")
ws_t.add_data_validation(dv_pri2); dv_pri2.add(f"F{FIRST}:F{FIRST + MAX_ROWS - 1}")
ws_t.add_data_validation(dv_member2); dv_member2.add(f"E{FIRST}:E{FIRST + MAX_ROWS - 1}")

status_rng_t = f"G{FIRST}:G{FIRST + MAX_ROWS - 1}"
for name, bg, fg in status_colors:
    ws_t.conditional_formatting.add(
        status_rng_t,
        CellIsRule(operator="equal", formula=[f'"{name}"'],
                   fill=PatternFill("solid", fgColor=bg),
                   font=Font(name=FONT, color=fg, bold=True)))
ws_t.conditional_formatting.add(
    f"J{FIRST}:J{FIRST + MAX_ROWS - 1}",
    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                color=C_ACCENT, showValue=True))
ws_t.conditional_formatting.add(
    f"K{FIRST}:K{FIRST + MAX_ROWS - 1}",
    FormulaRule(formula=[f'AND($B{FIRST}<>"",$G{FIRST}<>"완료",ISNUMBER($K{FIRST}),$K{FIRST}<0)'],
                fill=PatternFill("solid", fgColor=C_DELAY),
                font=Font(name=FONT, color=C_DELAY_FG, bold=True)))

ws_t.freeze_panes = f"E{FIRST}"
ws_t.auto_filter.ref = f"B{HR}:L{FIRST + MAX_ROWS - 1}"

# =====================================================================
# 3.5 프로젝트 상세 시트 (드롭다운으로 프로젝트 선택 → 작업 자동 필터)
# =====================================================================
ws_v = wb.create_sheet("프로젝트 상세")
ws_v.sheet_view.showGridLines = False

ws_v["B2"] = "🔍 프로젝트 상세"
ws_v["B2"].font = Font(name=FONT, bold=True, size=14, color=C_TITLE)
ws_v["B3"] = "C5 셀에서 프로젝트를 선택하면 해당 프로젝트의 정보와 작업 목록이 자동으로 표시됩니다. 이 시트에는 데이터를 직접 입력하지 마세요."
ws_v["B3"].font = Font(name=FONT, size=9, color="808080")

DETAIL_WIDTHS = {"B": 12, "C": 30, "D": 10, "E": 10, "F": 10,
                 "G": 12, "H": 12, "I": 10, "J": 10, "K": 26}
for col, w in DETAIL_WIDTHS.items():
    ws_v.column_dimensions[col].width = w
ws_v.column_dimensions["A"].width = 2
ws_v.column_dimensions["N"].hidden = True  # 헬퍼 열 숨김

# 프로젝트 선택 드롭다운
ws_v["B5"] = "프로젝트 선택"
header_style(ws_v["B5"])
ws_v["C5"] = "PRJ-001"
ws_v["C5"].font = Font(name=FONT, bold=True, size=12, color=C_ACCENT)
ws_v["C5"].fill = PatternFill("solid", fgColor=C_LIGHT)
ws_v["C5"].alignment = Alignment(horizontal="center", vertical="center")
ws_v["C5"].border = BORDER
ws_v.row_dimensions[5].height = 24
dv_proj = DataValidation(type="list", formula1="=프로젝트ID목록", allow_blank=True)
ws_v.add_data_validation(dv_proj)
dv_proj.add("C5")

# 프로젝트 요약 카드
SEL = "$C$5"
M_ROW = f"MATCH({SEL},프로젝트!$B$6:$B$105,0)"
info_headers = [("B", "프로젝트명"), ("D", "담당자"), ("E", "상태"), ("F", "시작일"),
                ("G", "마감일"), ("H", "진행률"), ("I", "작업 수"), ("J", "완료 작업")]
INFO_HR, INFO_VR = 7, 8
ws_v.merge_cells(f"B{INFO_HR}:C{INFO_HR}")
ws_v.merge_cells(f"B{INFO_VR}:C{INFO_VR}")
for col, title in info_headers:
    cell = ws_v[f"{col}{INFO_HR}"]
    cell.value = title
    header_style(cell)
ws_v[f"C{INFO_HR}"].border = BORDER

def proj_lookup(col_letter):
    return f'=IF({SEL}="","",IFERROR(INDEX(프로젝트!${col_letter}$6:${col_letter}$105,{M_ROW}),""))'

info_values = [
    ("B", proj_lookup("C"), None),
    ("D", proj_lookup("D"), None),
    ("E", proj_lookup("G"), None),
    ("F", proj_lookup("H"), "yyyy-mm-dd"),
    ("G", proj_lookup("I"), "yyyy-mm-dd"),
    ("H", proj_lookup("J"), "0%"),
    ("I", f'=IF({SEL}="","",COUNTIFS(작업!$C$6:$C$105,{SEL},작업!$B$6:$B$105,"<>"))', None),
    ("J", f'=IF({SEL}="","",COUNTIFS(작업!$C$6:$C$105,{SEL},작업!$G$6:$G$105,"완료"))', None),
]
for col, formula, fmt in info_values:
    cell = ws_v[f"{col}{INFO_VR}"]
    cell.value = formula
    body_style(cell)
    cell.font = Font(name=FONT, size=10, bold=True)
    if fmt:
        cell.number_format = fmt
ws_v[f"C{INFO_VR}"].border = BORDER
ws_v.row_dimensions[INFO_VR].height = 22
for name, bg, fg in status_colors:
    ws_v.conditional_formatting.add(
        f"E{INFO_VR}",
        CellIsRule(operator="equal", formula=[f'"{name}"'],
                   fill=PatternFill("solid", fgColor=bg),
                   font=Font(name=FONT, color=fg, bold=True)))

# 작업 목록 (자동 필터링)
DETAIL_HEADERS = ["작업 ID", "작업명", "담당자", "우선순위", "상태",
                  "시작일", "마감일", "진행률", "남은 일수", "비고"]
D_HR = 10          # 작업 테이블 헤더 행
D_FIRST = D_HR + 1
D_ROWS = 60
for i, h in enumerate(DETAIL_HEADERS):
    col = get_column_letter(2 + i)
    cell = ws_v[f"{col}{D_HR}"]
    cell.value = h
    header_style(cell)
ws_v.row_dimensions[D_HR].height = 24

# 헬퍼 열(N): 선택된 프로젝트의 n번째 작업이 있는 상대 행 번호.
# 작업 시트의 헬퍼 키(N열, "프로젝트ID|n")를 MATCH로 찾는다 —
# 배열 수식이 아니므로 Excel의 암시적 교차(@) 변환에 영향받지 않는다.
# 작업 시트 열 매핑: B=작업ID, D=작업명, E=담당자, F=우선순위, G=상태,
#                    H=시작일, I=마감일, J=진행률, K=남은일수, L=비고
SRC_COLS = ["B", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
for i in range(D_ROWS):
    r = D_FIRST + i
    ws_v[f"N{r}"] = (
        f'=IF({SEL}="","",IFERROR(MATCH({SEL}&"|"&ROWS($N${D_FIRST}:$N{r}),'
        f'작업!$N$6:$N$105,0),""))')
    for c_off, src in enumerate(SRC_COLS):
        col = get_column_letter(2 + c_off)
        cell = ws_v[f"{col}{r}"]
        cell.value = f'=IF($N{r}="","",INDEX(작업!{src}$6:{src}$105,$N{r}))'
        body_style(cell, horizontal="left" if col in ("C", "K") else "center")
    ws_v[f"G{r}"].number_format = "yyyy-mm-dd"
    ws_v[f"H{r}"].number_format = "yyyy-mm-dd"
    ws_v[f"I{r}"].number_format = "0%"

d_status_rng = f"F{D_FIRST}:F{D_FIRST + D_ROWS - 1}"
for name, bg, fg in status_colors:
    ws_v.conditional_formatting.add(
        d_status_rng,
        CellIsRule(operator="equal", formula=[f'"{name}"'],
                   fill=PatternFill("solid", fgColor=bg),
                   font=Font(name=FONT, color=fg, bold=True)))
ws_v.conditional_formatting.add(
    f"I{D_FIRST}:I{D_FIRST + D_ROWS - 1}",
    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                color=C_ACCENT, showValue=True))
ws_v.conditional_formatting.add(
    f"J{D_FIRST}:J{D_FIRST + D_ROWS - 1}",
    FormulaRule(formula=[f'AND($B{D_FIRST}<>"",$F{D_FIRST}<>"완료",ISNUMBER($J{D_FIRST}),$J{D_FIRST}<0)'],
                fill=PatternFill("solid", fgColor=C_DELAY),
                font=Font(name=FONT, color=C_DELAY_FG, bold=True)))
ws_v.conditional_formatting.add(
    f"J{D_FIRST}:J{D_FIRST + D_ROWS - 1}",
    FormulaRule(formula=[f'AND($B{D_FIRST}<>"",$F{D_FIRST}<>"완료",ISNUMBER($J{D_FIRST}),$J{D_FIRST}>=0,$J{D_FIRST}<=7)'],
                fill=PatternFill("solid", fgColor="FFE4C4"),
                font=Font(name=FONT, color="C55A11", bold=True)))

ws_v.freeze_panes = f"B{D_FIRST}"

# =====================================================================
# 4. 간트차트 시트 (주 단위, 조건부 서식 기반)
# =====================================================================
ws_g = wb.create_sheet("간트차트")
ws_g.sheet_view.showGridLines = False

ws_g["B2"] = "📊 간트차트 (주 단위 타임라인)"
ws_g["B2"].font = Font(name=FONT, bold=True, size=14, color=C_TITLE)
ws_g["B3"] = "F2 셀의 시작 기준일을 바꾸면 전체 타임라인이 이동합니다. 프로젝트 시트 데이터가 자동 반영됩니다."
ws_g["B3"].font = Font(name=FONT, size=9, color="808080")

ws_g["E2"] = "기준일:"
ws_g["E2"].font = Font(name=FONT, bold=True, size=10)
ws_g["E2"].alignment = Alignment(horizontal="right")
ws_g["F2"] = dt.date(2026, 5, 4)  # 월요일
ws_g["F2"].number_format = "yyyy-mm-dd"
ws_g["F2"].font = Font(name=FONT, bold=True, size=10, color=C_ACCENT)
ws_g["F2"].fill = PatternFill("solid", fgColor=C_LIGHT)
ws_g["F2"].border = BORDER

N_WEEKS = 26
GH = 5           # 간트 헤더 행
G_FIRST = GH + 1
N_PROJ = 30      # 프로젝트 시트에서 참조할 행 수

# 왼쪽 고정 열
gantt_left = [("B", "프로젝트명", 26), ("C", "상태", 9), ("D", "시작일", 11), ("E", "마감일", 11)]
for col, title, w in gantt_left:
    cell = ws_g[f"{col}{GH}"]
    cell.value = title
    header_style(cell)
    ws_g.column_dimensions[col].width = w
ws_g.column_dimensions["A"].width = 2
ws_g.row_dimensions[GH].height = 30

# 주차 헤더 (F열부터): 각 주의 월요일 날짜
for wk in range(N_WEEKS):
    col = get_column_letter(6 + wk)
    cell = ws_g[f"{col}{GH}"]
    cell.value = f"=$F$2+{wk * 7}"
    cell.number_format = "mm/dd"
    header_style(cell)
    ws_g.column_dimensions[col].width = 5.5

# 프로젝트 참조 행
for i in range(N_PROJ):
    r = G_FIRST + i
    src = FIRST + i  # 프로젝트 시트의 데이터 행
    ws_g[f"B{r}"] = f'=IF(프로젝트!$C{src}="","",프로젝트!$C{src})'
    ws_g[f"C{r}"] = f'=IF(프로젝트!$C{src}="","",프로젝트!$G{src})'
    ws_g[f"D{r}"] = f'=IF(프로젝트!$C{src}="","",프로젝트!$H{src})'
    ws_g[f"E{r}"] = f'=IF(프로젝트!$C{src}="","",프로젝트!$I{src})'
    ws_g[f"D{r}"].number_format = "yyyy-mm-dd"
    ws_g[f"E{r}"].number_format = "yyyy-mm-dd"
    for col in ("B", "C", "D", "E"):
        body_style(ws_g[f"{col}{r}"], horizontal="left" if col == "B" else "center")
    for wk in range(N_WEEKS):
        col = get_column_letter(6 + wk)
        cell = ws_g[f"{col}{r}"]
        cell.border = Border(left=Side(style="thin", color="E7EDF5"),
                             right=Side(style="thin", color="E7EDF5"),
                             top=Side(style="thin", color="E7EDF5"),
                             bottom=Side(style="thin", color="E7EDF5"))

bar_rng = f"F{G_FIRST}:{get_column_letter(5 + N_WEEKS)}{G_FIRST + N_PROJ - 1}"
# 완료 프로젝트 막대 (초록)
ws_g.conditional_formatting.add(
    bar_rng,
    FormulaRule(formula=[f'AND($B{G_FIRST}<>"",$C{G_FIRST}="완료",F${GH}+6>=$D{G_FIRST},F${GH}<=$E{G_FIRST})'],
                fill=PatternFill("solid", fgColor="70AD47"), stopIfTrue=True))
# 지연 프로젝트 막대 (빨강)
ws_g.conditional_formatting.add(
    bar_rng,
    FormulaRule(formula=[f'AND($B{G_FIRST}<>"",$C{G_FIRST}="지연",F${GH}+6>=$D{G_FIRST},F${GH}<=$E{G_FIRST})'],
                fill=PatternFill("solid", fgColor="C00000"), stopIfTrue=True))
# 일반 진행 막대 (파랑)
ws_g.conditional_formatting.add(
    bar_rng,
    FormulaRule(formula=[f'AND($B{G_FIRST}<>"",F${GH}+6>=$D{G_FIRST},F${GH}<=$E{G_FIRST})'],
                fill=PatternFill("solid", fgColor=C_ACCENT), stopIfTrue=True))
# 이번 주 열 하이라이트 (연한 노랑, 막대 없는 셀만)
ws_g.conditional_formatting.add(
    bar_rng,
    FormulaRule(formula=[f'AND(TODAY()>=F${GH},TODAY()<F${GH}+7)'],
                fill=PatternFill("solid", fgColor="FFF2CC")))

# 간트 상태 열 색상
for name, bg, fg in status_colors:
    ws_g.conditional_formatting.add(
        f"C{G_FIRST}:C{G_FIRST + N_PROJ - 1}",
        CellIsRule(operator="equal", formula=[f'"{name}"'],
                   fill=PatternFill("solid", fgColor=bg),
                   font=Font(name=FONT, color=fg, bold=True)))

ws_g.freeze_panes = f"F{G_FIRST}"

# =====================================================================
# 5. 대시보드 시트
# =====================================================================
ws_d = wb.create_sheet("대시보드", 0)  # 첫 번째 위치
ws_d.sheet_view.showGridLines = False

ws_d["B2"] = "📊 프로젝트 대시보드"
ws_d["B2"].font = Font(name=FONT, bold=True, size=18, color=C_TITLE)
ws_d["B3"] = '=TEXT(TODAY(),"yyyy년 mm월 dd일")&" 기준 · 프로젝트 시트 데이터가 자동 집계됩니다."'
ws_d["B3"].font = Font(name=FONT, size=10, color="808080")

P_ID = f"프로젝트!$B${FIRST}:$B${FIRST + MAX_ROWS - 1}"
P_STATUS = f"프로젝트!$G${FIRST}:$G${FIRST + MAX_ROWS - 1}"
P_PRI = f"프로젝트!$F${FIRST}:$F${FIRST + MAX_ROWS - 1}"
P_PROG = f"프로젝트!$J${FIRST}:$J${FIRST + MAX_ROWS - 1}"
P_END = f"프로젝트!$I${FIRST}:$I${FIRST + MAX_ROWS - 1}"
T_ID = f"작업!$B${FIRST}:$B${FIRST + MAX_ROWS - 1}"
T_STATUS = f"작업!$G${FIRST}:$G${FIRST + MAX_ROWS - 1}"

# ---- KPI 카드 ----
kpis = [
    ("B", "전체 프로젝트", f'=COUNTIF({P_ID},"<>")', C_HEADER_BG, "FFFFFF"),
    ("D", "진행중", f'=COUNTIF({P_STATUS},"진행중")', C_ACCENT, "FFFFFF"),
    ("F", "완료", f'=COUNTIF({P_STATUS},"완료")', "70AD47", "FFFFFF"),
    ("H", "지연", f'=COUNTIF({P_STATUS},"지연")', "C00000", "FFFFFF"),
    ("J", "마감 초과(미완료)", f'=COUNTIFS({P_ID},"<>",{P_STATUS},"<>완료",{P_END},"<"&TODAY())', "ED7D31", "FFFFFF"),
    ("L", "평균 진행률", f'=IFERROR(AVERAGEIF({P_ID},"<>",{P_PROG}),0)', "7030A0", "FFFFFF"),
]
KPI_R1, KPI_R2 = 5, 6
for col, label, formula, bg, fg in kpis:
    col2 = get_column_letter(ws_d[col + "1"].column + 1)
    ws_d.merge_cells(f"{col}{KPI_R1}:{col2}{KPI_R1}")
    ws_d.merge_cells(f"{col}{KPI_R2}:{col2}{KPI_R2}")
    lab = ws_d[f"{col}{KPI_R1}"]
    lab.value = label
    lab.font = Font(name=FONT, size=10, color=fg, bold=True)
    lab.fill = PatternFill("solid", fgColor=bg)
    lab.alignment = Alignment(horizontal="center", vertical="center")
    val = ws_d[f"{col}{KPI_R2}"]
    val.value = formula
    val.font = Font(name=FONT, size=20, color=fg, bold=True)
    val.fill = PatternFill("solid", fgColor=bg)
    val.alignment = Alignment(horizontal="center", vertical="center")
    if col == "L":
        val.number_format = "0%"
    # 병합 셀 배경 채우기
    ws_d[f"{col2}{KPI_R1}"].fill = PatternFill("solid", fgColor=bg)
    ws_d[f"{col2}{KPI_R2}"].fill = PatternFill("solid", fgColor=bg)
ws_d.row_dimensions[KPI_R1].height = 20
ws_d.row_dimensions[KPI_R2].height = 34
for c in "BCDEFGHIJKLM":
    ws_d.column_dimensions[c].width = 11
ws_d.column_dimensions["A"].width = 2

# ---- 상태별 현황 ----
ws_d["B9"] = "상태별 프로젝트 수"
ws_d["B9"].font = Font(name=FONT, bold=True, size=12, color=C_TITLE)
hdr_r = 10
for c, t in (("B", "상태"), ("C", "건수"), ("D", "비율")):
    cell = ws_d[f"{c}{hdr_r}"]
    cell.value = t
    header_style(cell)
total_ref = f"COUNTIF({P_ID},\"<>\")"
for i, (name, bg, fg) in enumerate(status_colors):
    r = hdr_r + 1 + i
    ws_d[f"B{r}"] = name
    ws_d[f"B{r}"].fill = PatternFill("solid", fgColor=bg)
    ws_d[f"B{r}"].font = Font(name=FONT, color=fg, bold=True, size=10)
    ws_d[f"B{r}"].alignment = Alignment(horizontal="center", vertical="center")
    ws_d[f"B{r}"].border = BORDER
    ws_d[f"C{r}"] = f'=COUNTIF({P_STATUS},"{name}")'
    ws_d[f"D{r}"] = f'=IFERROR(C{r}/{total_ref},0)'
    ws_d[f"D{r}"].number_format = "0%"
    body_style(ws_d[f"C{r}"]); body_style(ws_d[f"D{r}"])
ws_d.conditional_formatting.add(
    f"C{hdr_r + 1}:C{hdr_r + 5}",
    DataBarRule(start_type="num", start_value=0, end_type="max",
                color=C_ACCENT, showValue=True))

# ---- 우선순위별 현황 ----
ws_d["F9"] = "우선순위별 프로젝트 수"
ws_d["F9"].font = Font(name=FONT, bold=True, size=12, color=C_TITLE)
for c, t in (("F", "우선순위"), ("G", "건수")):
    cell = ws_d[f"{c}{hdr_r}"]
    cell.value = t
    header_style(cell)
pri_colors = [("높음", "C00000"), ("중간", "ED7D31"), ("낮음", "70AD47")]
for i, (name, color) in enumerate(pri_colors):
    r = hdr_r + 1 + i
    ws_d[f"F{r}"] = name
    ws_d[f"F{r}"].font = Font(name=FONT, color=color, bold=True, size=10)
    ws_d[f"F{r}"].alignment = Alignment(horizontal="center", vertical="center")
    ws_d[f"F{r}"].border = BORDER
    ws_d[f"G{r}"] = f'=COUNTIF({P_PRI},"{name}")'
    body_style(ws_d[f"G{r}"])

# ---- 작업 현황 ----
ws_d["I9"] = "작업(태스크) 현황"
ws_d["I9"].font = Font(name=FONT, bold=True, size=12, color=C_TITLE)
for c, t in (("I", "구분"), ("J", "건수")):
    cell = ws_d[f"{c}{hdr_r}"]
    cell.value = t
    header_style(cell)
task_rows = [
    ("전체 작업", f'=COUNTIF({T_ID},"<>")'),
    ("진행중", f'=COUNTIF({T_STATUS},"진행중")'),
    ("완료", f'=COUNTIF({T_STATUS},"완료")'),
    ("지연", f'=COUNTIF({T_STATUS},"지연")'),
]
for i, (label, formula) in enumerate(task_rows):
    r = hdr_r + 1 + i
    ws_d[f"I{r}"] = label
    body_style(ws_d[f"I{r}"])
    ws_d[f"J{r}"] = formula
    body_style(ws_d[f"J{r}"])

# ---- 마감 임박 프로젝트 (수동 확인 안내 + 수식 목록) ----
ws_d["B18"] = "⚠️ 확인이 필요한 프로젝트"
ws_d["B18"].font = Font(name=FONT, bold=True, size=12, color="C00000")
ws_d["B19"] = "프로젝트 시트에서 '남은 일수'가 빨간색(마감 초과) 또는 주황색(7일 이내)으로 표시된 항목을 확인하세요."
ws_d["B19"].font = Font(name=FONT, size=10, color="595959")

# ---- 사용 안내 ----
guide = [
    "📌 사용 방법",
    "1. [프로젝트] 시트에 프로젝트를 등록합니다. 상태·우선순위·분류·담당자는 드롭다운으로 선택하세요.",
    "2. [작업] 시트에 프로젝트별 세부 작업을 등록하고 프로젝트 ID로 연결합니다.",
    "3. [프로젝트 상세] 시트에서 프로젝트를 선택하면 해당 프로젝트의 정보와 작업만 모아서 볼 수 있습니다.",
    "4. [간트차트] 시트에서 전체 일정을 한눈에 확인합니다. (파랑=진행, 초록=완료, 빨강=지연, 노랑 열=이번 주)",
    "5. 이 대시보드는 자동으로 집계됩니다. 진행률은 프로젝트 시트에서 직접 입력하세요.",
    "6. [설정] 시트에서 팀원·분류 등 드롭다운 목록을 수정할 수 있습니다.",
]
for i, line in enumerate(guide):
    r = 22 + i
    ws_d[f"B{r}"] = line
    ws_d[f"B{r}"].font = Font(name=FONT, size=10,
                              bold=(i == 0),
                              color=C_TITLE if i == 0 else "404040")

# 시트 순서 정리: 설정 시트를 맨 뒤로
wb.move_sheet("설정", offset=len(wb.sheetnames) - 1 - wb.sheetnames.index("설정"))

wb.save("프로젝트_트랙커.xlsx")
print("완료: 프로젝트_트랙커.xlsx")
