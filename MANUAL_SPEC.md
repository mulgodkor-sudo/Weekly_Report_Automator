# MANUAL_SPEC — Weekly Report Automator

> 이 문서는 소스 파일을 직접 분석하여 작성한 사양서입니다.
> 모든 인용은 `파일경로:라인번호` 형식입니다. 확인되지 않은 항목은 **(미확인)** 으로 표기합니다.

---

## 1. 개요

| 항목 | 내용 |
|------|------|
| 프로그램명 | Weekly Report Automator |
| 버전 문자열 | `config.json`의 `version` 키 기준 (`src/config.py:41`) |
| 현재 버전 | `"1.11"` (`src/config.py:41`) |
| 주 기능 | Microsoft Outlook 캘린더에서 [실적]/[계획] 태그 일정을 읽어 Weekly Report Excel 파일 자동 생성 |
| 진입점 | `main.py` (PyInstaller 빌드 시 exe에 고정) |
| UI 프레임워크 | Python 3.13 + tkinter |
| Excel 출력 | xlsxwriter |
| Outlook 연동 | win32com.client (pywin32) |
| FC 파일 읽기 | openpyxl |

---

## 2. 시스템 요구사항

- **OS**: Windows 10/11 (`README.md:120`)
- **Python**: 3.13 (`README.md:119`)
- **Outlook**: 데스크탑 버전, 로그인된 상태에서 실행 중이어야 함 (`src/outlook_reader.py:20`)
  - `win32com.client.GetActiveObject("Outlook.Application")` 방식 사용 — Outlook이 종료된 상태에서는 오류 발생
- **패키지**: `xlsxwriter`, `pywin32`, `openpyxl` (`README.md:121`)

---

## 3. 파일 구조

```
Weekly_Report_Automator/
├── main.py                    ← PyInstaller 진입점 (패치 불가, exe에 고정)
├── build.bat                  ← PyInstaller 빌드 스크립트
├── patches/
│   ├── apply_patch_v1.11.bat  ← v1.1 → v1.11 패치
│   └── README.md              ← 패치 사용 안내
└── src/
    ├── app.py                 ← 메인 UI (tkinter)
    ├── config.py              ← 설정 관리 + FC 데이터 로더
    ├── fc_rules.py            ← 구분(C열) 분류 규칙
    ├── event_processor.py     ← Outlook 이벤트 → Excel 행 변환
    ├── excel_writer.py        ← xlsxwriter 기반 Excel 출력
    ├── outlook_reader.py      ← Outlook COM 연동
    ├── overrides.py           ← 되풀이 모임 override 저장/적용
    ├── override_dialog.py     ← 되풀이 모임 설정 UI
    ├── plant_mh_dialog.py     ← Plant M/H 입력 확인 UI
    ├── monthly_processor.py   ← 월간업무정리 처리 로직
    ├── monthly_dialog.py      ← 월간업무정리 UI + Excel 저장
    ├── config.json            ← 초기 설정값 (번들)
    ├── overrides.json         ← 초기 override 데이터 (번들)
    └── assets/
        ├── splash.png
        └── Schedule_Ico.ico
```

---

## 4. 설정 파일

### 4-1. config.json (사용자 설정)

- **저장 위치**: `C:\Users\{사용자명}\Documents\WeeklyReportAutomaker_config.json` (`src/config.py:32`)
- **번들 초기값**: `src/config.json` → 최초 실행 시 위 경로로 복사 (`src/config.py:58`)
- **구조**:

```json
{
  "func_code_excel_path": "",
  "splash_image": "assets/splash.png",
  "icon_file":    "assets/Schedule_Ico.ico",
  "version":      "1.11"
}
```

| 키 | 의미 |
|----|------|
| `func_code_excel_path` | FC 엑셀 파일 전체 경로 (사용자가 설정) |
| `splash_image` | 스플래시 이미지 상대경로 (번들) |
| `icon_file` | 아이콘 파일 상대경로 (번들) |
| `version` | 화면 표시용 버전 문자열 (`src/config.py:41`) |

### 4-2. overrides.json (되풀이 모임 설정)

- **저장 위치**: `C:\Users\{사용자명}\Documents\WeeklyReportAutomaker_overrides.json` (`src/overrides.py:17–25`)
- **번들 초기값**: `src/overrides.json` (예시 2건 포함)
- **구조**: 배열, 각 항목:

```json
{
  "project_code": "000000",
  "func_code":    "GA11-25",
  "subject":      "Team Meeting",
  "body":         "팀 미팅 진행"
}
```

---

## 5. 아웃룩 일정 작성 규칙

### 5-1. 제목 형식

```
[실적] 업무명 [#프로젝트코드#기능코드]
[계획] 업무명 [#프로젝트코드#기능코드]
```

- 앞쪽 태그: `[실적]` 또는 `[계획]` (`src/outlook_reader.py:75–97`)
- 코드 블록: `[#PC#FC]` 형식, PC는 5~8자 영숫자, FC는 `AA00-00` 패턴 (`src/outlook_reader.py:77`)
  - 정규식: `[A-Za-z0-9]{5,8}` (PC), `[A-Z]{2}\d{2}-\d{2}` (FC)
  - **공백 허용**: `[# 000000 # GA11 - 25]` 형식도 파싱됨 — 공백은 파싱 시 자동 제거

### 5-2. 시간 계산

| 조건 | 계산 방법 | 근거 |
|------|----------|------|
| AllDayEvent | 8.0H 고정 | `src/outlook_reader.py:116` |
| 시간 지정 일정 | (종료 - 시작) 시간, 30분 단위 반올림 | `src/outlook_reader.py:118` |

- 30분 단위 반올림: `round(h * 2) / 2` (`src/outlook_reader.py:118`)

### 5-3. 본문(상세) 처리

- Outlook HTMLBody → 텍스트 변환 우선, 실패 시 Body 사용 (`src/outlook_reader.py:105–112`)
- `*` 기호부터 이후 내용은 Excel I열에서 **붉은색**으로 출력 (`src/excel_writer.py:142–165`)

---

## 6. 메인 UI (app.py)

### 6-1. 창 기본 속성

| 항목 | 값 | 근거 |
|------|-----|------|
| 초기 창 크기 | `780x320` | `src/app.py:32` |
| 최소 창 크기 | `(780, 320)` | `src/app.py:33` |
| 배경색 | `#F0F2F5` | `src/app.py:44` |
| 창 제목 | `f"Weekly Report Automator  {버전}"` | `src/app.py:40` |

### 6-2. 타이틀 바

- 배경: `#1F4E79` (네이비), 높이 52px (`src/app.py:212`)
- 좌측 텍스트: `f"  📊  Weekly Report Automator  {self._version}"` — 폰트 맑은 고딕 14pt Bold, 흰색 (`src/app.py:215`)
- 우측 텍스트: `"DL이앤씨  플랜트본부 기계설계팀  "` — 폰트 맑은 고딕 9pt, `#9DC3E6` (`src/app.py:218`)

### 6-3. 상태바 (하단)

- 배경: `#1F4E79` (네이비), 높이 30px (`src/app.py:423`)
- 좌측: 동적 상태 메시지 (`self.v_status`, 초기값 `"  준비"`) (`src/app.py:64`)
- 우측 고정 텍스트: `"* 불편사항/개선사항은 이수신 차장에게 문의 바랍니다.  "` — 색상 `#9DC3E6` (`src/app.py:430`)

### 6-4. 섹션 구성 (위→아래)

1. **타이틀 바** (항상 표시)
2. **주별 추출** (좌) + **일별 추출** (우) — 항상 표시 (`src/app.py:186`)
3. **주요 버튼 행** — 항상 표시 (`src/app.py:190`)
4. **`▶  ⚠️   사전 검토 결과  ─  불러오기 후 표시`** — 접이식, 처음엔 접힘, 불러오기 후 자동 펼침 (`src/app.py:193–194`)
5. **`▶  🔁  되풀이 모임 설정 (계획)`** — 접이식, 처음엔 접힘 (`src/app.py:198–199`)
6. **`▶  ⚙️   상세 설정  (FC 파일 / 저장 경로)`** — 접이식, 처음엔 접힘 (`src/app.py:202–203`)
7. **상태바** (항상 표시)

접이식 섹션 헤더: 배경 `#D9E1F2`, 글씨 `#1F4E79`, 클릭 시 화살표 `▶`/`▼` 토글, 창 높이 자동 조정 최대 900px (`src/app.py:120–172`)

### 6-5. 주별 추출 섹션

- 섹션 제목: `"  📅  주별 추출  (Weekly Report 핵심 기능)"` (`src/app.py:245`)
- 체크박스: `"이번주  [실적] :"`, `"다음주  [계획] :"` (`src/app.py:255, 274`)
- 이번주 시작일 입력창: 강조 배경 `#EFF6FF`, 글씨 `#1F4E79`, 폰트 맑은 고딕 14pt Bold (`src/app.py:260–265`)
  - **자동완성**: 이번주 시작일(YYYY-MM-DD 10자) 입력 완료 시 나머지 날짜(종료, 다음주 시작/종료) 자동 계산 (`src/app.py:67–82`)
- 버튼:
  - `"↺ 날짜 초기화"` — 오늘 기준 이번주 월~금 자동 설정 (`src/app.py:297`)
  - `"🔍   주별 불러오기"` — 네이비 배경, Outlook에서 해당 주 일정 로드 (`src/app.py:304`)
- 힌트 텍스트: `"YYYY-MM-DD  |  이번주 시작일 입력 시 나머지 날짜 자동완성"` (`src/app.py:313`)

### 6-6. 일별 추출 섹션

- 섹션 제목: `"  🗓  일별 추출 (특정 날짜)"` (`src/app.py:224`)
- 날짜 입력: `"날  짜 :"`, 형식 힌트 `"(YYYY-MM-DD)"` (`src/app.py:230, 236`)
- 버튼: `"🔍  일별 불러오기"` (`src/app.py:238`)
- 기대 시간: 8.0H (일별 모드) (`src/app.py:540`)

### 6-7. 주요 버튼 행

| 버튼 텍스트 | 색상 | 초기 상태 | 근거 |
|------------|------|---------|------|
| `"📄   Excel 파일 생성"` | 배경 `#1F5E20` (초록) | **비활성** | `src/app.py:347–356` |
| `"📊  Plant M/H 입력 확인"` | 배경 `#5D4037` (갈색) | **비활성** | `src/app.py:360–370` |
| `"📋  월간업무정리"` | 배경 `#4A148C` (보라) | **항상 활성** | `src/app.py:373–380` |

- Excel 생성 / Plant M/H 버튼은 불러오기 성공 후 활성화됨 (`src/app.py:636–637`)
- 인디케이터(진행바): 불러오기/생성 중 indeterminate 모드 회전 (`src/app.py:382`)

### 6-8. 사전 검토 결과 섹션

- 텍스트 영역: 배경 `#FFFDE7` (연노랑), 높이 8줄, read-only (`src/app.py:319–323`)
- 초기 메시지:
  ```
  ⬅  [일별/주별 불러오기] 버튼을 클릭하면 아웃룩 일정을 가져옵니다.

  * Excel 생성 버튼을 누르면 Weekly Report 형식으로 출력됩니다.
  ```
  (`src/app.py:329–332`)
- 불러오기 성공 시 표시 내용:
  - `"✅  이번주 [실적] : {N}건  (합계 {H:.1f}H)"` (`src/app.py:624`)
  - `"✅  다음주 [계획] : {N}건  (합계 {H:.1f}H)"` (`src/app.py:625`)
  - `"🔁  되풀이 모임 override 적용 : {N}건"` (적용 건수 있을 때) (`src/app.py:626`)
  - 경고 항목 (아래 7절 참조)
  - 경고 없으면: `"✅  모든 검토 항목 이상 없음"` (`src/app.py:631`)

### 6-9. 되풀이 모임 설정 섹션 (접이식)

- 섹션 제목: `"  🔁  되풀이 모임 설정 (계획)"` (`src/app.py:199`)
- 내부 버튼: `"🔁   되풀이 모임 내용 고정 설정 열기"` — 배경 `#37474F` (`src/app.py:447`)
- 안내 텍스트: `"  [계획]으로만 새로 추가되는 되풀이 약속의 상세내용을 고정 설정합니다."` (`src/app.py:455`)
- 클릭 시 `OverrideDialog` 열림 (`src/app.py:440, 447`)

### 6-10. 상세 설정 섹션 (접이식)

#### FC 엑셀 파일 서브섹션

- 섹션 제목: `"  📂  Function Code 엑셀 파일"` (`src/app.py:389`)
- 레이블: `"파일 경로 :"` (`src/app.py:394`)
- 버튼: `"찾아보기"`, `"🔄 FC 재로드"` (`src/app.py:396, 398`)
- 힌트: `'파일명에 "기계설계" 포함 | 읽기 전용 | 수식 결과값 기준'` (`src/app.py:400`)
- FC 파일 조건: 파일명에 `"기계설계"` 포함, `.xlsx` 또는 `.xlsm` 확장자 (`src/config.py:85–89`)

#### 저장 설정 서브섹션

- 섹션 제목: `"  💾  저장 설정"` (`src/app.py:406`)
- 레이블: `"저장 경로 :"`, `"파일 이름 :"` (`src/app.py:411, 416`)
- 힌트: `".xlsx 자동 추가"` (`src/app.py:418`)
- 기본 저장 경로: 사용자 바탕화면 (`src/app.py:98–99`)
- 기본 파일명 패턴:
  - 주별: `f"WeeklyReport_{월}월{주차}주차.xlsx"` (예: `WeeklyReport_7월2주차.xlsx`) (`src/app.py:27`)
  - 일별: `f"WeeklyReport_{YYYYMMDD}.xlsx"` (`src/app.py:613`)

---

## 7. 사전 검토 (경고 생성)

`src/event_processor.py:194–301`의 `get_warnings()` 함수가 생성하는 경고:

### 7-1. 시간 합계 경고

- 허용 오차: `AUTO_TOL = 0.25` (15분) (`src/event_processor.py:212`)
- 기대값: 주별=40.0H, 일별=8.0H (`src/event_processor.py:197`)

| 조건 | 출력 형식 |
|------|---------|
| 차이 = 0 | 경고 없음 |
| 0 < 차이 ≤ 0.25H | `"ℹ️  {레이블} 합계 {H:.2f}H  →  {기대}H 자동보정 (30분 단위 반올림 오차)"` |
| 합계 < 기대 | `"⚠️  {레이블} 합계 {H:.1f}H  ←  {기대}H 미만. 시간 재확인 필요"` |
| 합계 > 기대 | `"⚠️  {레이블} 합계 {H:.1f}H  ←  {기대}H 초과. 시간 재확인 필요"` |

### 7-2. 코드 누락 경고

- 형식: `"⚠️  {mm월 dd일} - {프로젝트 코드 / Function Code} 없음. 재확인 필요  ※ {제목 앞 35자}"` (`src/event_processor.py:256–258`)
- 같은 (날짜, 제목) 조합은 중복 경고 제거 (`src/event_processor.py:248–249`)

### 7-3. 미등록 FC 경고

- 형식: `"⚠️  {mm월 dd일} - {FC코드} - 해당 Code는 존재하지 않음. 재확인 필요"` (`src/event_processor.py:272–274`)

### 7-4. FC 개수 정보 (항상 표시, 경고 목록 첫 줄)

- 형식: `"ℹ️  Function Code: 총 {N}개  (General {N}개 / Project {N}개)"` (`src/event_processor.py:292–295`)
- FC 2개 이하일 때 추가 경고: `"⚠️  Function Code {N}개 (3개 이상 되도록 점검 요망)"` (`src/event_processor.py:296–299`)

### 7-5. 아웃룩 접근 관련 (app.py에서 추가)

- FC 파일 미로드: `"⚠️  Function Code 파일이 로드되지 않았습니다.\n      하단 [FC 엑셀 파일] 경로를 확인하고 🔄FC 재로드를 눌러주세요."` (`src/app.py:566–568`)
- 0건 (일정 없음): `"⚠️  아웃룩에서 해당 기간({시작}~{종료}) 일정을 찾지 못했습니다.\n..."` (`src/app.py:572–577`)

---

## 8. 이벤트 처리 로직 (event_processor.py)

### 8-1. 이번주 [실적] 처리

1. `tag == "실적"` 필터 (`src/event_processor.py:55–58`)
2. **같은날 병합**: 날짜 + 제목 + 본문 동일 → 시간 합산, 단일 행 (`src/event_processor.py:31–47`)
3. `build_rows()`에서 (func_code, subject, body) 기준 그룹화, 그룹 내 날짜 오름차순 정렬 (`src/event_processor.py:110–145`)

### 8-2. 다음주 [계획] 처리

1. `tag == "계획"` 필터 (`src/event_processor.py:75–76`)
2. **(project_code, func_code, subject) 동일** → 시간 합산 (`src/event_processor.py:79–87`)
3. subject가 다르면 별도 행

### 8-3. 이번주↔다음주 연결

- 정규화된 subject 비교: `[실적/계획]` 태그 + `[#PC#FC]` 코드 블록 제거 후 비교 (`src/event_processor.py:20–24`)
- 매칭 시 이번주 첫 번째 행의 `next_week_h` 에 합산 (`src/event_processor.py:162–165`)
- 매칭 안 된 [계획] → `source = "next_week_only"` 신규 행 추가 (`src/event_processor.py:176–185`)

---

## 9. 구분(C열) 분류 로직 (fc_rules.py)

`GUBUN_OPTIONS = ["수행 (KPI)", "입찰", "기타 (KPI)", "기타"]` (`src/fc_rules.py:13`)

`GENERAL_PC = "000000"` (`src/fc_rules.py:10`)

| 조건 | 결과 | 근거 |
|------|------|------|
| FC 정보 없음 | `"기타"` | `src/fc_rules.py:26–27` |
| `kpi == "conditional"` + FC = GA08-01 + "매뉴얼" in 제목/본문 | `"기타 (KPI)"` | `src/fc_rules.py:30–33` |
| `kpi == "conditional"` (그 외) | `"기타"` | `src/fc_rules.py:34` |
| `kpi == "conditional"` (GA08-01 아닌 나머지) | `"기타"` | `src/fc_rules.py:35–36` |
| `kpi == True` + 구분 = "Project" | `"수행 (KPI)"` | `src/fc_rules.py:38–39` |
| `kpi == True` + 구분 = "General" + PC = 000000 | `"기타 (KPI)"` | `src/fc_rules.py:40–42` |
| `kpi == True` + 구분 = "General" + PC ≠ 000000 | `"수행 (KPI)"` (외부 프로젝트 투입) | `src/fc_rules.py:43–45` |
| `kpi == False` | `"기타"` | `src/fc_rules.py:46–47` |

### 9-1. FC 엑셀 파일 헤더 인식

필수 컬럼: `구분`, `CODE`, `KPI`, `업무명` — 최초 15행 안에서 자동 감지 (`src/config.py:107–117`)

KPI 컬럼 파싱:
- `CONDITIONAL_CODES = {"GA11-25", "GA08-01"}` → 항상 `kpi = "conditional"` (`src/config.py:77`)
- `KPI` 셀에 `/` 포함 + `O` 포함 (예: "O (LE미팅) / X(그룹리더/팀미팅)") → `kpi = "conditional"` (`src/config.py:136–138`)
- `KPI == "O"` → `kpi = True` (`src/config.py:139–140`)
- 그 외 → `kpi = False` (`src/config.py:141–142`)

---

## 10. Excel 출력 구조 (excel_writer.py)

### 10-1. 시트 구성

| 시트명 | 내용 | 근거 |
|--------|------|------|
| `Weekly Report` | 주간보고서 메인 시트 | `src/excel_writer.py:285` |
| `Plant MH` | 날짜별 실적 시간 집계 | `src/excel_writer.py:557` |

### 10-2. Weekly Report 시트 레이아웃

- 데이터 시작 행: 14행 (0-indexed: 13, `DATA_ROW = 13`) (`src/excel_writer.py:15`)
- 데이터 최대 행: 82행 (0-indexed: 81, `DATA_EROW = 81`) (`src/excel_writer.py:16`)

| 열 | 내용 | 정렬 |
|----|------|------|
| A (0) | (빈칸) | — |
| B (1) | (빈칸) | — |
| C (2) | 구분 | 가운데, 드롭다운 |
| D (3) | Project Code | 가운데 |
| E (4) | Function Code | 가운데 |
| F (5) | 업무명 (FC 기준) | 왼쪽 |
| G (6) | Date (`yyyy-mm-dd`) | 왼쪽 |
| H (7) | 수행 업무 (제목) | 왼쪽 |
| I (8) | 수행 업무 (상세) | 왼쪽, 병합 가능 |
| J (9) | 금주 시간 (`0.0`) | 가운데 |
| K (10) | 금주 비중 (`0%`) | 가운데, 수식 |
| L (11) | 구분선 (흰색) | — |
| M (12) | 차주 시간 (`0.0`) | 가운데, 검정 |
| N (13) | 차주 비중 (`0%`) | 가운데, 수식 |
| O (14) | 보직자 Feedback | 빈칸 |
| P (15) | 비고 | 빈칸 |

### 10-3. 헤더(1행)

- C~D 병합: `=IFERROR(TEXT(G14,"m")&"월 "&INT((DAY(G14)-1)/7)+1&"주차","")` — 주차 자동 계산 (`src/excel_writer.py:326–328`)
- E열: `"Weekly Report"` (`src/excel_writer.py:329`)
- O~P 병합: `"플랜트사업본부 설계실"` (`src/excel_writer.py:332`)

### 10-4. M/H 집계 영역 (3~9행)

- I열: `"주간 M/H 실적 집계"` (`src/excel_writer.py:335`)
- J~K 병합: `"금주"` (`src/excel_writer.py:336`)
- M~N 병합: `"차주"` (`src/excel_writer.py:337`)
- 카테고리 4개 (4~7행): `'수행 (KPI)'`, `'입찰'`, `'기타 (KPI)'`, `'기타'` — SUMIF 수식 (`src/excel_writer.py:340`)
- 9행 총계: `"총   계"` — SUM 수식 (`src/excel_writer.py:357`)

### 10-5. 컬럼 헤더 (12~13행)

- 12행 주헤더: `구분`, `Project\nCode`, `Function Code`, `Date`, `수행 업무 (제목)`, `수행 업무 (상세)`, `금주`, `차주`, `보직자 Feedback`, `비고` (`src/excel_writer.py:365–386`)
- 13행 서브헤더: `Code`, `Description` (FC), `시간`, `비중` (금주/차주) (`src/excel_writer.py:381–386`)
- 헤더 배경: `#595959` (진회색), 글씨: `#FFFFFF` (흰색) (`src/excel_writer.py:38–39`)

### 10-6. 데이터 행 정렬 규칙

1. 이번주 실적 행 먼저, 다음주만 있는 행(next_week_only) 나중에 (`src/excel_writer.py:278–280`)
2. 각각 내부: 구분 순서(수행KPI→입찰→기타KPI→기타) → 그룹 내 최초 날짜 → PC → FC → 제목 → 본문 (`src/excel_writer.py:250–251`)
3. **1행 보장**: 전체 최초 날짜를 포함한 그룹을 맨 앞으로 배치 (그룹 단위 이동, 병합 유지) (`src/excel_writer.py:260–270`)
4. 그룹 내부: 날짜 오름차순 (`src/excel_writer.py:248`)

### 10-7. I열(상세) 병합

- 같은 (gubun, pc, fc, subject, 별표제거_body)가 **연속** 배치된 경우 I열 병합 (`src/excel_writer.py:454–480`)
- 별표(`*`) 이후 내용은 병합 키 계산에서 제외 (`src/excel_writer.py:212–215`)
- 병합 그룹 행 높이: 총 추정 높이를 행 수로 균등 분배 (`src/excel_writer.py:492–500`)

### 10-8. C열 드롭다운

```
source: ['수행 (KPI)', '입찰', '기타 (KPI)', '기타']
error_message: '목록에서 선택하세요'
```
(`src/excel_writer.py:434–438`)

### 10-9. 색상 규칙

| 요소 | 색상 | 근거 |
|------|------|------|
| 금주 시간 집계 라벨 | 검정 (`#000000`) | `src/excel_writer.py:37` |
| 차주 시간/비중 | 검정 (`#000000`) | `src/excel_writer.py:37` |
| 합계 행 배경 | `#FFE699` (연노랑) | `src/excel_writer.py:40` |
| 별표 이후 텍스트 | `#FF0000` (붉은색) | `src/excel_writer.py:42` |

### 10-10. Plant MH 시트

- 헤더: `구분`, `Project Code`, `Func. Code`, `Description` + 날짜별 열(`월/일(요일)`) + `합계` (`src/excel_writer.py:594–598`)
- 날짜 표시: `f"{월}/{일}\n({요일})"`, 요일 = `['Mon','Tue','Wed','Thu','Fri','Sat','Sun']` (`src/excel_writer.py:597`)
- 이번주 실적 데이터만 포함 (`src/excel_writer.py:559`)
- 합계 행 배경: `#FFE699` (`src/excel_writer.py:580`)

---

## 11. 되풀이 모임 설정 (override_dialog.py / overrides.py)

### 11-1. 적용 조건

되풀이 모임 override는 **`source == "next_week_only"`** 행에만 적용 (`src/overrides.py:42–43`)
- 이번주 실적에 이미 상세내용이 있는 경우는 적용 안 됨

### 11-2. 매칭 키

`(project_code, func_code, normalized_subject)` — 정규화: `[실적/계획]` 태그 + `[#PC#FC]` 블록 제거 (`src/overrides.py:35–36`)

### 11-3. 적용 효과

매칭 시 해당 행의 `body` 필드를 override 설정값으로 교체 (`src/overrides.py:44`)

### 11-4. 다이얼로그 (OverrideDialog)

- 창 제목: `"되풀이 모임 내용 고정 설정"` (`src/override_dialog.py:43`)
- 창 크기: `700x490`, 최소 `600x430` (`src/override_dialog.py:44–45`)
- 타이틀 바 텍스트: `"  🔁  되풀이 모임 내용 고정 설정"` (`src/override_dialog.py:64`)

#### 섹션 구성

| 섹션 | 제목 |
|------|------|
| 빠른 파싱 | `"  ⚡  빠른 파싱 (아웃룩 제목 붙여넣기 → 자동 분석)"` |
| 등록 목록 | `"  📋  등록된 항목 (클릭하면 수정)"` |
| 입력/수정 폼 | `"  ✏️  내용 입력 / 수정"` |

(`src/override_dialog.py:76–117`)

#### 폼 필드

| 레이블 | 필드 |
|--------|------|
| `"Project Code :"` | 텍스트 입력 |
| `"Function Code :"` | 텍스트 입력 |
| `"수행업무 (제목) :"` | 텍스트 입력 |
| `"수행업무 (상세) :"` | 3줄 텍스트 박스 |

(`src/override_dialog.py:125–152`)

#### 버튼 행

`"➕  새로 추가"`, `"💾  저장"`, `"🗑  삭제"`, `"✅  닫기"` (`src/override_dialog.py:163–176`)

- 힌트: `"제목: 태그·코드 제외한 업무명만"` (`src/override_dialog.py:171`)
- 저장 확인 메시지: `"저장되었습니다."` (`src/override_dialog.py:275`)

#### 빠른 파싱

- 버튼: `"🔍 분석"` (`src/override_dialog.py:89`)
- 예시 힌트: `"예) [실적] [General] Team Meeting [#000000#GA11-25]"` (`src/override_dialog.py:93`)
- 코드 미발견 시: 파싱 결과 메시지박스 표시 후 PC/FC 직접 입력 유도 (`src/override_dialog.py:192–198`)

---

## 12. Plant M/H 입력 확인 (plant_mh_dialog.py)

- 창 제목: `"Plant M/H 입력 확인"` (`src/plant_mh_dialog.py:20`)
- 창 크기: `920x420`, 최소 `700x300` (`src/plant_mh_dialog.py:21–22`)
- 타이틀 바 텍스트: `"  📊  Plant M/H 입력 확인  (이번주 [실적] 기준)"` (`src/plant_mh_dialog.py:41`)
- 데이터 없을 때 메시지: `"이번주 [실적] 데이터가 없습니다.\n먼저 주별 불러오기를 실행하세요."` (`src/plant_mh_dialog.py:48–49`)

### 컬럼 구성

| 컬럼 ID | 헤더 | 너비 |
|---------|------|------|
| no | No. | 40 |
| gubun | 구분 | 85 |
| pc | Project Code | 80 |
| fc | Func. Code | 80 |
| desc | Description | 180 |
| d0..dN | `요일(월/일)` 형식 | 60 |
| total | 합계 | 55 |

(`src/plant_mh_dialog.py:79–82`, `src/plant_mh_dialog.py:77`)

- 날짜 헤더 형식: `f"{DAY_KO[요일]}({월}/{일})"`, DAY_KO = `["월","화","수","목","금","토","일"]` (`src/plant_mh_dialog.py:14, 77`)
- 합계 행 레이블: `"합  계"` — 배경 `#FFE699` (`src/plant_mh_dialog.py:124–129`)
- 하단 안내: `"※ ST(정규시간) 기준  |  OT는 별도 확인하세요"` (`src/plant_mh_dialog.py:144–146`)
- 닫기 버튼: `"✅  닫기"` (`src/plant_mh_dialog.py:147`)

---

## 13. 월간업무정리 (monthly_dialog.py / monthly_processor.py)

### 13-1. 처리 로직

`monthly_processor.process_monthly()` (`src/monthly_processor.py:56–154`)

- 입력: `[실적]` 이벤트 (`src/monthly_processor.py:76`)
- **휴가/공가 FC**: `HOLIDAY_FC = "GE04-02"` — 별도 집계 (`src/monthly_processor.py:17`)
- 그룹화 키: `(project_code, func_code, normalized_subject)` (`src/monthly_processor.py:92`)
- 본문 별표 제거: `*` 이후 전체 삭제 (`src/monthly_processor.py:38–43`)
- **같은 날 + 같은 내용**: 시간 합산 (`src/monthly_processor.py:115`)
- **다른 날 + 같은 내용**: 날짜 목록에 추가 (별도 행 아님) (`src/monthly_processor.py:114`)

### 13-2. 정렬 순서

구분 우선순위: `수행(KPI)=0`, `입찰=1`, `기타(KPI)=2`, `기타=3` → 매뉴얼 류 인접 → PC → FC → 제목 (`src/monthly_processor.py:143–149`)

매뉴얼 류 키워드: `["표준서", "메뉴얼", "매뉴얼", "절차서"]` (`src/monthly_processor.py:18`)

### 13-3. 다이얼로그 (MonthlyDialog)

- 창 제목: `"월간 업무 정리"` (`src/monthly_dialog.py:32`)
- 창 크기: `1080x760`, 최소 `860x480` (`src/monthly_dialog.py:33–34`)
- 타이틀 바 텍스트: `"  📋  월간 업무 정리"` (`src/monthly_dialog.py:45`)

#### 기간 설정 섹션

- 섹션 제목: `"  📅  기간 설정"` (`src/monthly_dialog.py:55`)
- 레이블: `"시작일 :"`, `" ~ 종료일 :"` (`src/monthly_dialog.py:59, 63`)
- 자동 입력: 전월 1일~말일 (`src/monthly_dialog.py:155–162`)
- 수정 가능 형식: `YYYY-MM-DD` (`src/monthly_dialog.py:83`)
- 버튼: `"🔍  불러오기"` (`src/monthly_dialog.py:70`)
- 힌트: `"* 전월 날짜 자동 입력 / YYYY-MM-DD 형식으로 수정 가능 / Non-KPI 순서는 FC 엑셀 파일 기준"` (`src/monthly_dialog.py:82–83`)

#### 결과 Treeview 섹션

- 섹션 제목: `"  📊  결과  (같은 내용은 날짜 병합 / 별표 이후 내용 제거)"` (`src/monthly_dialog.py:88`)
- 컬럼: `구분`, `Project Code`, `FC`, `업무명`, `날짜`, `수행내용`, `시간` (`src/monthly_dialog.py:94`)
- 행 색상: odd=흰색, even=`#F0F4FF`, subtotal=`#FFE699`, holiday=`#FCE4EC`, grand_total=`#1F4E79`(흰 글씨) (`src/monthly_dialog.py:111–118`)
- 소계 레이블: `"소  계"` (`src/monthly_dialog.py:238`)
- 월간 총 합계 레이블: `"★  월간 총 합계"` (`src/monthly_dialog.py:253`)
- PC/FC 별 소요시간 헤더: `"[ PC / FC 별 소요시간 ]"` (`src/monthly_dialog.py:262`)
- 상태바 형식: `"업무 그룹 {N}건  |  휴가/공가 {H:.1f}H"` (`src/monthly_dialog.py:207`)

#### 하단 버튼

- `"✅  닫기"` — 배경 `#455A64` (`src/monthly_dialog.py:134–141`)
- `"📄  Excel 저장"` — 배경 `#1F5E20`, 초기 **비활성**, 불러오기 성공 후 활성화 (`src/monthly_dialog.py:143–153`)

### 13-4. Excel 저장 (월간업무정리)

- 기본 파일명: `f"월간업무정리_{년도}년{월}월.xlsx"` (예: `월간업무정리_2025년6월.xlsx`) (`src/monthly_dialog.py:295`)
- 저장 후 바로 열기 여부 묻는 대화상자 (`src/monthly_dialog.py:300–303`)

#### 시트 구성

| 시트명 | 내용 | 근거 |
|--------|------|------|
| `월간업무정리` | 상세 업무 목록 | `src/monthly_dialog.py:340` |
| `PC_FC별 소요시간` | (구분, PC, FC) 별 시간 요약 | `src/monthly_dialog.py:436` |

#### 월간업무정리 시트 컬럼

`구분`, `Project Code`, `FC`, `업무명`, `날짜`, `수행내용`, `시간(H)` (`src/monthly_dialog.py:375`)

- 타이틀 행: `f"월간 업무 정리  —  {년도}년 {월}월"` (`src/monthly_dialog.py:372`)
- 소계 레이블: `"소  계"` (`src/monthly_dialog.py:403`)
- 총합계 행: `f"★  {년도}년 {월}월  월간 총 업무 시간"` (`src/monthly_dialog.py:431`)

#### PC_FC별 소요시간 시트 컬럼

`구분`, `Project Code`, `FC`, `업무명(FC 기준)`, `시간(H)` (`src/monthly_dialog.py:454`)

- 합계 행 레이블: `"월간 총 업무 시간"` (`src/monthly_dialog.py:509`)

---

## 14. 빌드 및 패치

### 14-1. 빌드 (`build.bat`)

- PyInstaller `--onedir` 방식
- 출력: `dist/Weekly_Report_Automator_V1.11/Weekly_Report_Automator_V1.11.exe`
- `main.py`가 exe에 고정되며, 이후 패치로 갱신 불가 (`README.md:96–98`)
- `main.py`의 `_setup_src()`가 `exe_dir/src`를 `sys.path` 앞에 삽입 → `src/*.py` 패치 가능

### 14-2. 패치 (`patches/apply_patch_v1.11.bat`)

- 적용 위치: exe와 같은 폴더 (= `src/` 폴더가 있는 폴더)
- 방식: base64 → zip 디코딩 → `src/` 전체 교체 (부분 수정 아님)
- exe 파일명 `V1.1` → `V1.11` 자동 변경 (멱등: 이미 V1.11이면 건너뜀)
- 상위 폴더명 `V1.1` → `V1.11` 자동 변경 (`patches/README.md`)
- 임시파일 (`%TEMP%\wra_patch_*.b64`, `*.zip`) 성공/실패 무관 자동 삭제
- 백업 폴더 미생성

### 14-3. v1.11 패치 대상 파일

| 파일 | 변경 내용 |
|------|---------|
| `src/app.py` | 버전 표시 `config.json` 기반으로 변경, "DL이엔씨" → "DL이앤씨" 오타 수정 |
| `src/config.py` | `version` 키 추가, `get_version_str()` 헬퍼 추가 |
| `src/config.json` | `version: "1.11"` 추가 |
| `src/excel_writer.py` | 날짜 정렬 버그 수정, I열 병합 회귀 버그 수정, 테두리 누락 수정 |

---

## 15. 죽은 코드 (Dead Code / 비활성 기능)

| 위치 | 내용 | 상태 |
|------|------|------|
| `src/event_processor.py:138` | `highlight = False` (row dict에 항상 False 고정) | 노란 음영 제거됨, 필드는 남아있음 |
| `src/config.py:185–188` | `is_highlight()` — 항상 `return False` | 호출되더라도 기능 없음 |
| `src/excel_writer.py:91` | `'data_i_red'` format — I열 전체 붉은색 (`color=RED`) | 실제로는 `_write_body()` 내부에서 단색 붉은 셀에 사용됨, 활성 |
| `src/event_processor.py:141–144` | `is_group_first`, `is_group_last`, `group_size`, `group_key` 필드 | `build_rows()`에서 설정하나, `excel_writer.py`는 이를 사용하지 않고 자체 병합 로직 사용 — 사실상 미사용 |

---

## 16. 스레딩 모델

- 모든 Outlook COM 호출과 Excel 생성은 **백그라운드 스레드**에서 실행 (`src/app.py:542–596`)
- COM 스레드 초기화: `pythoncom.CoInitialize()` / 정리: `pythoncom.CoUninitialize()` (`src/app.py:549–552, 589–593`)
- UI 업데이트는 `self.root.after(0, callback)` 로 메인 스레드에서 처리 (`src/app.py:583, 586`)
- 월간업무정리 / 되풀이 모임 설정도 동일 패턴 적용 (`src/monthly_dialog.py:181–199`)

---

## 17. 오류 처리

### 알려진 오류 패턴

| 상황 | 오류 메시지 / 동작 | 근거 |
|------|-----------------|------|
| Outlook 미실행 | COM 예외 → 오류 다이얼로그 표시 | `src/app.py:584–586` |
| FC 파일 미설정 | 상태바 `"FC 엑셀 파일 경로를 설정하세요"` | `src/app.py:112` |
| FC 파일 경로 없음 | `"파일이 없습니다:\n{경로}"` 오류 다이얼로그 | `src/app.py:487–489` |
| Excel 파일 열려 있음 | `"파일이 열려 있습니다. Excel을 닫고 다시 시도하세요."` | `src/app.py:669–671` |
| 날짜 형식 오류 | `"날짜 형식 오류: '{입력값}'\nYYYY-MM-DD 형식으로 입력하세요."` | `src/app.py:500–502` |
| v1.1 미패치 상태에서 Excel 생성 | `'Worksheet' object has no attribute 'workbook'` — v1.1 `excel_writer.py:174` 버그 | 패치 적용으로 해결 |
