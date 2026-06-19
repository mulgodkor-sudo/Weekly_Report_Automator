# 📊 Weekly Report Automator  `Ver.1.11`

> DL이앤씨 플랜트본부 기계설계팀 — 아웃룩 캘린더 기반 주간보고서 자동 생성 도구

---

## 📁 프로젝트 구조

```
Weekly_Report_Automator/
├── main.py                  ← 런처 (PyInstaller 진입점)
├── build.bat                ← PyInstaller 빌드 스크립트
├── patches/
│   └── apply_patch_v1.11.bat ← v1.1 → v1.11 패치 (exe 옆 src/ 교체)
└── src/
    ├── app.py               ← 메인 UI (tkinter)
    ├── config.py            ← 설정 관리 + FC 데이터 로더
    ├── fc_rules.py          ← ★ FC 구분 분류 규칙 (패치 주요 대상)
    ├── event_processor.py   ← 아웃룩 이벤트 처리 + 경고 생성
    ├── excel_writer.py      ← xlsxwriter 기반 Excel 출력
    ├── outlook_reader.py    ← Outlook COM 연동
    ├── overrides.py         ← 되풀이 모임 override 관리
    ├── override_dialog.py   ← 되풀이 모임 설정 UI
    ├── plant_mh_dialog.py   ← Plant M/H 입력 확인 UI
    ├── monthly_processor.py ← 월간업무정리 처리 로직
    ├── monthly_dialog.py    ← 월간업무정리 UI + Excel 저장
    ├── config.json          ← 초기 설정값
    ├── overrides.json       ← 초기 override 데이터
    └── assets/
        ├── splash.png
        └── Schedule_Ico.ico
```

---

## ⚙️ 주요 기능

| 기능 | 설명 |
|------|------|
| **주별 추출** | 이번주 실적 + 다음주 계획을 아웃룩에서 자동 추출 |
| **일별 추출** | 특정 날짜 하루치 실적 추출 |
| **Excel 생성** | Weekly Report 양식 엑셀 자동 생성 (I열 병합, 별표 붉은색) |
| **사전 검토** | 시간 합계, FC 누락, 코드 오류 자동 점검 |
| **Plant M/H** | 날짜별 업무 시간 확인 |
| **월간업무정리** | 전월 실적 월 단위 집계 + Excel 저장 |
| **되풀이 모임** | 반복 회의 상세내용 고정 설정 |

---

## 🔧 구분 분류 로직 (`fc_rules.py`)

```
Project  + KPI=O                  → 수행 (KPI)
General  + KPI=O + PC=000000      → 기타 (KPI)
General  + KPI=O + PC≠000000      → 수행 (KPI)  ← 외부 프로젝트 투입
KPI=X                             → 기타
conditional (GA08-01 등)          → 조건부 판단
```

> **패치 시** `fc_rules.py` 하나만 수정·배포하면 됩니다.

---

## 🏗️ 빌드 방법

```cmd
# 1. 소스 폴더에서 실행
build.bat

# 결과물
dist/Weekly_Report_Automator_V1.11/
├── Weekly_Report_Automator_V1.11.exe
├── _internal/      ← DLL
└── src/            ← .py 파일 (패치 가능)
```

---

## 🩹 패치 배포 방법

`patches/apply_patch_vX.X.bat` 파일 하나로 배포한다. 내부에 변경된 `src/*.py`,
`src/*.json` 전체 파일이 base64로 인코딩되어 들어 있고, 실행하면 zip으로 복원한
뒤 `src/` 폴더 파일을 **통째로 덮어쓴다** (부분 수정이 아니라 전체 교체이므로,
팀원마다 src/ 내용이 조금씩 달라도 항상 같은 결과로 맞춰진다).

```cmd
# 팀원 — 패치 적용
1. Weekly_Report_Automator_VX.X.exe 가 있는 폴더(= src/ 폴더가 있는 폴더)에
   apply_patch_vX.X.bat 복사
2. apply_patch_vX.X.bat 더블클릭
   - 패치 전 src/ 폴더를 src_backup_v* 폴더로 자동 백업
   - 실패 시 자동으로 백업본 복원
3. 프로그램 재시작
```

> ⚠️ `main.py` 안의 코드(스플래시 화면 등 PyInstaller 진입점)는 exe에 직접
> 빌드되어 있어 패치로 갱신되지 않는다. main.py가 바뀐 경우는 `build.bat`로
> 다시 빌드해서 exe 자체를 교체해야 한다.

---

## 📋 아웃룩 일정 작성 규칙

```
[실적] 업무명 [#프로젝트코드#기능코드]
[계획] 업무명 [#프로젝트코드#기능코드]

예) [실적] MHS 표준서 작성 [#000000#GA07-20]
예) [계획] FAT 지원 [#P2600O#PF03-05]
```

- PC/FC 코드 내 공백 허용: `[# 000000 # GA11 - 25]` 도 자동 인식
- `*` 이후 내용은 Excel 출력 시 붉은색으로 표시

---

## 🛠️ 개발 환경

- Python 3.13
- Windows 10/11
- Microsoft Outlook (데스크탑, 로그인 필수)
- 주요 패키지: `xlsxwriter`, `pywin32`, `openpyxl`

---

## 📂 설정 파일 저장 위치

```
C:\Users\{사용자명}\Documents\WeeklyReportAutomaker_config.json
C:\Users\{사용자명}\Documents\WeeklyReportAutomaker_overrides.json
```

> 프로그램 재설치·패치 후에도 설정이 유지됩니다.

---

*DL이앤씨 플랜트본부 기계설계팀 — 이수신 차장*
