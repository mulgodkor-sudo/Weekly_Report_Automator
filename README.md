# 📊 Weekly Report Automator  `Ver.1.2`

> DL이앤씨 플랜트본부 기계설계팀 — 아웃룩 캘린더 기반 주간보고서 자동 생성 도구

---

## 📁 프로젝트 구조

```
Weekly_Report_Automator/
├── main.py                  ← 런처 (PyInstaller 진입점)
├── build_msi.bat            ← ★ 한 번에 onedir 폴더 + MSI 설치파일 빌드
├── build.bat                ← onedir 폴더만 빌드 (build_msi.bat onedir 호출)
├── installer/
│   ├── msi_tool.py          ← 버전 읽기 / MSI 버전 변환 / 설치 안내문 생성
│   ├── msi_installed.txt    ← MSI 설치 표시 파일 (패치가 이름 변경을 건너뜀)
│   └── wix/                 ← WiX v5 설치 정의 (Product.wxs, Variables.wxi, ui.ko-KR.wxl, 브랜딩 이미지)
├── patches/
│   ├── apply_patch_v1.11.bat ← v1.1 → v1.11 패치 (exe 옆 src/ 교체)
│   ├── apply_patch_v1.2.bat  ← v1.1 / v1.11 → v1.2 패치
│   ├── make_patch.py         ← 패치 .bat 생성기
│   └── README.md             ← 패치 안내
└── src/
    ├── app.py               ← 메인 UI (tkinter)
    ├── config.py            ← 설정 관리 + FC 데이터 로더
    ├── fc_rules.py          ← ★ FC 구분 분류 규칙 (패치 주요 대상)
    ├── event_processor.py   ← 아웃룩 이벤트 처리 + 경고 생성
    ├── excel_writer.py      ← xlsxwriter 기반 Excel 출력
    ├── outlook_reader.py    ← Outlook COM 연동
    ├── overrides.py         ← 되풀이 모임 override 관리
    ├── override_dialog.py   ← 되풀이 모임 설정 UI
    ├── plant_mh.py          ← Plant M/H 집계 + Excel 시트 (공통)
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

소스 폴더(이 README가 있는 폴더)에서 **`build_msi.bat` 더블클릭** 한 번이면 끝.

```cmd
build_msi.bat          ← onedir 폴더 + MSI 설치파일
build.bat              ← onedir 폴더만 (= build_msi.bat onedir)

# 결과물
dist/Weekly_Report_Automator_V1.2/          ← onedir 프로그램 폴더 (압축해서 배포 가능)
├── Weekly_Report_Automator_V1.2.exe
├── _internal/      ← DLL
└── src/            ← .py 파일 (패치 가능)
dist/Weekly_Report_Automator_V1.2.msi       ← 설치파일
```

- **버전은 `src/config.py`의 `APP_VERSION` 하나만 바꾸면** 폴더명·exe명·MSI 버전에 모두 반영된다.
- 필요한 것: Python 3 (PATH 등록). `pyinstaller / pywin32 / xlsxwriter / openpyxl`이 없으면 자동 설치.
- MSI 도구: **WiX Toolset v5.0.2** (무료, MS-RL. v6 이상은 유료 OSMF 라이선스라 사용 안 함).
  `.NET SDK`가 있으면 `dotnet tool`로 자동 설치/버전 고정 (MHS_Searcher와 같은 방식).
  .NET SDK가 없으면: `winget install Microsoft.DotNet.SDK.8` 또는 https://dotnet.microsoft.com/download
- 예전 빌드 오류 재발 방지:
  - bat 파일은 영문(ASCII)만 사용 → 한글 코드페이지 오류 없음
  - 매번 `dist/…`, `build/`, `.spec` 삭제 + PyInstaller `--clean` → 이전 빌드의 `src/`가 새 코드를 가리는 문제 없음
  - 프로젝트 폴더에 도구를 내려받지 않음 (바탕화면 "제어된 폴더 액세스"로 인한 `액세스가 거부되었습니다` 방지)
  - exe가 백신에 훼손됐는지 크기 검사, 바탕화면/OneDrive/한글 경로면 경고 표시
  - WiX 확장 버전 명시(`/5.0.2`) → 여러 버전이 섞여 생기던 WIX0094 방지

### MSI 설치파일 (MHS_Searcher 설치 구성과 동일)

- **한국어 설치 마법사**: 시작 → 안내문 → 설치 경로 선택 → 진행 → 완료(+ 프로그램 실행 체크박스), DL E&C 브랜딩 이미지
- 기본 설치 위치: `C:\Program Files\Autotools_Mechanical\Weekly Report Automator\` (관리자 권한으로 설치)
- 시작 메뉴 / 바탕화면 바로가기 `Weekly Report Automator 1.2`, "앱 및 기능"에 버전·문의처 표시
- 새 버전 MSI를 설치하면 이전 버전은 자동 제거 후 설치. 같은 버전 MSI 재실행 시 [업데이트]/[프로그램 삭제] 선택
- 설치 폴더의 `src\`는 일반 사용자도 수정 가능 → **패치(`apply_patch_vX.X.bat`)를 설치 폴더에 복사해 관리자 권한 없이 적용**.
  바로가기가 깨지지 않도록 MSI 설치본에서는 exe·폴더 이름을 바꾸지 않는다.
- 설정(`Documents\WeeklyReportAutomaker_*.json`)은 설치/제거와 무관하게 유지
- 설치 정의 파일: `installer/wix/Product.wxs`, `Variables.wxi`(제품명·제조사·UpgradeCode), `ui.ko-KR.wxl`(한국어 문구)

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
   - src/ 파일을 즉시 덮어쓴다 (백업 없음, 임시파일은 종료 시 자동 삭제)
   - .exe 파일명과 그 exe가 들어있는 폴더명도 새 버전 표기로 자동 변경
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
