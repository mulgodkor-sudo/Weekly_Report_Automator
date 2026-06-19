# 패치 안내 — apply_patch_v1.11.bat

> 기준선(baseline): **v1.1** (최초 README 작성 시점)
> 패치 후 버전: **v1.11**

---

## 사용 방법

1. `Weekly_Report_Automator_VX.X.exe` 가 있는 폴더(= `src/` 폴더가 있는 폴더)에
   `apply_patch_v1.11.bat` 파일을 복사한다.
2. `apply_patch_v1.11.bat` 더블클릭 실행.
   - `src/` 폴더가 없으면 즉시 에러로 중단 (잘못된 위치에서 실행한 경우).
   - 패치 데이터는 base64 → zip으로 디코딩 후 `src/` 폴더에 **전체 파일 교체** 방식으로 적용된다 (부분 수정 없음).
   - 적용 중 생성되는 임시 파일(`%TEMP%\wra_patch_*.b64`, `*.zip`)은 성공/실패와 관계없이
     스크립트 종료 시 자동 삭제된다. **백업 폴더는 생성하지 않는다.**
3. 패치 완료 후 프로그램을 재시작하면 적용된다.

> ⚠️ `main.py`는 PyInstaller가 `.exe`에 직접 빌드해 넣는 진입점이라 이 패치로
> 갱신할 수 없다. main.py가 바뀐 경우는 `build.bat`로 다시 빌드해서 `.exe` 자체를
> 교체해야 한다. (이번 v1.11 패치에는 main.py 변경 내용이 없으므로 해당 없음)

---

## 패치 대상 파일 (4개, 전체 교체)

| 파일 | 변경 내용 |
|------|-----------|
| `src/app.py` | 화면 표시 버전을 `config.json`의 `version` 값을 따르도록 변경 (하드코딩 제거), "DL이엔씨" → "DL이앤씨" 오타 수정 |
| `src/config.py` | `version` 키 추가 + `get_version_str()` 헬퍼 추가 (버전 정보를 config.json 기준으로 관리) |
| `src/config.json` | `version: "1.11"` 추가 |
| `src/excel_writer.py` | 날짜 정렬 버그 수정(가장 빠른 날짜를 1행에 고정 후 구분별 정렬), I열 병합 회귀 버그 및 테두리 누락 버그 수정 |

---

## 패치에 포함되지 않는 변경

- `main.py` (스플래시 화면 버전 표시 로직) — 위 경고 참고, 이번 버전에는 변경 없음
- `build.bat`, `README.md` 등 빌드/문서 파일 — exe 동작에 영향 없음

---

## 검증 상태

- base64 → zip 디코딩 결과는 원본 `src/*.py`, `src/*.json` 파일과 바이트 단위로 동일함을
  개발 환경에서 직접 확인했다.
- 실제 팀원 PC(Windows + PowerShell)에서 1회 실행 테스트 완료, 정상 동작 확인.
