# 엑셀의 수치를 매번 한글 보고서 파일로 옮겨야 하는 피로 해소용 컨버터

단순 반복 작업에서 발생하는 휴먼 에러를 원천 차단하기 위해 제작한 업무 자동화 도구입니다.<br>
많은 양의 엑셀(.xlsx) 데이터를 한 번의 실행으로 한글(.hwpx) 보고서 양식에 자동 기입합니다.

<br><br>

### ✨ Features

- **HWP 필드 매칭** — 엑셀에 작성된 수치(카테고리명과 연결하거나 특정 row,column의 셀과 연결)와 한글 문서의 '누름틀'을 1:1 매칭하여 정확한 위치에 입력
- **출력물 스타일링** — 정해진 양식에 따라 천 단위 콤마, 단일행/다중행, 취소선, 글자 색상, 증감표시, 굵기 등을 자동 적용
  - 단일 수치: 요구액과 확정액이 동일할 경우 한 줄로 표기
  - 수정 수치: 요구액(취소선) + 확정액(파란색)으로 두 줄 병기
  - 음수 기호: 음수 수치는 △ 기호를 부착하고, 빨간색으로 표시
- **증감액과 증감률 계산** — 전년 대비 증감액 및 증감률을 계산하여 기입
- **진행률 표시** — 터미널 내 프로그레스 바를 통해 진행 상태 시각화
- **디버깅 로그 생성** — 오류 발생 시 debug_log.txt로 추출하여 유지보수 편의성 확보

<br><br>

### 🏃 Getting Started

⚠️ Windows 전용이며, 한글 프로그램 설치와 실행 권한이 필요합니다.

1. 핵심 파일은 `양식파일.hwpx`, `*.xlsx`, `convert.py` 이렇게 3가지 파일입니다.

```
.
├── constants.py          # 필드 매핑
├── convert.py            # 실행 스크립트
├── example.xlsx          # 수치가 들어있는 엑셀 파일
└── 양식파일.hwpx          # 한글 보고서 템플릿
```

<br>

2. constants.py에 데이터의 위치를 기입합니다.

```
# column
COL_NOW = 9    # column J (올해 예산)
COL_REQ = 13   # column N (내년 요구)
COL_FIX = 17   # column R (최종 검토)
```

엑셀 파일 내 카테고리명과 한글 문서 내 '누름틀'을 연결합니다.

```
PROGRAMS = {
    ("[1000]CATEGORY_NAME"): "1000_TOTAL",
    ...
}
```

<br><br>

3. 필요한 라이브러리 설치 후 스크립트를 실행합니다.

```
pip install pandas pyhwpx openpyxl
python convert.py
```

파이썬 설치 없이 EXE 파일로 실행할 수도 있습니다.
EXE 파일은 아래 명령어로 생성합니다.

```
pyinstaller --onefile convert.py
```

이후 생성된 dist 폴더에서 convert.exe를 꺼내 convert.py와 동일한 위치에 두고 실행하세요.

<br><br>

### 🔨 Tech Stack

- Language: `Python`
- Library: `pandas`, `pyhwpx`, `tkinter`
- Stability: `logging`, `traceback`
- Build Tool: `PyInstaller`

<br><br>

### 🔑 Security & Privacy

- 본 도구는 로컬 환경에서 구동되는 독립형 스크립트로, 외부 서버와 통신하지 않습니다.
