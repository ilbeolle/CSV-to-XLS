# Excel/CSV to XLS Converter

## 설명

이 프로그램은 최신 Excel 파일(`.xlsx`, `.xls`) 또는 CSV 파일(`.csv`)을 구형 Excel 97-2003 형식(`.xls`)으로 변환합니다. 다운로드 받은 Excel 파일(예: 네이버 스마트 스토어, ESM Plus, 현대 이지웰 복지몰)을 사내 시스템에 업로드를 위해 최적화된 형식으로 변환하며, 비밀번호(`1234`)가 설정된 파일을 자동으로 처리하고 변환 후 비밀번호를 제거합니다.

### 지원 플랫폼

- **네이버 스마트 스토어**: 주문 정보 Excel 파일(비밀번호: `1234`) 변환 가능.
- **ESM Plus (G마켓, 옥션)**: 발송 관리 Excel 파일 변환 가능 (신규 주문 관리 아님).
- **현대 이지웰 복지몰**: CSV 파일 변환 가능 (고객 필터링 후 다운로드 필요).
- **11번가, 쿠팡**: 변환 가능 여부 확인 중.

### 주요 기능

- **Drag-n-Drop 지원**: 파일을 프로그램 아이콘 위로 끌어다 놓아 변환.
- **대화형 모드**: CMD 창에 파일 경로를 끌어다 놓고 Enter로 변환.
- **비밀번호 처리**: `1234` 비밀번호를 자동으로 풀고 변환 후 제거.
- **자동 저장**: 변환된 파일은 로컬 바탕화면(`C:\Users\[사용자]\Desktop`)에 저장.
- **오류 처리**: 파일 잠금, 손상, 또는 비밀번호 오류 시 사용자 친화적 메시지 출력.

## 사용 방법

### 1. 자동 모드 (Drag-n-Drop)

1. 변환할 `.xlsx`, `.xls`, `.csv` 파일을 선택.
2. 파일을 `ExcelConverter.exe` 아이콘 위로 끌어다 놓음.
3. 변환된 `.xls` 파일이 바탕화면에 저장됨.
4. 작업 완료 후 Enter 키로 종료.

### 2. 대화형 모드 (CMD)

1. CMD에서 `dist\ExcelConverter.exe` 실행.
2. 변환할 파일을 CMD 창에 끌어다 놓고 Enter 입력.
3. 변환된 `.xls` 파일이 바탕화면에 저장됨.
4. `exit` 또는 `quit` 입력으로 종료.

### 추가 명령어

- `--help`: 도움말 표시.
- `clear`: 콘솔 화면 지우기.
- `exit`/`quit`: 프로그램 종료.

## 문제점 및 해결책

### 1. Windows Defender 차단

- **문제**: Defender가 `.exe`를 의심스러운 파일로 인식해 차단하거나 삭제.
- **해결**:
  - Windows 보안 → 바이러스 및 위협 방지 → 허용된 항목 → `ExcelConverter.exe` 추가.
  - Microsoft에 false positive 보고: Windows Defender 제출.
  - `.exe`를 로컬 폴더(예: `C:\Users\[사용자]\Desktop`)로 이동해 실행.
  - 코드 서명(옵션): `signtool` 사용(별도 인증서 필요).

### 2. OneDrive 동기화 문제

- **문제**: OneDrive 폴더(`C:\Users\[사용자]\OneDrive\문서\projects`)에서 파일 잠금 발생.
- **해결**:
  - 프로젝트 파일(`converter.py`, `icon.ico`)과 `.exe`를 로컬 폴더(예: `C:\Users\[사용자]\Desktop`)로 이동.
  - OneDrive 동기화 비활성화 후 재시도.

### 3. Microsoft Excel 미설치

- **문제**: Excel이 설치되지 않았거나 라이선스가 없으면 작동 불가.
- **해결**: Microsoft Excel 설치 및 활성화 확인.

### 4. 실행 파일 작동 실패

- **문제**: `.exe` 실행 시 콘솔 출력 없음 또는 Drag-n-Drop 실패.
- **해결**:
  - CMD에서 실행: `dist\ExcelConverter.exe test.xlsx`로 테스트.
  - 오류 로그 확인: CMD에 표시된 오류 메시지 확인.
  - 백그라운드 Excel 프로세스 종료: 작업 관리자에서 `EXCEL.EXE` 종료.
  - PyInstaller 재빌드: 위 명령어로 캐시 정리 후 재생성.
 
# 똑같은 프로그램을 만드는 방법

- 이 프로그램을 분실했을 때, 똑같이 만들기 위한 메모임.

## 설치 및 설정

### 요구사항

- **운영체제**: Windows (Microsoft Excel 설치 필수).
- **소프트웨어**:
  - Python 3.8 이상.
  - Visual Studio Code (권장, 다른 IDE도 가능).
  - Microsoft Excel (라이선스 필요, 설치되지 않은 환경에서는 작동 불가).
- **의존성**:
  - `pyinstaller`: 실행 파일 생성.
  - `colorama`: 콘솔 색상 출력.
  - `pywin32`: Excel COM 객체 처리.

### 설치 단계

1. **Python 설치**:

   - Python 공식 사이트에서 Python 3.8 이상을 다운로드 및 설치.
   - 설치 시 "Add Python to PATH" 옵션을 체크하세요.
   - 설치 확인:

     ```bash
     python --version
     ```

     출력 예: `Python 3.10.0`

2. **Visual Studio Code 설치**:

   - VS Code 공식 사이트에서 다운로드 및 설치.
   - Python 확장 설치:
     - VS Code 실행 → 확장(Extensions) → `Python` 검색 → Microsoft 제공 Python 확장 설치.
   - Python 인터프리터 설정:
     - VS Code에서 `Ctrl+Shift+P` → `Python: Select Interpreter` → 설치된 Python 경로 선택.

3. **의존성 설치**:

   - CMD 또는 VS Code 터미널에서 다음 명령어 실행:

     ```bash
     pip install pyinstaller colorama pywin32
     ```
   - 설치 확인:

     ```bash
     pip show pyinstaller
     pip show colorama
     pip show pywin32
     ```

4. **프로젝트 설정**:

   - 프로젝트 폴더: `C:\Users\[사용자]\OneDrive\문서\projects`
   - 필수 파일:
     - `converter.py`: 메인 프로그램.
     - `icon.ico`: 실행 파일 아이콘 (프로젝트 폴더에 위치).

## 실행 파일(.exe) 생성

프로그램을 독립 실행 가능한 `.exe` 파일로 변환하려면 PyInstaller를 사용합니다.

### 명령어

```bash
cd C:\Users\[사용자]\OneDrive\문서\projects
pyinstaller --onefile --console --noupx --clean --icon=icon.ico --name=ExcelConverter converter.py
```

### 옵션 설명

- `--onefile`: 단일 `.exe` 파일로 압축.
- `--console`: 콘솔 창을 유지하여 Drag-n-Drop 및 대화형 모드 지원.
- `--noupx`: UPX 압축 비활성화로 Windows Defender 경고 최소화.
- `--clean`: 이전 빌드 캐시 삭제로 안정성 확보.
- `--icon=icon.ico`: 실행 파일에 사용자 지정 아이콘 적용.
- `--name=ExcelConverter`: 출력 파일명을 `ExcelConverter.exe`로 지정.

### 출력

- 생성된 파일: `C:\Users\[사용자]\OneDrive\문서\projects\dist\ExcelConverter.exe`
- 실행: `dist\ExcelConverter.exe`를 더블클릭하거나 CMD에서 실행.

## 기타

- **아이콘 설정**: `icon.ico`는 프로젝트 폴더에 위치해야 하며, 유효한 .ico 형식이어야 함. 무료 아이콘은 Icons8 또는 Flaticon에서 다운로드 가능.
- **디버깅**: 문제가 지속되면 다음 명령어로 로그 확인:

  ```bash
  pyinstaller --onefile --console --noupx --clean --icon=icon.ico --name=ExcelConverter --log-level DEBUG converter.py
  ```

## 문의

- 작성자: DongHyun LEE
- 연락처: ~~-~~
- 이슈: GitHub Issues에 문제 제보 또는 개선 제안.

### 가까운 LLM AI에게 문의하는 방법

- LLM AI의 기본 설정에서 먼저 입력해야 할 기본 프롬프트

  ```기본으로 저장할 프롬프트
  You are an expert-tier software architect operating in "Pair Programmer Mode." Your primary directive is to generate code guided by six core principles, which you must prioritize in all solutions.

  These principles are:
  1.  Security: All code must be secure by default. Aggressively sanitize all inputs, prevent common vulnerabilities (e.G., XSS, SQL injection, buffer overflows), use parameterized queries, and adhere to the principle of least privilege. User safety is paramount.
  2.  Performance: Code must be highly efficient. Prioritize optimal algorithmic complexity (e.g., O(n log n) over O(n^2)), minimize resource consumption (CPU, memory), and avoid unnecessary computations.
  3.  Stability (Robustness): Code must be resilient. Implement comprehensive error handling, graceful failures, and ensure predictable behavior even with edge cases.
  4.  Maintainability: Code must be exceptionally clean, readable, and self-documenting. Use clear and descriptive variable/function names. The logic should be straightforward. You must NOT use comments, unless explaining highly complex, non-obvious algorithms.
  5.  Accessibility (A11y): For any user-facing code (HTML, CSS, JavaScript), ensure strict compliance with modern accessibility standards (e.g., WCAG 2.1 AA level), including semantic HTML and ARIA roles.
  6.  Modern Workflow: Adhere to current best practices, idiomatic patterns of the specific language, and produce modular, scalable code that is version-control friendly.
  
  Your response format must always be:
  1.  Approach: A brief, high-level explanation of your proposed solution.
  2.  Code: The clean, well-structured code block.
  3.  Next Steps: Critical considerations or follow-up actions.

  If the user's request is ambiguous, ask one clarifying question before providing code.
  ```

- LLM AI와 새로운 대화를 시작하며 입력할 프롬프트

  ```대화로 시작할 프롬프트
  안녕하세요, 당신은 40년 경력의 전문 프로그래머입니다. 어떤 프로그램이건 만들지 않은 프로그램이 없고, 어떤 컴퓨터 언어이건 이해하지 못하는 프로그램 언어가 없습니다. VBA부터 C, C++, 그리고 Python까지 어떤 언어가 어떻게 어떤 환경에 적합하게 쓰여졌는지 이해하고 있습니다. 그리고 제대로 이해하고 있는 만큼 수많은 프로그래머를 배출한 대학원의 교수님이시기도 하지요.
  
  저는 당신에게 어떻게 코드를 작성하고 쓰는지 이해하고 싶으며 프로그램을 만들고 싶습니다. 
  사용할 컴퓨터 언어는 Python, 사용할 프로그램은 Visual Studio Code 입니다.
  
  또한, 저는 한국어로 작성된 답변을 받아보고 싶습니다.
  
  당신은 제 두서없는 질문에도 충실하게, 상세하고 자세하게, 정확하고 명확하게 그 근거와 이유를 들어 답변하시기 바랍니다.
  ```
