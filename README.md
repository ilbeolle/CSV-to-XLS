# Excel/CSV to XLS Converter v2.1.0

## 설명

이 프로그램은 최신 Excel 파일(.xlsx, .xls) 또는 CSV 파일(.csv)을 구형 Excel 97-2003 형식(.xls)으로 변환합니다. 다운로드 받은 Excel 파일(예: 네이버 스마트 스토어, ESM Plus, 현대 이지웰 복지몰)을 사내 시스템에 업로드를 위해 최적화된 형식으로 변환하며, 비밀번호(1234)가 설정된 파일을 자동으로 처리하고 변환 후 비밀번호를 제거합니다. 

### 지원 플랫폼

- **네이버 스마트 스토어**: 선택주문발주발송관리 Excel 파일(비밀번호 `1234` 자동 해제) 변환 가능.
- **ESM Plus (G마켓, 옥션)**: 발송 관리 Excel 파일 변환 가능.
- **현대 이지웰 복지몰**: EUC-KR로 인코딩된(?) CSV 파일 변환 가능 (고객 필터링 후 다운로드 필요).
- **11번가**: 발주확인을 하기 전의 엑셀 파일로 업로드할 것.
- **쿠팡**: 변환 가능 여부 확인 중 (추가 테스트 필요).

### 주요 기능

- **Drag-n-Drop 지원**: 파일을 프로그램 아이콘 위로 끌어다 놓아 변환.
- **대화형 모드**: CMD 창에 파일 경로를 끌어다 놓고 Enter로 변환.
- **비밀번호 처리**: 1234 비밀번호를 자동으로 풀고 변환 후 제거.
- **자동 저장**: 변환된 파일은 로컬 바탕화면(C:\Users\[사용자]\Desktop)에 저장 (OneDrive 우회).
- **오류 처리**: 파일 잠금, 손상, 또는 비밀번호 오류 시 사용자 친화적 메시지 출력 (추가: 시스템 권한, 메모리 부족 등 세부 안내).
- **자동 종료**: 성공 시 3초 카운트다운 후 창 닫힘.
- **충돌 방지**: 저장하려는 파일명이 이미 존재할 경우, 덮어쓰지 않고 `파일명_1.xls`, `파일명_2.xls`와 같이 숫자를 붙여 안전하게 저장합니다.

## 사용 방법

### 1. 자동 모드 (Drag-n-Drop)

1. 변환할 .xlsx, .xls, .csv 파일(들)을 선택.
2. 파일을 ExcelConverter.exe 아이콘 위로 끌어다 놓음.
3. 변환된 .xls 파일이 바탕화면에 저장됨.
4. 성공 시 3초 후 자동 종료; 실패 시 Enter 키로 종료.

### 2. 대화형 모드 (CMD)

1. CMD에서 dist\ExcelConverter.exe 실행.
2. 변환할 파일을 CMD 창에 끌어다 놓고 Enter 입력.
3. 변환된 .xls 파일이 바탕화면에 저장됨 (바탕화면 자동 열림).
4. exit 또는 quit 입력으로 종료.

### 추가 명령어

- `--help`: 도움말 표시.
- `clear`: 콘솔 화면 지우기.
- `exit` 또는 `quit`: 프로그램 종료.

## 주의점
- 변환 작업 이후 엑셀을 강제 종료하는 기능이 내장되어 있기 때문에, 반드시 작업 중인 엑셀 파일은 안전하게 종료할 것.

## 문제점 및 해결책

### 1. Windows Defender 차단

- **문제**: Defender가 .exe를 의심스러운 파일로 인식해 차단하거나 삭제.
- **해결**:
  - Windows 보안 → 바이러스 및 위협 방지 → 허용된 항목 → ExcelConverter.exe 추가.
  - Microsoft에 false positive 보고: Windows Defender 제출.
  - .exe를 로컬 폴더(예: C:\projects\ExcelConverter)로 이동해 실행.
  - 코드 서명(옵션): signtool 사용(별도 인증서 필요).

### 2. OneDrive 동기화 문제

- **문제**: OneDrive 폴더에서 파일 잠금 발생.
- **해결**:
  - 프로젝트 파일(converter.py, icon.ico)과 .exe를 로컬 폴더(예: C:\projects\ExcelConverter)로 이동.
  - OneDrive 동기화 비활성화 후 재시도.

### 3. Microsoft Excel 미설치

- **문제**: 이 프로그램은 Microsoft Excel의 내부 엔진(COM Interface)을 빌려 사용합니다. 따라서 Excel이 설치되지 않았거나 라이선스가 없으면 작동되지 않습니다.
- **해결**: 반드시 PC에 정품 Microsoft Excel이 설치되어 있어야 하며, Hancom Office 한셀 등은 호환되지 않습니다.

### 4. 실행 파일 작동 실패

- **문제**: .exe 실행 시 콘솔 출력 없음 또는 Drag-n-Drop 실패.
- **해결**:
  - CMD에서 실행: dist\ExcelConverter.exe "test.xlsx"로 테스트.
  - 오류 로그 확인: CMD에 표시된 오류 메시지 (추가: COM 초기화 실패 시 Pythoncom 확인).
  - 백그라운드 Excel 프로세스 종료: 작업 관리자에서 EXCEL.EXE 종료.
  - PyInstaller 재빌드: 아래 명령어로 캐시 정리 후 재생성.

### 5. 확인된 오류

- **문제**: 
  - 손상된 Excel로 열 수 없는 .xlsx 파일로 저장된 경우
  - < > ? [ ] : | 등의 특수문자가 포함된 Excel 파일인 경우
  - 파일 이름 및 폴더 경로가 218개를 초과하는 문자를 포함한 경우
  - 프로그램은 실행되나 변환 작업이 진행되지 않는 경우
- **해결**:
  - Excel로 열 수 있는 올바른 파일(문서)인지 확인할 것.
  - Windows 및 이 프로그램에서 파일 이름으로 사용 가능한 이름인지 확인하고 수정할 것.
  - 파일 이름을 짧은 이름으로 줄이거나 C 드라이브 등 찾기 쉬운 상위 폴더로 이동할 것.
  - 본 프로그램과 엑셀 프로세스 자체를 완전히 강제 종료한 후 작업을 다시 실행할 것.
  - 프로그램 실행 중에 프로그램 상에서,  Enter 키를 눌러줄 것

# 똑같은 프로그램을 만드는 방법

- 이 프로그램을 분실했을 때, 똑같이 만들기 위한 메모이자 그 구성 순서
- 또는 소스 코드(`converter.py`)를 기반으로 실행 파일(`.exe`)을 직접 생성(Build)하는 상세 가이드

## 설치 및 설정

### 요구사항

- **운영체제**: Windows 10/11 (x64) (Microsoft Excel 설치 필수).
- **소프트웨어**:
  - Python 3.8 이상.
  - Visual Studio Code (권장, 다른 IDE도 가능).
  - Microsoft Excel (라이선스 필요, 설치되지 않은 환경에서는 작동 불가).
- **의존성**:
  - pyinstaller: 실행 파일 생성.
  - colorama: 콘솔 색상 출력.
  - pywin32: Excel COM 객체 처리.

### 설치 단계 (선행 작업: .EXE 생성 전 필수)

1. **Python 설치**:
   - Python 공식 사이트에서 Python 3.8 이상을 다운로드 및 설치.
   - 설치 시 "Add Python to PATH" 옵션을 체크하세요.
   - 설치 확인: CMD에서 python --version 실행 (출력 예: Python 3.12.3).

2. **Visual Studio Code 설치**:
   - VS Code 공식 사이트에서 다운로드 및 설치.
   - Python 확장 설치: VS Code 실행 → 확장(Extensions) → Python 검색 → Microsoft 제공 Python 확장 설치.
   - Python 인터프리터 설정: VS Code에서 Ctrl+Shift+P → Python: Select Interpreter → 설치된 Python 경로 선택.

3. **의존성 설치**:
   - CMD 또는 VS Code 터미널에서 다음 명령어 실행:
     `pip install pyinstaller colorama pywin32`
   - 설치 확인:
     `pip show pyinstaller`
     `pip show colorama`
     `pip show pywin32`

4. **프로젝트 설정**:
   - 프로젝트 폴더 생성: C:\projects\ExcelConverter (로컬 경로 사용, OneDrive 피함).
   - 필수 파일:
     - converter.py: 메인 프로그램 (v1.6.0 beta 코드 복사).
     - icon.ico: 실행 파일 아이콘 (해당 폴더와 같은 위치, 무료 사이트에서 다운로드. 필수 사항이 아닌 선택 사항). 

## 실행 파일(.exe) 생성

프로그램을 독립 실행 가능한 .exe 파일로 변환하려면 `PyInstaller`를 사용합니다.

### 명령어 (프로젝트 경로 수정 적용)

CMD에서 다음 실행:
  ```
  cd C:\projects\ExcelConverter
  pyinstaller --onefile --console --noupx --clean --icon=icon.ico --name=ExcelConverter converter.py
  ```
  다른 명령어로도 실행이 가능 `pyinstaller --noconfirm --onefile --console --clean --name "ExcelConverter" --icon "icon.ico" converter.py`

### 옵션 설명

- `--noconfirm`: 기존 `dist` 폴더 내 파일 덮어쓰기 확인 없이 진행.
- `--onefile`: 단일 `.exe` 파일로 압축.
- `--console`: 콘솔 창을 유지하여 Drag-n-Drop 및 대화형 모드 지원.
- `--noupx`: UPX 압축 비활성화로 Windows Defender 경고 최소화.
- `--clean`: 이전 빌드 캐시 삭제로 안정성 확보.
- `--icon=icon.ico`: 실행 파일에 사용자 지정 아이콘 적용.
- `--name=ExcelConverter`: 출력 파일명을 `ExcelConverter.exe`로 지정.

### 출력

명령어가 성공적으로 실행되면 프로젝트 폴더 내 `dist` 폴더가 생성됨.
- 생성된 파일: `C:\projects\ExcelConverter\dist\ExcelConverter.exe`
- 실행: `dist\ExcelConverter.exe`를 더블클릭하거나 CMD에서 실행.

## 기타

- **아이콘 설정**: icon.ico는 프로젝트 폴더에 위치해야 하며, 유효한 .ico 형식이어야 함. 무료 아이콘은 Icons8 또는 Flaticon에서 다운로드 가능.
- **디버깅**: 문제가 지속되면 다음 명령어로 로그 확인:
  ```
  pyinstaller --onefile --console --noupx --clean --icon=icon.ico --name=ExcelConverter --log-level DEBUG converter.py
  ```

## 문의

- 작성자: DongHyun LEE
- 연락처: ~~-~~
- 이슈: GitHub Issues에 문제 제보 또는 개선 제안.

### 가까운 LLM AI에게 문의하는 방법

 - LLM AI의 기본 설정에서 먼저 입력해야 할 프롬프트
 ```
You are an expert-tier software architect operating in "Pair Programmer Mode." Your primary directive is to generate code guided by the ISO/IEC 25002:2024 quality model, prioritizing the following eight core characteristics from the SQuaRE framework.

These characteristics are:
1. Functional Suitability: Ensure code fully realizes specified functions with complete, accurate, and appropriate functionality, avoiding over- or under-implementation.
2. Performance Efficiency: Optimize for time and resource behavior, prioritizing efficient algorithms (e.g., O(n log n) over O(n^2)), minimizing CPU/memory usage, and scaling under load.
3. Compatibility: Design for seamless interoperability with other systems, data formats, and environments, using standard protocols and avoiding vendor lock-in.
4. Usability: Prioritize user experience with intuitive, learnable, and accessible interfaces; for user-facing code (HTML, CSS, JavaScript), comply with WCAG 2.1 AA standards, including semantic HTML and ARIA roles.
5. Reliability: Build resilient code with fault tolerance, error recovery, and predictable behavior; implement comprehensive error handling, graceful degradation, and support for edge cases.
6. Security: Enforce security by default with input sanitization, prevention of vulnerabilities (e.g., XSS, SQL injection, buffer overflows), parameterized queries, and least privilege principles.
7. Maintainability: Produce modular, readable code with descriptive names, straightforward logic, and version-control-friendly structure; adhere to language idioms without unnecessary comments unless for complex algorithms.
8. Portability: Ensure adaptability across environments with platform-agnostic design, avoiding hard-coded dependencies, and supporting easy deployment and migration.

Your response format must always be:
1. Approach: A brief, high-level explanation of your proposed solution.
2. Code: The clean, well-structured code block.
3. Next Steps: Critical considerations or follow-up actions.

If the user's request is ambiguous, ask one clarifying question before providing code.
 ```

 - LLM AI와 새로운 대화를 시작하며 부여할 페르소나
 ```
[WHO: 역할 정의]
당신은 40년 경력의 전문 프로그래머이자, 수많은 프로그래머를 배출한 대학원 교수입니다. VBA, C, C++, Python 등 모든 컴퓨터 언어를 완벽히 이해하며, 각 언어의 환경 적합성(예: Python의 데이터 처리 효율성)을 깊이 알습니다. 특히 Python 전문가로서, Excel 파일 변환 프로젝트(ExcelConverter)에서 보안, 성능, 안정성을 최우선으로 고려합니다.

[WHAT: 작업 내용]
주요 작업은 사용자가 제공한 기존 코드(converter.py)를 기반으로 개선된 Python 프로그램을 개발하는 것입니다. Excel 파일을 다른 형식으로 변환하는 기능을 중심으로 코드 작성, 디버깅, 최적화, icon.ico 통합을 수행합니다. 출력은 단계별 설명, 코드 블록, 테스트 사례를 포함하며, 보안(입력 검증, SQL 인젝션 방지), 성능(O(n) 알고리즘 우선), 안정성(예외 처리)을 준수합니다.

[WHEN: 시기 및 지식 범위]
채팅 시작부터 지속적으로 적용되며, 지식은 2025년 10월 24일 현재까지의 최신 Python 베스트 프랙티스(예: Python 3.12+, pandas/openpyxl 라이브러리)를 기반으로 합니다. 실시간 질문에 즉시 응답하며, 미래 예측 시 명확히 밝히세요.

[WHERE: 환경 및 컨텍스트]
개발 환경은 Visual Studio Code(VS Code)이며, 프로젝트 경로는 C:\projects\ExcelConverter입니다. 주요 파일은 converter.py(메인 스크립트)와 icon.ico(아이콘 파일)입니다. Windows 환경을 가정하며, 필요한 라이브러리(예: pandas, openpyxl)는 pip로 설치 가능성을 고려합니다.

[WHY: 목적 및 목표]
사용자가 코드를 이해하고, Excel 변환 프로그램을 성공적으로 개발할 수 있도록 돕기 위함입니다. 이는 교육적 가치(근거 설명)를 통해 사용자의 프로그래밍 스킬을 향상시키고, 실무적 프로그램(ExcelConverter)을 완성하는 데 초점을 맞춥니다. 궁극적으로 안전하고 효율적인 소프트웨어를 만들기 위함입니다.

[HOW: 응답 방식 및 방법론]
- 모든 응답은 한국어로 작성하며, 사용자의 모국어(한국어)를 존중합니다.
- 두서없는 질문에도 충실히 답변: 질문을 재구성하여 확인 후, 상세·정확·명확하게 설명. 근거(예: PEP 8 스타일 가이드, 알고리즘 복잡도 분석)를 들어 이유 제시.
- 단계별 접근(Chain-of-Thought): 1) 문제 분석, 2) 솔루션 설계, 3) 코드 구현, 4) 테스트/최적화, 5) 잠재적 이슈 설명.
- 코드 형식: Markdown 코드 블록 사용, 주석 최소화(자명한 로직), 현대적 패턴(모듈화, 타입 힌트) 적용.
- 접근성: 사용자 친화적 설명, 에러 시 graceful failure.
- 이해 확인: "이해했습니다. 코드를 공유해 주세요."로 시작.

이 페르소나를 이해했다면, 사용자의 코드를 기다립니다.
 ```
