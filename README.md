# CSV-to-XLS
설명
---
신 버전의 엑셀 파일을 구 버전의 엑셀 파일로 변형합니다.

사내 시스템에 업로드하기 위한 알맞은 엑셀 형식으로 바꾸기 위해 사용합니다.

네이버 스마트 스토어의 주문 정보 엑셀 파일., 비밀번호는 항상 '**1234**'이어야 합니다. (대응 가능)

ESM Plus (Gmarket, Auction)의 발송관리 엑셀 파일., 신규주문 관리가 아닙니다. (대응 가능)

현대 이지웰의 복지몰의 CSV 파일. 이때, 원하는 고객만 추려서 다운로드할 필요가 있습니다. (대응 가능)

11번가 엑셀 파일 확인 중... (대응 가능 여부 모름)


수정 예정
---
v1.4.0 보안 등급 수정 중...


만드는 방법...
---
해당 코드의 컴파일 코드는 아래와 같습니다.


```
pyinstaller --onefile --console --icon="icon.ico" --name="Excel_to_XLS_Converter" converter.py
```

[projects] 폴더 내의 [dist] 내의 실행 파일로 작업을 시작합니다...

[projects] 폴더 바깥의 파일은 모두 구버전입니다...


Python 및 Visual Studio Code 설치 방법 설명 추가 예정...

Visual Studio Code 내 Python 연동을 위한 Pyinstaller 설치 방법 추가 예정...

의존성 설치 방법 추가 예정...

아이콘(icon.ico) 설정 방법 추가 예정...


```
python --version
pip install pyinstaller
pip show pyinstaller
pip install colorama pywin32
```


기타...

```
pyinstaller --onefile --console --icon="icon.ico" --name="Excel_to_XLS_Converter" --log-level DEBUG converter.py
```
