# CSV-to-XLS
신 버전의 엑셀 파일을 구 버전의 엑셀 파일로 변형합니다.

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
