"""
Excel/CSV to XLS (Excel 97-2003) Converter
------------------------------------------------
이 프로그램은 .xlsx, .xls, .csv 파일을 구형 .xls 포맷으로 변환합니다.
Drag-n-Drop을 지원하며, '1234' 암호를 자동으로 처리하고 제거합니다.

- 작성자: DongHyun LEE
- 버전: 1.5.2
- 최종 수정일: 2025-10-22
"""

import sys
import os
import traceback
import win32com.client
from colorama import init, Fore, Style, Back
import time
import logging

# --- 상수 정의 (Constants) ---
__author__ = "DongHyun LEE"
__version__ = "1.5.2"
__last_modified__ = "2025-10-22"
__contact__ = "ilbeolle@gmail.com"

# Excel 파일 포맷 상수
XL_EXCEL8 = 56  # .xls (Excel 97-2003)
XL_FORCE_OVERWRITE = 2  # 덮어쓰기 강제 (xlLocalSessionChanges)
XL_UPDATE_LINKS_NEVER = 3  # 링크 업데이트 안함

# 하드코딩된 비밀번호 (사용자 지정 예외)
DEFAULT_PASSWORD = "1234"

# 색상 정의 (Color Palette)
init(autoreset=True)
C_TEXT = Fore.WHITE + Back.BLACK
C_TITLE = Style.BRIGHT + Fore.CYAN + Back.BLACK
C_AUTHOR = Style.BRIGHT + Fore.YELLOW + Back.BLACK
C_SUCCESS = Style.BRIGHT + Fore.GREEN + Back.BLACK
C_ERROR = Style.BRIGHT + Fore.RED + Back.BLACK
C_HELP = Fore.LIGHTBLACK_EX + Back.BLACK
C_CMD = Fore.MAGENTA + Back.BLACK
C_WARN = Fore.LIGHTYELLOW_EX + Back.BLACK

# 로깅 설정 (파일 생성 제거, 콘솔에만 출력)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

# --- 유틸리티 모듈 (Utilities Module) ---

def print_splash():
    """초기 스플래시 화면 및 작성자 정보를 출력합니다."""
    print(C_TITLE + "==========================================================")
    print(C_TITLE + f"  Excel/CSV to XLS Converter (v{__version__})")
    print(C_TITLE + "==========================================================")
    print(C_AUTHOR + f"  - 작성자: {__author__}")
    print(C_AUTHOR + f"  - 최종 수정일: {__last_modified__}")
    print(C_AUTHOR + f"  - 연락처: {__contact__}")
    print(C_TEXT)

def print_help():
    """사용 방법(Help)을 더 쉽게 인식할 수 있도록 단계별로 출력합니다."""
    print(f"{C_TITLE}\n[ 사용 가이드 ]")
    print(f"{C_HELP}----------------------------------------------------------")
    print(f"{C_TEXT}이 프로그램은 최신 Excel(.xlsx, .xls) 또는 CSV(.csv) 파일을")
    print(f"{C_TEXT}오래된 Excel 97-2003(.xls) 형식으로 바꿔줍니다.")
    print(f"{C_TEXT}암호가 '1234'인 파일도 자동으로 풀고, 변환 후 암호를 없앱니다.")
    print(f"{C_TEXT}네이버 스마트스토어, ESM Plus(G마켓, 옥션), 현대이지웰(복지몰)에 대응 가능합니다.")
    print(f"{C_TEXT}쿠팡, 11번가는 확인이 필요합니다.")
    
    print(f"{C_TITLE}\n[ 쉬운 방법: 자동 모드 ]")
    print(f"{C_TEXT}  1. 변환할 파일(들)을 마우스로 선택하세요.")
    print(f"{C_TEXT}  2. 선택한 파일을 이 프로그램(.exe) 아이콘 위로 끌어다 놓으세요.")
    print(f"{C_TEXT}  3. 자동으로 변환되어 바탕화면에 저장됩니다.")
    print(f"{C_TEXT}  4. 작업 끝나면 Enter 키를 눌러 창을 닫으세요.")

    print(f"{C_TITLE}\n[ 고급 방법: 대화형 모드 ]")
    print(f"{C_TEXT}  1. 프로그램(.exe)을 더블클릭해서 검은 창(CMD)을 열으세요.")
    print(f"{C_TEXT}  2. 변환할 파일을 이 검은 창으로 끌어다 놓으세요.")
    print(f"{C_TEXT}  3. 검은 창에서 Enter 키를 누르면 변환 시작합니다!")
    
    print(f"{C_TITLE}\n[ 변환 시 알아둘 점 ]")
    print(f"{C_HELP}  - 저장 위치: 항상 로컬 바탕화면(C:\\Users\\[사용자(컴퓨터)의 이름]\\Desktop)에 저장됩니다.")
    print(f"{C_HELP}  - 파일 이름: 원본 이름 그대로 (예: orderList.xlsx → orderList.xls).")
    print(f"{C_HELP}  - 이미 같은 이름의 .xls 파일이 있다면 자동으로 지우고 새로 저장합니다.")
    print(f"{C_HELP}  - 지울 수 없으면 숫자 붙여서 저장합니다. (예: orderList_1.xls), Excel에서 파일을 닫아주세요.")
    print(f"{C_HELP}  - 변환을 마치면 바탕화면 폴더가 자동으로 열립니다.")
    
    print(f"{C_TITLE}\n[ 명령어 (검은 창에서 입력) ]")
    print(f"{C_TEXT}  {C_CMD}--help{C_TEXT}: 이 가이드 다시 보기.")
    print(f"{C_TEXT}  {C_CMD}clear{C_TEXT}: 화면 깨끗이 지우기.")
    print(f"{C_TEXT}  {C_CMD}exit 또는 quit{C_TEXT}: 프로그램 종료.")
    
    print(f"{C_WARN}\n[ 주의사항 (문제 발생 시 확인) ]")
    print(f"{C_WARN}  - Microsoft Excel이 컴퓨터에 설치되어 있어야 합니다.")
    print(f"{C_WARN}  - 변환 전에, 저장될 파일(예: orderList.xls)이 Excel에서 열려 있지 않은지 확인하세요.")
    print(f"{C_WARN}  - OneDrive나 클라우드 폴더에 파일이 있으면 잠길 수 있습니다. 로컬 바탕화면에 복사해서 사용하세요.")
    print(f"{C_WARN}  - 문제가 생기면 오류 메시지를 읽고, Excel을 닫거나 파일 경로를 확인하세요.")
    print(f"{C_HELP}----------------------------------------------------------\n")

def get_desktop_path() -> str:
    """로컬 바탕화면 경로를 반환하며, OneDrive를 우회합니다. (안정성 강화)"""
    try:
        user_profile = os.path.expanduser("~")
        local_desktop = os.path.join(user_profile, "Desktop")
        
        if os.path.exists(local_desktop):
            logging.info("로컬 Desktop 경로 사용.")
            return local_desktop
        
        korean_desktop = os.path.join(user_profile, "바탕화면")
        if os.path.exists(korean_desktop):
            logging.info("한국어 바탕화면 경로 사용.")
            return korean_desktop
        
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
        desktop_path = winreg.QueryValueEx(key, "Desktop")[0]
        winreg.CloseKey(key)
        if not desktop_path.lower().startswith(os.path.join(user_profile, "onedrive").lower()):
            logging.info("레지스트리 Desktop 경로 사용.")
            return desktop_path
        
        return local_desktop
    except Exception as e:
        logging.error(f"Desktop 경로 조회 실패: {e}")
        raise RuntimeError("바탕화면 경로를 찾을 수 없습니다. 해결: 사용자 폴더에 'Desktop' 또는 '바탕화면' 폴더가 있는지 확인하세요.")

def clear_console():
    """콘솔 화면을 지웁니다. (OS 독립적)"""
    os.system('cls' if os.name == 'nt' else 'clear')
    logging.info("콘솔 화면 지움.")

# --- 코어 로직 모듈 (Core Logic Module) ---

def create_excel_instance():
    """Excel COM 인스턴스를 백그라운드 모드로 생성합니다. (인스턴스 재사용으로 성능 최적화)"""
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        excel.AlertBeforeOverwriting = False
        logging.info("Excel 인스턴스 생성 성공.")
        return excel
    except Exception as e:
        logging.error(f"Excel 인스턴스 생성 실패: {e}")
        raise RuntimeError(f"Excel을 시작할 수 없습니다. 해결: Microsoft Excel이 설치되어 있는지, 또는 실행 중인지 확인하세요.")

def open_workbook_silent(excel, file_path: str):
    """파일을 열고, 입력을 안전하게 처리합니다. (보안: 경로 정규화)"""
    file_abs_path = os.path.abspath(file_path.strip())  # 입력 sanitization
    file_ext = os.path.splitext(file_abs_path)[1].lower()
    
    logging.info(f"파일 열기 시도: {file_abs_path}")
    print(C_TEXT + "  - 파일 여는 중...")
    
    if file_ext == '.csv':
        excel.Workbooks.OpenText(
            Filename=file_abs_path,
            Origin=65001,  # UTF-8
            DataType=1,    # xlDelimited
            Comma=True,
            Local=True
        )
        logging.info("CSV 파일 열기 성공.")
        return excel.ActiveWorkbook
    else:
        try:
            wb = excel.Workbooks.Open(
                Filename=file_abs_path,
                Password=DEFAULT_PASSWORD,
                UpdateLinks=XL_UPDATE_LINKS_NEVER,
                ReadOnly=False,
                Format=1,
                IgnoreReadOnlyRecommended=False
            )
            print(C_HELP + f"  - '{DEFAULT_PASSWORD}' 암호로 열기 성공.")
            logging.info("암호로 파일 열기 성공.")
            return wb
        except Exception:
            try:
                wb = excel.Workbooks.Open(
                    Filename=file_abs_path,
                    UpdateLinks=XL_UPDATE_LINKS_NEVER,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=False
                )
                print(C_HELP + "  - 암호 없이 열기 성공.")
                logging.info("암호 없이 파일 열기 성공.")
                return wb
            except Exception as e:
                logging.error(f"파일 열기 실패: {e}")
                raise ValueError(f"파일을 열 수 없습니다. 해결: 파일이 손상되었거나, 암호가 '1234'가 아닌지, Excel에서 이미 열려 있는지 확인하세요.")

def remove_password(wb):
    """워크북 암호를 제거합니다. (안정성: 속성 확인)"""
    try:
        wb.WriteResPassword = ""
        if hasattr(wb, 'Password'):
            wb.Password = ""
        print(C_HELP + "  - 암호 제거 완료.")
        logging.info("암호 제거 성공.")
    except Exception as e:
        print(C_HELP + "  - 암호 제거 시도됨 (오류 무시).")
        logging.warning(f"암호 제거 중 오류: {e}")

def get_alternative_filename(output_path: str) -> str:
    """대체 파일명 생성 (성능: 최소 반복)"""
    base, ext = os.path.splitext(output_path)
    counter = 1
    while True:
        new_path = f"{base}_{counter}{ext}"
        if not os.path.exists(new_path):
            logging.info(f"대체 파일명 생성: {new_path}")
            return new_path
        counter += 1

def save_as_xls(wb, output_path: str):
    """XLS로 저장 (보안: 덮어쓰기 안전 처리)"""
    print(C_TEXT + "  - '.xls'로 변환 및 저장 중...")
    final_output_path = output_path
    if os.path.exists(output_path):
        try:
            os.remove(output_path)
            print(C_HELP + "  - 기존 파일 삭제 완료.")
            logging.info("기존 파일 삭제 성공.")
        except OSError as e:
            print(C_WARN + f"  - 파일 삭제 실패 (사용 중). 해결: '{os.path.basename(output_path)}' 파일을 Excel에서 닫아주세요.")
            logging.warning(f"파일 삭제 실패: {e}")
            final_output_path = get_alternative_filename(output_path)
            print(C_HELP + f"  - 대체 이름 사용: {os.path.basename(final_output_path)}")
    
    try:
        wb.SaveAs(
            Filename=final_output_path,
            FileFormat=XL_EXCEL8,
            ConflictResolution=XL_FORCE_OVERWRITE,
            Password="",
            WriteResPassword=""
        )
        logging.info(f"저장 성공: {final_output_path}")
        return final_output_path
    except Exception as e:
        logging.error(f"저장 실패: {e}")
        raise RuntimeError(f"파일 저장 실패. 해결: 바탕화면에 쓰기 권한이 있는지, 또는 Excel에서 파일이 열려 있는지 확인하세요.")

def close_resources(wb=None, excel=None):
    """리소스 정리 (안정성: 무조건 실행)"""
    try:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.ScreenUpdating = True
            excel.Quit()
        logging.info("Excel 리소스 정리 완료.")
        print(C_HELP + "  - Excel 정리 완료.\n")
    except Exception as e:
        logging.warning(f"리소스 정리 중 오류: {e}")

def process_file(excel, file_path: str):
    """단일 파일 처리 (모듈화: 독립 로직)"""
    logging.info(f"파일 처리 시작: {file_path}")
    print(C_TEXT + f"\n[ 작업 시작 ] \"{os.path.basename(file_path)}\"")
    
    wb = None
    try:
        if not os.path.exists(file_path) or not os.path.isfile(file_path):
            raise FileNotFoundError(f"파일 없음: {file_path}. 해결: 파일 경로를 확인하고, 파일이 실제로 존재하는지 확인하세요.")
        
        wb = open_workbook_silent(excel, file_path)
        remove_password(wb)
        
        desktop_path = get_desktop_path()
        file_base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(desktop_path, f"{file_base_name}.xls")
        
        final_output_path = save_as_xls(wb, output_path)
        
        print(C_SUCCESS + "\n[ 성공 ]")
        print(C_SUCCESS + f"  - 저장: {final_output_path}")
        
        os.startfile(desktop_path)
        print(C_HELP + "  - 바탕화면 열림.")
        logging.info("작업 성공.")
    
    except FileNotFoundError as e:
        print(C_ERROR + f"  오류: {e}")
        logging.error(f"파일 없음: {e}")
    except ValueError as e:
        print(C_ERROR + f"  오류: {e}")
        logging.error(f"열기 실패: {e}")
    except RuntimeError as e:
        print(C_ERROR + f"  오류: {e}")
        logging.error(f"저장 실패: {e}")
    except Exception as e:
        print(C_ERROR + "\n[ 실패 ]")
        print(C_ERROR + f"  - 파일: {file_path}")
        print(C_ERROR + f"  - 상세: {e}")
        print(C_ERROR + "  해결: Microsoft Excel이 설치되어 있는지, 파일이 손상되었는지, 또는 OneDrive 동기화로 잠겼는지 확인하세요.")
        logging.error(f"예외: {traceback.format_exc()}")
    
    finally:
        if wb:
            wb.Close(SaveChanges=False)
            wb = None

# --- 메인 모듈 (Main Module) ---

def main():
    """메인 진입점 (인스턴스 재사용)"""
    os.system(f'cmd /c "color 0F"')
    clear_console()
    print_splash()
    logging.info("프로그램 시작.")

    excel = None
    try:
        excel = create_excel_instance()
        
        if len(sys.argv) > 1:
            args = [arg.strip('"') for arg in sys.argv[1:]]  # 입력 sanitization
            
            if args[0].lower() == '--help':
                print_help()
                input(C_HELP + "\nEnter로 종료...")
                return

            print(C_TITLE + f"[ 자동 모드 ] {len(args)}개 파일 처리.\n")
            for file_path in args:
                process_file(excel, file_path)
            
            print(C_SUCCESS + "완료.")
            input(C_HELP + "Enter로 종료...")
            logging.info("자동 모드 완료.")

        else:
            print_help()
            print(C_TITLE + "[ 대화형 모드 ] 파일 끌어다 놓기 + Enter.")
            
            while True:
                try:
                    user_input = input(C_TEXT + f"입력 (exit 종료): ").strip().strip('"')
                    
                    if not user_input:
                        continue
                    
                    cmd = user_input.lower()
                    if cmd in ('exit', 'quit'):
                        print(C_TEXT + "종료.")
                        logging.info("사용자 종료.")
                        break
                    elif cmd == 'clear':
                        clear_console()
                        print_splash()
                        print_help()
                    elif cmd == '--help':
                        print_help()
                    else:
                        process_file(excel, user_input)
                
                except EOFError:
                    logging.info("EOF 종료.")
                    break
                except Exception as e:
                    print(C_ERROR + f"오류: {e}. 해결: 입력한 파일 경로가 올바른지 확인하세요.")
                    logging.error(f"대화형 오류: {e}")
    
    finally:
        close_resources(None, excel)
        logging.info("프로그램 종료.")

if __name__ == "__main__":
    main()