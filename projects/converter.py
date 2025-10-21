"""
Excel/CSV to XLS (Excel 97-2003) Converter
------------------------------------------------
이 프로그램은 .xlsx, .xls, .csv 파일을 구형 .xls 포맷으로 변환합니다.
Drag-n-Drop을 지원하며, '1234' 암호를 자동으로 처리하고 제거합니다.

- 작성자: DongHyun LEE
- 버전: 1.4.0  # 개선 버전
- 최종 수정일: 2025-10-21
"""

import sys
import os
import traceback
import win32com.client
from colorama import init, Fore, Style, Back
import time
import logging

# --- 로깅 설정 (개발자 친화성 강화) ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- 상수 정의 (Constants) ---
__author__ = "DongHyun LEE"
__version__ = "1.4.0"
__last_modified__ = "2025-10-21"
__contact__ = "ilbeolle@gmail.com"

# Excel 파일 포맷 상수
XL_EXCEL8 = 56  # .xls (Excel 97-2003)
XL_FORCE_OVERWRITE = 2  # 덮어쓰기 강제 (xlLocalSessionChanges)
XL_UPDATE_LINKS_NEVER = 3  # 링크 업데이트 안함

# 하드코딩된 비밀번호 (사용자 지정 예외)
DEFAULT_PASSWORD = "1234"

# --- 색상 정의 (Color Palette) ---
init(autoreset=True)
C_TEXT = Fore.WHITE + Back.BLACK
C_TITLE = Style.BRIGHT + Fore.CYAN + Back.BLACK
C_AUTHOR = Style.BRIGHT + Fore.YELLOW + Back.BLACK
C_SUCCESS = Style.BRIGHT + Fore.GREEN + Back.BLACK
C_ERROR = Style.BRIGHT + Fore.RED + Back.BLACK
C_HELP = Fore.LIGHTBLACK_EX + Back.BLACK
C_CMD = Fore.MAGENTA + Back.BLACK
C_WARN = Fore.LIGHTYELLOW_EX + Back.BLACK

# --- 유틸리티 함수 (Utility Functions) ---

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
    """사용 방법(Help)을 간결하고 명확히 출력합니다."""
    print(C_TITLE + "\n[ 사용 방법 ]")
    print(C_HELP + "----------------------------------------------------------")
    print(C_TEXT + "Excel(.xlsx, .xls) 또는 CSV(.csv) 파일을")
    print(C_TEXT + "구형 Excel 97-2003(.xls) 포맷으로 변환합니다.")
    
    print(C_TITLE + "\n[ 1. 자동 모드 (권장) ]")
    print(C_TEXT + "  1. 변환할 파일을 선택 후 프로그램(.exe) 위로 드래그 앤 드롭.")
    print(C_TEXT + "  2. 자동으로 변환 후 바탕화면에 저장.")
    print(C_TEXT + "  3. Enter 키로 종료.")

    print(C_TITLE + "\n[ 2. 대화형 모드 ]")
    print(C_TEXT + "  1. 프로그램(.exe)을 더블클릭 실행.")
    print(C_TEXT + "  2. 파일을 CMD 창에 드래그 앤 드롭 후 Enter.")
    
    print(C_TITLE + "\n[ 변환 규칙 ]")
    print(C_HELP + "  - 암호 '1234' 자동 해제, 변환 후 암호 제거.")
    print(C_HELP + f"  - 저장: 원본 이름으로 로컬 바탕화면(C:\\Users\\[사용자]\\Desktop)에 저장.")
    print(C_HELP + f"  - {C_WARN}덮어쓰기: 기존 .xls 파일은 삭제 후 저장. 파일이 사용 중이면 숫자를 추가 (예: orderList_1.xls).{C_HELP}")
    print(C_HELP + "  - 변환 성공 시 바탕화면 폴더 자동 열림.")
    
    print(C_TITLE + "\n[ 명령어 ]")
    print(C_TEXT + f"  {C_CMD}--help{C_TEXT}: 이 도움말 표시.")
    print(C_TEXT + f"  {C_CMD}clear{C_TEXT}: 화면 지우기.")
    print(C_TEXT + f"  {C_CMD}exit/quit{C_TEXT}: 프로그램 종료.")
    print(C_TEXT + f"  {C_CMD}--output <경로>{C_TEXT}: 출력 경로 커스텀 (예: --output C:\\Temp).")
    
    print(C_WARN + "\n[ 참고 ]")
    print(C_WARN + "  - Microsoft Excel 설치 필수.")
    print(C_WARN + "  - 변환 전, 출력 파일(예: orderList.xls)이 Excel에서 열려 있지 않은지 확인.")
    print(C_WARN + "  - OneDrive 동기화 폴더 사용 시 파일 잠금 발생 가능. 로컬 바탕화면 권장.")
    print(C_HELP + "----------------------------------------------------------\n")

def get_desktop_path() -> str:
    """
    로컬 바탕화면 경로를 반환하며, OneDrive 경로를 명시적으로 우회합니다.
    """
    user_profile = os.path.expanduser("~")
    local_desktop = os.path.join(user_profile, "Desktop")
    
    if os.path.exists(local_desktop):
        return local_desktop
    
    korean_desktop = os.path.join(user_profile, "바탕화면")
    if os.path.exists(korean_desktop):
        return korean_desktop
    
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
        desktop_path = winreg.QueryValueEx(key, "Desktop")[0]
        winreg.CloseKey(key)
        if not desktop_path.lower().startswith(os.path.join(user_profile, "onedrive").lower()):
            return desktop_path
    except:
        pass
    
    return local_desktop

def clear_console():
    """운영체제에 맞춰 콘솔 화면을 지웁니다."""
    os.system('cls' if os.name == 'nt' else 'clear')

# --- 핵심 로직 함수 (Core Logic Functions) ---

def create_excel_instance():
    """Excel COM 인스턴스를 완전 백그라운드 모드로 생성합니다."""
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        excel.AskToUpdateLinks = False
        excel.AlertBeforeOverwriting = False
        logger.info("Excel 인스턴스 생성 성공.")
        return excel
    except Exception as e:
        logger.error(f"Excel 인스턴스 생성 실패: {e}")
        raise RuntimeError(f"Excel 인스턴스 생성 실패: {e}")

def open_workbook_silent(excel, file_path: str):
    """파일을 자동 모드로 엽니다. 모든 다이얼로그 차단."""
    file_abs_path = os.path.abspath(file_path)
    file_ext = os.path.splitext(file_abs_path)[1].lower()
    
    print(C_TEXT + "  - 파일 여는 중...")
    logger.info(f"파일 열기 시도: {file_abs_path}")
    
    if file_ext == '.csv':
        excel.Workbooks.OpenText(
            Filename=file_abs_path,
            Origin=65001,  # UTF-8
            DataType=1,    # xlDelimited
            Comma=True,
            Local=True
        )
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
            print(C_HELP + f"  - '{DEFAULT_PASSWORD}' 암호로 파일 열기 성공.")
            logger.info("암호로 파일 열기 성공.")
            return wb
        except Exception:
            try:
                wb = excel.Workbooks.Open(
                    Filename=file_abs_path,
                    UpdateLinks=XL_UPDATE_LINKS_NEVER,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=False
                )
                print(C_HELP + "  - 암호 없이 파일 열기 성공.")
                logger.info("암호 없이 파일 열기 성공.")
                return wb
            except Exception as e:
                logger.error(f"파일 열기 실패: {e}")
                raise ValueError(f"파일 열기 실패 (암호 불일치 또는 파일 손상): {e}")

def remove_password(wb):
    """워크북의 암호를 제거합니다."""
    try:
        wb.WriteResPassword = ""
        if hasattr(wb, 'Password'):
            wb.Password = ""
        print(C_HELP + "  - 파일 암호 제거 완료.")
        logger.info("파일 암호 제거 완료.")
    except Exception as e:
        print(C_HELP + "  - (암호 제거 시도됨)")
        logger.warning(f"암호 제거 실패: {e}")

def get_alternative_filename(output_path: str) -> str:
    """파일이 잠겨 있을 경우 대체 파일명을 생성합니다 (예: orderList_1.xls)."""
    base, ext = os.path.splitext(output_path)
    counter = 1
    while True:
        new_path = f"{base}_{counter}{ext}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1

def save_as_xls(wb, output_path: str):
    """워크북을 XLS 형식으로 저장합니다. 기존 파일 존재 시 강제 삭제 또는 대체 파일명 사용."""
    print(C_TEXT + "  - '.xls' 포맷으로 변환 및 저장 중...")
    logger.info(f"저장 시도: {output_path}")
    
    final_output_path = output_path
    if os.path.exists(output_path):
        try:
            os.remove(output_path)
            print(C_HELP + "  - 기존 파일 삭제 완료 (강제 덮어쓰기).")
            logger.info("기존 파일 삭제 완료.")
        except OSError as e:
            print(C_WARN + f"  - 기존 파일 삭제 실패 (파일 사용 중): {e}")
            logger.warning(f"기존 파일 삭제 실패: {e}")
            final_output_path = get_alternative_filename(output_path)
            print(C_HELP + f"  - 대체 파일명 사용: {os.path.basename(final_output_path)}")
            logger.info(f"대체 파일명 사용: {final_output_path}")
    
    wb.SaveAs(
        Filename=final_output_path,
        FileFormat=XL_EXCEL8,
        ConflictResolution=XL_FORCE_OVERWRITE,
        Password="",
        WriteResPassword=""
    )
    return final_output_path

def close_resources(wb=None, excel=None):
    """Excel 리소스를 안전하게 정리합니다."""
    try:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.ScreenUpdating = True
            excel.Quit()
        logger.info("Excel 리소스 정리 완료.")
    except Exception as e:
        logger.warning(f"리소스 정리 실패: {e}")
    print(C_HELP + "  - Excel 리소스 정리 완료.\n")

def process_file(excel, file_path: str, output_dir: str = None):
    """단일 파일을 변환하는 핵심 로직. 출력 디렉토리 커스텀 지원."""
    print(C_TEXT + f"\n[ 작업 시작 ] \"{os.path.basename(file_path)}\"")
    logger.info(f"작업 시작: {file_path}")
    
    wb = None
    try:
        if not os.path.exists(file_path) or not os.path.isfile(file_path):
            raise FileNotFoundError(f"유효하지 않은 파일 경로: {file_path}")
        
        wb = open_workbook_silent(excel, file_path)
        remove_password(wb)
        
        desktop_path = output_dir or get_desktop_path()
        file_base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(desktop_path, f"{file_base_name}.xls")
        
        final_output_path = save_as_xls(wb, output_path)
        
        print(C_SUCCESS + "\n[ 변환 성공 ]")
        print(C_SUCCESS + f"  - 저장 위치: {final_output_path}")
        logger.info(f"변환 성공: {final_output_path}")
        
        os.startfile(desktop_path)
        print(C_HELP + f"  - 출력 폴더({desktop_path})를 열었습니다.")
    
    except FileNotFoundError as e:
        print(C_ERROR + f"  오류: 파일을 찾을 수 없습니다. {e}")
        logger.error(f"파일 찾기 실패: {e}")
    except ValueError as e:
        print(C_ERROR + f"  오류: 파일 열기 실패. {e}")
        logger.error(f"파일 열기 실패: {e}")
    except Exception:
        print(C_ERROR + "\n[ 변환 실패 ]")
        print(C_ERROR + f"  - 파일: {file_path}")
        print(C_ERROR + "  - 오류 내용:")
        print(C_ERROR + traceback.format_exc())
        logger.error(f"변환 실패: {traceback.format_exc()}")
    
    finally:
        if wb:
            wb.Close(SaveChanges=False)
        wb = None

# --- 메인 실행 함수 ---

def parse_arguments(args):
    """커맨드라인 인자 파싱 (출력 경로 커스텀 지원)."""
    output_dir = None
    files_to_process = []
    for i, arg in enumerate(args[1:]):
        if arg == '--output' and i + 1 < len(args) - 1:
            output_dir = args[i + 2].strip('"')
            continue
        files_to_process.append(arg.strip('"'))
    return files_to_process, output_dir

def main():
    """프로그램의 메인 진입점. Excel 인스턴스 재사용으로 성능 최적화."""
    os.system(f'cmd /c "color 0F"')
    clear_console()
    print_splash()

    excel = None
    try:
        excel = create_excel_instance()
        
        args = sys.argv
        files_to_process, output_dir = parse_arguments(args)
        
        if files_to_process and files_to_process[0].lower() == '--help':
            print_help()
            input(C_HELP + "\n도움말 출력이 완료되었습니다. Enter 키를 눌러 종료합니다...")
            return

        if len(files_to_process) > 0:
            print(C_TITLE + f"[ 완전 자동 모드 ] {len(files_to_process)}개의 파일을 처리합니다.\n")
            for file_path in files_to_process:
                process_file(excel, file_path, output_dir)
            
            print(C_SUCCESS + "모든 작업이 완료되었습니다.")
            input(C_HELP + "프로그램을 종료하려면 Enter 키를 누르십시오...")

        else:
            print_help()
            print(C_TITLE + "[ 대화형 모드 ] 변환할 파일을 이 창으로 끌어다 놓은 후 Enter를 누르세요.")
            
            while True:
                try:
                    user_input = input(C_TEXT + f"파일 경로 입력 (종료: {C_CMD}exit{C_TEXT}): ").strip().strip('"')

                    if not user_input:
                        continue

                    cmd = user_input.lower()
                    if cmd in ('exit', 'quit'):
                        print(C_TEXT + "프로그램을 종료합니다.")
                        break
                    elif cmd == 'clear':
                        clear_console()
                        print_splash()
                        print_help()
                    elif cmd == '--help':
                        print_help()
                    else:
                        process_file(excel, user_input, output_dir)

                except EOFError:
                    break
                except Exception as e:
                    print(C_ERROR + f"예상치 못한 오류: {e}")
                    logger.error(f"예상치 못한 오류: {e}")
    
    finally:
        close_resources(None, excel)

if __name__ == "__main__":
    main()