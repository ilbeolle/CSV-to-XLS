import sys
import os
import time
import logging
import traceback
import winreg
from pathlib import Path
import win32com.client
import pythoncom
from colorama import init, Fore, Style, Back

class Config:
    APP_NAME = "Excel/CSV to XLS Converter"
    VERSION = "2.1.0"
    AUTHOR = "DongHyun LEE"
    LAST_MODIFIED = "2025-11-24"
    DEFAULT_PASSWORD = "1234"
    XL_EXCEL8 = 56
    XL_FORCE_OVERWRITE = 2
    XL_UPDATE_LINKS_NEVER = 3
    
    C_TITLE = Style.BRIGHT + Fore.LIGHTCYAN_EX + Back.BLACK
    C_SUCCESS = Style.BRIGHT + Fore.LIGHTGREEN_EX + Back.BLACK
    C_ERROR = Style.BRIGHT + Fore.LIGHTRED_EX + Back.BLACK
    C_WARN = Fore.LIGHTYELLOW_EX + Back.BLACK
    C_TEXT = Fore.LIGHTWHITE_EX + Back.BLACK
    C_HELP = Fore.LIGHTBLACK_EX + Back.BLACK
    C_CMD = Fore.LIGHTMAGENTA_EX + Back.BLACK
    C_RESET = Style.RESET_ALL

logging.basicConfig(level=logging.INFO, format="%(message)s", handlers=[logging.StreamHandler(sys.stdout)])

class SystemUtils:
    @staticmethod
    def get_desktop_path() -> Path:
        user_home = Path.home()
        local_desktop = user_home / "Desktop"
        korean_desktop = user_home / "바탕화면"

        if local_desktop.exists():
            return local_desktop
        if korean_desktop.exists():
            return korean_desktop

        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
            desktop_reg, _ = winreg.QueryValueEx(key, "Desktop")
            winreg.CloseKey(key)
            desktop_path_str = os.path.expandvars(desktop_reg)
            return Path(desktop_path_str)
        except Exception:
            return local_desktop

    @staticmethod
    def clear_console():
        os.system('cls' if os.name == 'nt' else 'clear')

class ExcelAutomation:
    def __init__(self):
        self.app = None

    def __enter__(self):
        try:
            pythoncom.CoInitialize()
            self.app = win32com.client.Dispatch("Excel.Application")
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self.app.ScreenUpdating = False
            self.app.AskToUpdateLinks = False
            return self
        except Exception:
            raise RuntimeError("Excel initialization failed.")

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.app:
            try:
                if self.app.Workbooks.Count > 0:
                    for wb in self.app.Workbooks:
                        wb.Close(SaveChanges=False)
                self.app.Quit()
            except Exception:
                pass
            finally:
                self.app = None
                pythoncom.CoUninitialize()

    def process_file(self, file_path: Path) -> bool:
        if not file_path.exists():
            print(f"{Config.C_ERROR} 오류: 파일이 존재하지 않습니다: {file_path.name}")
            return False

        wb = None
        try:
            print(f"{Config.C_TEXT} [작업 중] {file_path.name}")
            wb = self._open_workbook(file_path)
            self._remove_password(wb)
            
            desktop = SystemUtils.get_desktop_path()
            output_filename = f"{file_path.stem}.xls"
            output_path = desktop / output_filename
            
            final_path = self._save_workbook(wb, output_path)
            print(f"{Config.C_SUCCESS} [성공] 저장 완료: {final_path.name}")
            return True

        except Exception as e:
            print(f"{Config.C_ERROR} [실패] {file_path.name}")
            print(f"{Config.C_ERROR}  - 사유: {str(e)}")
            return False
        finally:
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass

    def _open_workbook(self, file_path: Path):
        ext = file_path.suffix.lower()
        abs_path = str(file_path.resolve())

        if ext == '.csv':
            self.app.Workbooks.OpenText(Filename=abs_path, Origin=65001, DataType=1, Comma=True, Local=True)
            return self.app.ActiveWorkbook
        else:
            try:
                return self.app.Workbooks.Open(
                    Filename=abs_path, Password=Config.DEFAULT_PASSWORD, UpdateLinks=Config.XL_UPDATE_LINKS_NEVER,
                    ReadOnly=False, Format=1, IgnoreReadOnlyRecommended=True
                )
            except Exception:
                return self.app.Workbooks.Open(
                    Filename=abs_path, UpdateLinks=Config.XL_UPDATE_LINKS_NEVER,
                    ReadOnly=False, IgnoreReadOnlyRecommended=True
                )

    def _remove_password(self, wb):
        try:
            wb.WriteResPassword = ""
            if hasattr(wb, 'Password'):
                wb.Password = ""
            print(f"{Config.C_HELP}  - 암호 제거/확인 완료")
        except Exception:
            pass

    def _save_workbook(self, wb, target_path: Path) -> Path:
        final_path = target_path
        counter = 1
        while final_path.exists():
            try:
                final_path.unlink()
                print(f"{Config.C_HELP}  - 기존 파일 덮어쓰기 준비 완료")
            except OSError:
                print(f"{Config.C_WARN}  - 기존 파일 사용 중. 이름 변경 시도.")
                final_path = target_path.with_name(f"{target_path.stem}_{counter}{target_path.suffix}")
                counter += 1
        
        wb.SaveAs(
            Filename=str(final_path), FileFormat=Config.XL_EXCEL8,
            ConflictResolution=Config.XL_FORCE_OVERWRITE, Password="", WriteResPassword=""
        )
        return final_path

class ConverterApp:
    def __init__(self):
        init(autoreset=True)

    def run(self):
        SystemUtils.clear_console()
        self._print_splash()
        
        if len(sys.argv) > 1:
            args = [arg.strip('"') for arg in sys.argv[1:]]
            if args[0].lower() == '--help':
                self._print_help()
                input(f"\n{Config.C_WARN}Enter를 누르면 종료합니다...")
                return
            self._run_batch_mode(args)
        else:
            self._print_help()
            self._run_interactive_mode()

    def _print_splash(self):
        print(Config.C_TITLE + "==========================================================")
        print(Config.C_TITLE + f"  {Config.APP_NAME} (v{Config.VERSION})")
        print(Config.C_TITLE + "==========================================================")
        print(Config.C_WARN  + f"  - 작성자: {Config.AUTHOR}")
        print(Config.C_WARN  + f"  - 최종 수정일: {Config.LAST_MODIFIED}")
        print(Config.C_TEXT)

    def _print_help(self):
        print(f"{Config.C_TITLE}\n[ 사용 가이드 ]")
        print(f"{Config.C_TEXT}  - 이 프로그램은 최신 Excel(.xlsx, .xls) 또는 CSV(.csv) 파일을 오래된 Excel 97-2003(.xls) 형식으로 바꿉니다.")
        print(f"{Config.C_TEXT}  - 암호가 '1234'인 파일을 자동으로 풀고, 변환 후 암호를 제거합니다.")
        print(f"{Config.C_TEXT}  - 네이버 스마트스토어, ESM Plus(G마켓, 옥션), 현대이지웰(복지몰)에 대응합니다.")
        
        print(f"{Config.C_TITLE}\n[ 쉬운 방법: 복수 자동 모드 ]")
        print(f"{Config.C_TEXT}  1. 변환할 파일(들)을 마우스로 선택하세요.")
        print(f"{Config.C_TEXT}  2. 선택한 파일을 이 프로그램(.exe) 아이콘 위로 끌어다 놓으세요.")
        print(f"{Config.C_TEXT}  3. 자동으로 변환되어 바탕화면에 저장합니다.")

        print(f"{Config.C_TITLE}\n[ 고급 방법: 단일 대화 모드 ]")
        print(f"{Config.C_TEXT}  1. 프로그램(.exe)을 실행하세요.")
        print(f"{Config.C_TEXT}  2. 변환할 파일을 이 검은 창으로 끌어다 놓으세요.")
        print(f"{Config.C_TEXT}  3. Enter 키를 누르면 변환을 시작합니다!")
        
        print(f"{Config.C_TITLE}\n[ 변환 시 알아둘 점 ]")
        print(f"{Config.C_HELP}  - 저장 위치: 항상 로컬 바탕화면(C:\\Users\\[사용자]\\Desktop)에 저장합니다.")
        print(f"{Config.C_HELP}  - 이미 같은 이름의 파일이 있다면 덮어쓰거나, 사용할 수 없을 경우 숫자를 붙여 저장합니다.")
        
        print(f"{Config.C_WARN}\n[ 주의사항 (문제 발생 시 확인) ]")
        print(f"{Config.C_WARN}  - 반드시 Microsoft Excel이 컴퓨터에 설치되어 있어야 합니다.")
        print(f"{Config.C_WARN}  - OneDrive나 클라우드 폴더에 파일이 있으면 오류가 날 수 있습니다. 바탕화면으로 복사 후 시도하세요.")
        
        print(f"{Config.C_TITLE}\n[ 명령어 ]")
        print(f"{Config.C_TEXT}  {Config.C_CMD}exit{Config.C_TEXT}: 종료 | {Config.C_CMD}clear{Config.C_TEXT}: 화면 지움 | {Config.C_CMD}--help{Config.C_TEXT}: 도움말 다시 보기")
        print(Config.C_HELP + "-" * 58 + "\n")

    def _run_batch_mode(self, files):
        print(Config.C_TITLE + f"[ 자동 모드 ] {len(files)}개 파일 처리 시작.\n")
        success_count = 0
        with ExcelAutomation() as bot:
            for file_str in files:
                if bot.process_file(Path(file_str)):
                    success_count += 1
        
        self._open_desktop()
        self._countdown_exit(success_count == len(files))

    def _run_interactive_mode(self):
        print(Config.C_TITLE + "[ 대화형 모드 ] 파일을 끌어다 놓으세요 (종료: exit).")
        
        with ExcelAutomation() as bot:
            while True:
                try:
                    user_input = input(f"{Config.C_WARN}>>> 입력: {Config.C_RESET}").strip().strip('"')
                    
                    if not user_input: continue
                    if user_input.lower() in ('exit', 'quit'):
                        print("프로그램을 종료합니다.")
                        break
                    if user_input.lower() == 'clear':
                        SystemUtils.clear_console()
                        self._print_splash()
                        self._print_help()
                        continue
                    if user_input.lower() == '--help':
                        self._print_help()
                        continue

                    path = Path(user_input)
                    if bot.process_file(path):
                        self._open_desktop()
                        
                except KeyboardInterrupt:
                    print("\n강제 종료.")
                    break
                except Exception as e:
                    print(f"{Config.C_ERROR} 오류 발생: {e}")

    def _open_desktop(self):
        try:
            os.startfile(SystemUtils.get_desktop_path())
            print(f"{Config.C_HELP}  - 바탕화면 폴더를 열었습니다.")
        except Exception:
            pass

    def _countdown_exit(self, success):
        if success:
            print(f"\n{Config.C_SUCCESS} 모든 작업이 완료되었습니다. 3초 후 자동으로 닫힙니다.")
            for i in range(3, 0, -1):
                print(f"{i}...", end=' ', flush=True)
                time.sleep(1)
        else:
            input(f"\n{Config.C_WARN} 일부 오류가 발생했습니다. 내용을 확인하고 Enter를 누르면 종료합니다.")

if __name__ == "__main__":
    try:
        app = ConverterApp()
        app.run()
    except Exception:
        traceback.print_exc()
        input("치명적인 오류 발생. Enter를 눌러 종료하세요.")