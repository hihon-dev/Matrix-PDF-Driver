"""Matrix PDF-Driver — 인쇄하면 PDF를 만들어 주는 가상 프린터.

Windows 프린터가 127.0.0.1의 RAW 포트로 보내는 PostScript를 받아
Ghostscript로 PDF로 바꾼 뒤, 정해 둔 폴더에 저장하거나 저장 위치를 묻는다.

실행 형태는 셋이다.
  --run       SYSTEM 계정으로 상시 실행. 인쇄 데이터를 받아 변환한다.
  --agent     로그인한 사용자마다 실행. 저장 위치를 묻는 창을 띄운다.
  --settings  설정 화면.

Python 3.8(Windows 7 지원 마지막 버전)에서 동작해야 하므로 표준 라이브러리만 쓴다.
"""

from __future__ import annotations

import argparse
import ctypes
import json
import logging
import os
import queue
import re
import shutil
import socket
import subprocess
import sys
import threading
import time
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

APP_NAME = "Matrix PDF-Driver"
APP_VERSION = "1.0.0"
PUBLISHER = "Hihon Inc."
EXE_NAME = "MatrixPdfDriver.exe"
PORT_NAME = "MatrixPdfDriver"
DEFAULT_TCP_PORT = 9100
DEFAULT_PRINTER_NAME = APP_NAME
DEFAULT_FILENAME_PREFIX = "Cube_"
DEFAULT_FILENAME_PATTERN = "{datetime}"
DEFAULT_OUTPUT_FOLDER = r"C:\MatrixPDF"
SAVE_MODES = ("auto", "ask")
PAGE_ROTATIONS = ("keep", "auto")
COLOR_MODES = ("keep", "gray")
# 값은 Ghostscript의 -dPDFA 번호. 0은 보통 PDF.
PDF_FORMATS = {"pdf": 0, "pdfa1": 1, "pdfa2": 2, "pdfa3": 3}
GS_FILES = ("gswin32c.exe", "gsdll32.dll")
LICENSE_FILES = ("LICENSE.txt", "THIRD-PARTY.txt")

# 파일 이름과 폴더 규칙에 넣을 수 있는 항목. 설정 화면의 "항목 넣기" 목록 순서이기도 하다.
TOKENS = (
    ("{title}", "인쇄한 문서 이름"),
    ("{datetime}", "날짜와 시각"),
    ("{date}", "날짜"),
    ("{time}", "시각"),
    ("{year}", "연도"),
    ("{month}", "월"),
    ("{day}", "일"),
    ("{user}", "인쇄한 사람"),
    ("{computer}", "PC 이름"),
    ("{printer}", "프린터 이름"),
    ("{counter}", "일련번호"),
)

# (드라이버 이름, 그 드라이버가 들어 있는 Windows 기본 inf 후보)
DRIVER_CANDIDATES = (
    ("MS Publisher Color Printer", ("prnge001.inf", "ntprint.inf")),
    ("Microsoft PS Class Driver", ("prnms005.inf",)),
)

HEADER_SCAN_BYTES = 64 * 1024
TAIL_SCAN_BYTES = 8 * 1024
RECEIVE_IDLE_TIMEOUT = 300
OTHER_USER_GRACE_SECONDS = 30
STALE_NOTICE_SECONDS = 120
CREATE_NO_WINDOW = 0x08000000
DETACHED_PROCESS = 0x00000008
UNINSTALL_KEY = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Matrix PDF-Driver"
RUN_KEY = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

log = logging.getLogger("matrix-pdf-driver")
KEEP_PIDS: set = set()

# 화면 언어. 문구는 한글로 적어 두고, 영어일 때 아래 표로 옮긴다.
LANGUAGES = (("en", "English"), ("ko", "한국어"))
LANGUAGE_CODES = tuple(code for code, _label in LANGUAGES)
DEFAULT_LANGUAGE = "en"
LANGUAGE = DEFAULT_LANGUAGE


def set_language(code: str) -> None:
    global LANGUAGE
    LANGUAGE = code if code in LANGUAGE_CODES else DEFAULT_LANGUAGE


def tr(text: str) -> str:
    """지금 화면 언어에 맞는 문구. 표에 없으면 한글 그대로 돌려준다."""
    return text if LANGUAGE == "ko" else TRANSLATIONS.get(text, text)


TRANSLATIONS = {
    '저장':
        'Save',
    '파일 이름':
        'File name',
    '품질':
        'Quality',
    '출력':
        'Output',
    '정보':
        'About',
    '닫기':
        'Close',
    '설치':
        'Install',
    '찾아보기':
        'Browse',
    '취소':
        'Cancel',
    '다음':
        'Next',
    '이전':
        'Back',
    '마침':
        'Finish',
    '언어':
        'Language',
    '저장 방식':
        'Save mode',
    '매번 묻기':
        'Ask every time',
    '저장 폴더에서 저장 창이 열립니다.':
        'A save dialog opens in the save folder.',
    '자동 저장':
        'Save automatically',
    '저장 폴더에 묻지 않고 바로 저장합니다.':
        'Saves to the save folder without asking.',
    '저장 폴더':
        'Save folder',
    '서브 폴더 생성':
        'Add subfolder',
    '고른 항목이 저장 폴더 뒤에 이어 붙고, 그 이름의 폴더가 인쇄할 때 자동으로 만들어집니다. 폴더를 한 단계 더 나누려면 사이에 \\ 를 직접 넣으세요.':
        'The item you pick is added to the end of the save folder, and that folder is created when you print. To add another level, type \\ between items yourself.',
    'PDF를 저장할 폴더':
        'Folder for PDF files',
    '앞 이름':
        'Prefix',
    '추가 이름':
        'Name items',
    '추가 이름 생성':
        'Add name item',
    '전체':
        'Full path',
    '위 세 칸 중 어디를 고쳐도 서로 맞춰집니다. 같은 이름의 파일이 이미 있으면 뒤에 _2, _3이 붙습니다.':
        'Editing any of the three fields updates the others. If a file with the same name exists, _2, _3 is added.',
    '인쇄한 문서 이름':
        'Document name',
    '날짜와 시각':
        'Date and time',
    '날짜':
        'Date',
    '시각':
        'Time',
    '연도':
        'Year',
    '월':
        'Month',
    '일':
        'Day',
    '인쇄한 사람':
        'Printed by',
    'PC 이름':
        'Computer name',
    '프린터 이름':
        'Printer name',
    '일련번호':
        'Counter',
    '시험성적서':
        'Report',
    '형식':
        'Format',
    'PDF/A-2b (권장)':
        'PDF/A-2b (recommended)',
    '그림 품질':
        'Image quality',
    '원본 그대로':
        'Keep original',
    '그림을 손대지 않습니다.':
        'Images are left untouched.',
    '파일 크기 줄이기':
        'Reduce file size',
    '그림을 압축합니다. 조금 흐려질 수 있습니다.':
        'Images are compressed and may look slightly softer.',
    '페이지 방향':
        'Page rotation',
    '인쇄한 그대로':
        'As printed',
    '글자 방향에 맞춰 자동으로 돌리기':
        'Rotate to match text direction',
    '색상':
        'Color',
    '인쇄한 그대로 (컬러)':
        'As printed (color)',
    '흑백':
        'Grayscale',
    '세로·가로 용지 방향은 인쇄하는 프로그램의 인쇄 창에서 고릅니다.':
        'Portrait or landscape paper is chosen in the print dialog of the program you print from.',
    '할 일':
        'After printing',
    '변환 과정 보여주기':
        'Show conversion progress',
    '생성 후 PDF 열기':
        'Open the PDF when it is ready',
    '동시 출력':
        'Also print to',
    '쓸 수 있는 프린터가 없습니다':
        'No printer available',
    '변환 과정은 화면 가운데에 진행 바로 잠깐 표시됩니다. 동시 출력은 PDF를 만들면서 고른 프린터로도 인쇄합니다.':
        'Progress appears briefly as a bar in the middle of the screen. "Also print to" sends each PDF to the selected printer as well.',
    '제품':
        'Product',
    '제작':
        'Publisher',
    '주식회사 히온 (Hihon Inc.)':
        'Hihon Inc.',
    '라이선스':
        'License',
    '저장 폴더 열기':
        'Open save folder',
    '실패한 인쇄물 보기':
        'View failed jobs',
    '다음 인쇄물이 저장될 곳':
        'Where the next printout will be saved',
    '저장 창이 처음 열릴 폴더와 이름':
        'Folder and name the save dialog starts with',
    '●  동작 중':
        '●  Running',
    '●  멈춰 있음':
        '●  Stopped',
    '저장했습니다. 다음 인쇄부터 적용됩니다.':
        'Saved. Changes apply from the next printout.',
    '설치하는 중입니다. 잠시만 기다려 주세요.':
        'Installing. Please wait.',
    '설치하지 못했습니다. %s':
        'Installation failed. %s',
    '설치가 끝났습니다. 프린터를 "%s"로 골라 인쇄해 보세요.':
        'Installation finished. Print to "%s" from any program.',
    '저장 방식을 골라 주세요.':
        'Choose a save mode.',
    '저장 폴더는 C:\\ 처럼 드라이브부터 시작하는 전체 경로로 적어 주세요.':
        'Enter the save folder as a full path that starts with a drive, such as C:\\.',
    '앞 이름에는 \\ / : * ? " < > | 를 쓸 수 없습니다.':
        'The prefix cannot contain \\ / : * ? " < > |.',
    '앞 이름이나 추가 이름 중 하나는 있어야 합니다.':
        'Enter a prefix or at least one name item.',
    '동시에 출력할 프린터를 골라 주세요.':
        'Choose the printer to also print to.',
    '규칙에 쓸 수 없는 항목이 있습니다. 목록에서 골라 넣어 주세요.':
        'The rule contains an item that cannot be used. Pick items from the list.',
    'PDF 저장 — %s':
        'Save PDF — %s',
    'PDF 문서':
        'PDF document',
    '저장하지 않고 이 인쇄물을 버릴까요?':
        'Discard this printout without saving?',
    '저장하지 못했습니다.\n\n%s\n\n다른 위치를 골라 주세요.':
        'Could not save the file.\n\n%s\n\nChoose another location.',
    'PDF를 만드는 중입니다':
        'Creating the PDF',
    'PDF를 만들지 못했습니다':
        'Could not create the PDF',
    '저장했습니다':
        'Saved',
    'PDF를 열지 못했습니다':
        'Could not open the PDF',
    '%s 프린터로 출력하지 못했습니다':
        'Could not print to %s',
    '인쇄하지 못했습니다.':
        'Printing failed.',
    'PostScript 인쇄 데이터가 아닙니다. 프린터 드라이버 설정을 확인하세요.':
        'The print data is not PostScript. Check the printer driver setting.',
    'Ghostscript 실행 파일을 찾을 수 없습니다.':
        'The Ghostscript program was not found.',
    '변환 제한 시간(%d초)을 넘었습니다.':
        'Conversion took longer than the limit (%d seconds).',
    'Ghostscript를 실행하지 못했습니다: %s':
        'Could not run Ghostscript: %s',
    'Ghostscript 오류(코드 %d)\n%s':
        'Ghostscript error (code %d)\n%s',
    'PDF를 저장하지 못했습니다: %s':
        'Could not save the PDF: %s',
    '관리자 권한으로 실행해야 합니다.':
        'Administrator rights are required.',
    '프린터 포트 등록 도구(prnport.vbs)를 찾을 수 없습니다.':
        'The printer port tool (prnport.vbs) was not found.',
    '프린터 포트를 등록하지 못했습니다.\n%s':
        'Could not add the printer port.\n%s',
    'Windows 기본 PostScript 프린터 드라이버를 찾을 수 없습니다.':
        'No built-in Windows PostScript printer driver was found.',
    '프린터를 등록하지 못했습니다.\n%s':
        'Could not add the printer.\n%s',
    '자동 실행을 등록하지 못했습니다.\n%s':
        'Could not register automatic start.\n%s',
    '사용할 수 있는 포트 번호(9100~9119)가 없습니다.':
        'No free port number is available (9100-9119).',
    '설치는 빌드된 실행 파일(%s)로만 할 수 있습니다.':
        'Installation only works from the built program (%s).',
    '설치 파일에 Ghostscript가 들어 있지 않습니다.':
        'The installer does not contain Ghostscript.',
    '설치 폴더는 C:\\Program Files\\%s 처럼 드라이브와 폴더 이름까지 적어 주세요.':
        'Enter the install folder with a drive and a folder name, such as C:\\Program Files\\%s.',
    'Windows에서만 설치할 수 있습니다.':
        'Installation is only possible on Windows.',
    '설치가 끝났습니다.':
        'Installation finished.',
    '%s를 제거할까요?\n저장된 PDF와 설정은 지우지 않습니다.':
        'Remove %s?\nSaved PDF files and settings are kept.',
    '제거했습니다.':
        'Removed.',
    '%s 설치를 시작합니다':
        'Welcome to %s Setup',
    '인쇄할 때 프린터를 "%s"로 고르면 PDF 파일이 만들어집니다.':
        'Choose "%s" as the printer when you print, and a PDF file is created.',
    '버전 %s\n제작 주식회사 히온 (Hihon Inc.)':
        'Version %s\nPublisher: Hihon Inc.',
    '이 PC에 이미 설치되어 있습니다. 계속하면 새 파일로 바꾸고, 쓰던 설정은 그대로 둡니다.':
        'It is already installed on this PC. Continuing replaces the program files and keeps your settings.',
    '계속하려면 "다음"을 누르세요.':
        'Click Next to continue.',
    '이 프로그램은 GNU AGPL v3 조건으로 제공됩니다. 함께 설치되는 다른 프로그램의 라이선스는 설치 폴더의 THIRD-PARTY.txt에 있습니다.':
        'This program is provided under the GNU AGPL v3. Licenses of the other programs installed with it are in THIRD-PARTY.txt in the install folder.',
    '위 라이선스 내용에 동의합니다':
        'I accept the license',
    '설치 폴더':
        'Install folder',
    '프로그램 파일을 넣을 폴더입니다. 바꾸지 않아도 됩니다.':
        'The program files go into this folder. You do not need to change it.',
    '설치할 폴더':
        'Install folder',
    '필요한 공간: 약 35MB':
        'Space required: about 35 MB',
    '만들어진 PDF가 저장되는 폴더는 설치 뒤 설정에서 정합니다.':
        'The folder for the PDF files is chosen in Setting after installation.',
    '설치하는 중입니다':
        'Installing',
    '설치가 끝났습니다':
        'Installation finished',
    '프린터 목록에 "%s"가 추가됐습니다. 아무 프로그램에서나 이 프린터로 인쇄해 보세요.':
        '"%s" was added to your printers. Print to it from any program.',
    '저장 폴더, 파일 이름, 품질은 시작 메뉴의 "%s Setting"에서 바꿉니다.':
        'Change the save folder, file name and quality in "%s Setting" on the Start menu.',
    '지금 설정 열기':
        'Open Setting now',
    '라이선스에 동의해야 설치할 수 있습니다.':
        'Accept the license to install.',
    '프린터를 등록하고 파일을 복사하고 있습니다. 10~20초쯤 걸립니다.':
        'Adding the printer and copying files. This takes 10 to 20 seconds.',
}


# ---------------------------------------------------------------- 경로와 설정

def app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", app_dir()))
    return base / "assets" / name


def data_dir() -> Path:
    root = os.environ.get("PROGRAMDATA") or os.environ.get("ALLUSERSPROFILE")
    return Path(root or Path.home()) / APP_NAME


def install_dir() -> Path:
    # 32비트 실행 파일이 64비트 Windows에서 돌 때 ProgramFiles는 (x86)을 가리킨다.
    root = os.environ.get("ProgramW6432") or os.environ.get("ProgramFiles") or r"C:\Program Files"
    return Path(root) / APP_NAME


def installed_dir() -> Path:
    """실제로 설치된 폴더. 설치할 때 고른 위치를 프로그램 목록 정보에서 읽고, 없으면 기본 위치."""
    try:
        import winreg
        access = winreg.KEY_READ | winreg.KEY_WOW64_64KEY
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, UNINSTALL_KEY, 0, access) as key:
            location = str(winreg.QueryValueEx(key, "InstallLocation")[0] or "")
            if location:
                return Path(location)
    except (OSError, ImportError):
        pass
    return install_dir()


def program_files_in(folder: Path) -> List[Path]:
    """설치 폴더에 이 프로그램이 넣는 파일들. 제거할 때 이것만 지운다."""
    return [folder / EXE_NAME] + [folder / "gs" / name for name in GS_FILES] + [folder / name for name in LICENSE_FILES]


def normalize_install_folder(chosen: str) -> str:
    """찾아보기로 고른 폴더 아래에 제품 이름 폴더를 붙인다. 이미 그 이름이면 그대로 둔다."""
    path = os.path.normpath(chosen.strip())
    return path if os.path.basename(path).casefold() == APP_NAME.casefold() else os.path.join(path, APP_NAME)


def check_install_folder(folder: str) -> Path:
    folder = folder.strip()
    if not re.match(r"^[A-Za-z]:\\[^\\]", folder) or re.search(r'[*?"<>|]', folder):
        raise ValueError(tr("설치 폴더는 C:\\Program Files\\%s 처럼 드라이브와 폴더 이름까지 적어 주세요.") % APP_NAME)
    return Path(os.path.normpath(folder))


def default_config_path() -> Path:
    return data_dir() / "config.json"


def split_name_rule(raw: Dict[str, Any]) -> Tuple[str, str]:
    """파일 이름 규칙을 (항상 같은 앞 이름, 그 뒤에 붙는 항목 규칙)으로 나눈다.

    앞 이름을 따로 저장하지 않던 설정은 첫 항목 앞의 글자를 앞 이름으로 본다.
    """
    if "filename_prefix" in raw:
        return str(raw.get("filename_prefix") or ""), str(raw.get("filename_pattern") or "")
    rule = str(raw.get("filename_pattern") or (DEFAULT_FILENAME_PREFIX + DEFAULT_FILENAME_PATTERN))
    index = rule.find("{")
    return (rule, "") if index < 0 else (rule[:index], rule[index:])


class Config:
    def __init__(self, raw: Dict[str, Any]):
        mode = str(raw.get("save_mode") or "auto").lower()
        self.save_mode = mode if mode in SAVE_MODES else "auto"
        self.output_folder = os.path.expandvars(str(raw.get("output_folder") or DEFAULT_OUTPUT_FOLDER))
        self.port = int(raw.get("port") or DEFAULT_TCP_PORT)
        self.printer_name = str(raw.get("printer_name") or DEFAULT_PRINTER_NAME)
        self.filename_prefix, self.filename_pattern = split_name_rule(raw)
        quality = str(raw.get("image_quality") or "lossless").lower()
        self.image_quality = quality if quality in ("lossless", "compact") else "lossless"
        rotation = str(raw.get("page_rotation") or "keep").lower()
        self.page_rotation = rotation if rotation in PAGE_ROTATIONS else "keep"
        color = str(raw.get("color_mode") or "keep").lower()
        self.color_mode = color if color in COLOR_MODES else "keep"
        pdf_format = str(raw.get("pdf_format") or "pdf").lower()
        self.pdf_format = pdf_format if pdf_format in PDF_FORMATS else "pdf"
        language = str(raw.get("language") or DEFAULT_LANGUAGE).lower()
        self.language = language if language in LANGUAGE_CODES else DEFAULT_LANGUAGE
        self.show_progress = bool(raw.get("show_progress", False))
        self.open_after_save = bool(raw.get("open_after_save", False))
        self.also_print = bool(raw.get("also_print", False))
        self.also_print_printer = str(raw.get("also_print_printer") or "")
        self.log_dir = Path(os.path.expandvars(str(raw.get("log_dir") or (data_dir() / "logs"))))
        self.convert_timeout_seconds = max(30, int(raw.get("convert_timeout_seconds") or 600))
        self.ghostscript_path = str(raw.get("ghostscript_path") or "")

    @property
    def name_rule(self) -> str:
        """앞 이름과 항목 규칙을 합친 전체 규칙. 앞 이름의 중괄호는 글자 그대로 쓴다."""
        return self.filename_prefix.replace("{", "{{").replace("}", "}}") + self.filename_pattern

    @property
    def notifies_helper(self) -> bool:
        return self.show_progress or self.open_after_save or self.also_print

    @property
    def spool_dir(self) -> Path:
        return data_dir() / "spool"

    @property
    def failed_dir(self) -> Path:
        return data_dir() / "failed"

    @property
    def pending_dir(self) -> Path:
        return data_dir() / "pending"


def read_raw_config(path: Path) -> Dict[str, Any]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def load_config(path: Path) -> Config:
    return Config(read_raw_config(path))


def find_ghostscript(config: Config) -> Optional[Path]:
    if config.ghostscript_path:
        explicit = Path(os.path.expandvars(config.ghostscript_path))
        return explicit if explicit.exists() else None
    base = app_dir()
    for folder in (base / "gs", base, base / "vendor" / "ghostscript"):
        candidate = folder / GS_FILES[0]
        if candidate.exists():
            return candidate
    return None


def setup_logging(log_dir: Path, name: str = "driver.log") -> None:
    log.setLevel(logging.INFO)
    log.handlers.clear()
    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        handler = RotatingFileHandler(str(log_dir / name), maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8")
        handler.setFormatter(formatter)
        log.addHandler(handler)
    except OSError:
        pass
    if sys.stdout is not None:
        console = logging.StreamHandler(sys.stdout)
        console.setFormatter(formatter)
        log.addHandler(console)


# ---------------------------------------------------------------- 인쇄 데이터 읽기

def find_postscript_start(head: bytes) -> int:
    """PostScript 본문 시작 위치. 드라이버가 앞에 붙이는 PJL·Ctrl-D를 건너뛴다. 없으면 -1."""
    return head.find(b"%!PS")


def trim_trailer(path: Path) -> None:
    """마지막 %%EOF 뒤에 붙은 Ctrl-D·PJL 종료 문자열을 잘라낸다."""
    size = path.stat().st_size
    scan = min(size, TAIL_SCAN_BYTES)
    with path.open("r+b") as handle:
        handle.seek(size - scan)
        tail = handle.read(scan)
        index = tail.rfind(b"%%EOF")
        if index < 0:
            return
        end = size - scan + index + len(b"%%EOF")
        if end < size:
            handle.truncate(end)
            handle.seek(end)
            handle.write(b"\n")


def _decode_text(raw: bytes) -> str:
    for encoding in ("utf-8", "mbcs" if os.name == "nt" else "cp949", "cp949"):
        try:
            return raw.decode(encoding)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("latin-1")


def parse_dsc(head: bytes, key: str) -> str:
    """DSC 주석(%%Title, %%For 등) 값을 읽는다. 한글은 <16진수> 형태로 들어온다."""
    match = re.search(rb"^%%" + key.encode("ascii") + rb":[ \t]*(.*?)[ \t]*\r?$", head, re.MULTILINE)
    if not match:
        return ""
    value = match.group(1)
    if value.startswith(b"<") and value.endswith(b">"):
        try:
            value = bytes.fromhex(value[1:-1].decode("ascii"))
        except ValueError:
            return ""
        if value.startswith(b"\xef\xbb\xbf"):
            value = value[3:]
    elif value.startswith(b"(") and value.endswith(b")"):
        value = _unescape_ps_string(value[1:-1])
    return _decode_text(value).strip()


def _unescape_ps_string(raw: bytes) -> bytes:
    """PostScript 문자열의 \\270 같은 8진수 표기와 \\( \\) \\\\ 를 원래 바이트로 되돌린다."""
    named = {b"n": b"\n", b"r": b"\r", b"t": b"\t", b"b": b"\b", b"f": b"\f"}

    def restore(match):
        code = match.group(1)
        if code[:1].isdigit():
            return bytes([int(code, 8) & 0xFF])
        return named.get(code, code)

    return re.sub(rb"\\([0-7]{1,3}|.)", restore, raw)


def notice_order(path: Path) -> Tuple[str, int]:
    """한 인쇄물의 알림은 '만드는 중'이 '저장했습니다'보다 먼저 처리되어야 한다."""
    stem, _, kind = path.name[:-len(".json")].rpartition(".")
    return stem, {"start": 0, "ask": 1, "saved": 2, "failed": 3}.get(kind, 9)


def parse_title(head: bytes) -> str:
    return parse_dsc(head, "Title")


# ---------------------------------------------------------------- 이름 규칙

def sanitize_filename(text: str, limit: int = 80) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\x00-\x1f]', " ", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned[:limit].rstrip(" .")


def token_values(title: str, user: str, when: datetime, counter: int, printer: str) -> Dict[str, Any]:
    return {
        "title": sanitize_filename(title) or "print",
        "datetime": when.strftime("%Y%m%d_%H%M%S"),
        "date": when.strftime("%Y%m%d"),
        "time": when.strftime("%H%M%S"),
        "year": when.strftime("%Y"),
        "month": when.strftime("%m"),
        "day": when.strftime("%d"),
        "user": sanitize_filename(user) or "unknown",
        "computer": sanitize_filename(socket.gethostname()) or "PC",
        "printer": sanitize_filename(printer) or APP_NAME,
        "counter": counter,
    }


def sample_values() -> Dict[str, Any]:
    user = os.environ.get("USERNAME") or "user"
    return token_values(tr("시험성적서"), user, datetime.now(), 1, DEFAULT_PRINTER_NAME)


def check_template(template: str) -> None:
    """규칙에 모르는 항목이나 깨진 괄호가 있으면 ValueError."""
    try:
        template.format(**sample_values())
    except (KeyError, IndexError, ValueError):
        raise ValueError(tr("규칙에 쓸 수 없는 항목이 있습니다. 목록에서 골라 넣어 주세요."))


def build_filename(pattern: str, values: Dict[str, Any]) -> str:
    try:
        name = pattern.format(**values)
    except (KeyError, IndexError, ValueError):
        name = (DEFAULT_FILENAME_PREFIX + DEFAULT_FILENAME_PATTERN).format(**values)
    return sanitize_filename(name, limit=150) or "print"


def render_folder(template: str, values: Dict[str, Any]) -> Path:
    try:
        rendered = template.format(**values)
    except (KeyError, IndexError, ValueError):
        rendered = static_folder_root(template)
    return Path(os.path.normpath(rendered))


def static_folder_root(template: str) -> str:
    """폴더 규칙에서 항목이 들어가기 전까지의, 항상 같은 앞부분."""
    parts = []
    for part in re.split(r"[\\/]", template):
        if "{" in part:
            break
        parts.append(part)
    return "\\".join(parts) if parts else template


def unique_pdf_path(folder: Path, base: str) -> Path:
    candidate = folder / (base + ".pdf")
    number = 2
    while candidate.exists() or candidate.with_name(candidate.name + ".tmp").exists():
        candidate = folder / ("%s_%d.pdf" % (base, number))
        number += 1
    return candidate


def next_counter() -> int:
    path = data_dir() / "counter.txt"
    try:
        value = int(path.read_text(encoding="ascii").strip()) + 1
    except (OSError, ValueError):
        value = 1
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(str(value), encoding="ascii")
    except OSError:
        pass
    return value


# ---------------------------------------------------------------- 변환

def pdfa_definition(gray: bool) -> str:
    """PDF/A에 꼭 들어가야 하는 색 기준(ICC 프로필)을 붙이는 PostScript. 프로필은 Ghostscript에 내장된 것을 쓴다."""
    return (
        "%%!\n"
        "[/_objdef {icc_PDFA} /type /stream /OBJ pdfmark\n"
        "[{icc_PDFA} << /N %d >> /PUT pdfmark\n"
        "[{icc_PDFA} (%%rom%%iccprofiles/%s) (r) file /PUT pdfmark\n"
        "[/_objdef {OutputIntent_PDFA} /type /dict /OBJ pdfmark\n"
        "[{OutputIntent_PDFA} << /Type /OutputIntent /S /GTS_PDFA1 /DestOutputProfile {icc_PDFA} "
        "/OutputConditionIdentifier (%s) >> /PUT pdfmark\n"
        "[{Catalog} << /OutputIntents [ {OutputIntent_PDFA} ] >> /PUT pdfmark\n"
    ) % ((1, "default_gray.icc", "sGray") if gray else (3, "default_rgb.icc", "sRGB"))


def rewrite_font_names_as_utf8(source: Path, target: Path) -> int:
    """인쇄 데이터 안의 한글 글꼴 이름을 UTF-8로 바꿔 target에 쓴다. 바꾼 개수를 돌려준다.

    Windows 드라이버는 글꼴 이름을 한글 Windows 인코딩으로 넣는데, PDF/A는 UTF-8이 아니면
    Ghostscript가 PDF/A를 포기하고 보통 PDF를 만든다.
    """
    changed = 0

    def fix(match):
        nonlocal changed
        try:
            raw = bytes.fromhex(match.group(1).decode("ascii"))
            raw.decode("utf-8")
            return match.group(0)
        except UnicodeDecodeError:
            changed += 1
            return b"/OrigFontName <" + _decode_text(raw).encode("utf-8").hex().upper().encode("ascii") + b">"
        except ValueError:
            return match.group(0)

    data = re.sub(rb"/OrigFontName\s*<([0-9A-Fa-f\s]+)>", fix, source.read_bytes())
    target.write_bytes(data)
    return changed


def ghostscript_args(gs: Path, quality: str, output: Path, source: Path, rotation: str = "keep",
                     color: str = "keep", pdfa: int = 0, pdfa_def: Optional[Path] = None) -> List[str]:
    args = [
        str(gs), "-dSAFER", "-dBATCH", "-dNOPAUSE", "-q",
        "-sDEVICE=pdfwrite", "-dEmbedAllFonts=true", "-dSubsetFonts=true",
        # 자동 회전은 글자 방향을 보고 쪽을 돌리므로 좌표 기반 추출이 어긋날 수 있어 기본은 끈다.
        "-dAutoRotatePages=" + ("/PageByPage" if rotation == "auto" else "/None"),
    ]
    if color == "gray":
        args += ["-sColorConversionStrategy=Gray", "-dProcessColorModel=/DeviceGray"]
    elif pdfa:
        args.append("-sColorConversionStrategy=RGB")
    if pdfa:
        # 규격에 맞지 않는 요소는 빼고서라도 PDF/A로 만든다.
        args += ["-dPDFA=%d" % pdfa, "-dPDFACompatibilityPolicy=1"]
    if quality == "compact":
        args.append("-dPDFSETTINGS=/printer")
    else:
        args += [
            "-dAutoFilterColorImages=false", "-dColorImageFilter=/FlateEncode",
            "-dAutoFilterGrayImages=false", "-dGrayImageFilter=/FlateEncode",
            "-dDownsampleColorImages=false", "-dDownsampleGrayImages=false",
            "-dDownsampleMonoImages=false",
        ]
    args.append("-sOutputFile=" + str(output))
    if pdfa and pdfa_def is not None:
        args.append(str(pdfa_def))
    args.append(str(source))
    return args


def print_pdf(gs: Path, pdf: Path, printer: str, timeout: int = 600) -> None:
    """PDF를 실제 프린터로도 보낸다. 실패하면 OSError."""
    result = subprocess.run(
        [str(gs), "-dBATCH", "-dNOPAUSE", "-q", "-dNoCancel", "-sDEVICE=mswinpr2",
         "-sOutputFile=%printer%" + printer, str(pdf)],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL, timeout=timeout,
        creationflags=CREATE_NO_WINDOW if os.name == "nt" else 0,
    )
    if result.returncode != 0:
        raise OSError(result.stdout.decode("utf-8", "replace").strip()[-500:] or tr("인쇄하지 못했습니다."))


def list_printers() -> List[str]:
    """이 사용자가 쓸 수 있는 프린터 이름. 자기 자신(가상 프린터)은 뺀다."""
    from ctypes import wintypes

    class PrinterInfo4(ctypes.Structure):
        _fields_ = [("pPrinterName", wintypes.LPWSTR), ("pServerName", wintypes.LPWSTR), ("Attributes", wintypes.DWORD)]

    try:
        spooler = ctypes.WinDLL("winspool.drv")
        flags = 0x2 | 0x4  # 이 PC의 프린터 + 연결해 둔 네트워크 프린터
        needed, count = wintypes.DWORD(0), wintypes.DWORD(0)
        spooler.EnumPrintersW(flags, None, 4, None, 0, ctypes.byref(needed), ctypes.byref(count))
        if not needed.value:
            return []
        buffer = ctypes.create_string_buffer(needed.value)
        if not spooler.EnumPrintersW(flags, None, 4, buffer, needed, ctypes.byref(needed), ctypes.byref(count)):
            return []
        entries = ctypes.cast(buffer, ctypes.POINTER(PrinterInfo4))
        names = [entries[index].pPrinterName for index in range(count.value)]
    except (OSError, AttributeError):
        return []
    return sorted(name for name in names if name and name != DEFAULT_PRINTER_NAME)


def _remove_quietly(path: Path) -> None:
    try:
        path.unlink()
    except OSError:
        pass


class Converter:
    def __init__(self, config: Config):
        self.config = config
        self.gs = find_ghostscript(config)

    def _fail(self, job: Path, reason: str) -> None:
        log.error("변환 실패: %s — %s", job.name, reason)
        try:
            self.config.failed_dir.mkdir(parents=True, exist_ok=True)
            target = self.config.failed_dir / (job.stem + ".ps")
            shutil.move(str(job), str(target))
            target.with_suffix(".reason.txt").write_text(reason, encoding="utf-8")
        except OSError as exc:
            log.error("실패한 인쇄 데이터를 보관하지 못했습니다: %s", exc)

    def _notify(self, kind: str, stem: str, **details: Any) -> None:
        """로그인한 사용자 쪽 도우미에게 알릴 일을 남긴다 (진행 표시, 저장 후 열기·동시 출력)."""
        try:
            pending = self.config.pending_dir
            pending.mkdir(parents=True, exist_ok=True)
            notice = dict(details, type=kind, received_at=time.time())
            partial = pending / ("%s.%s.tmp" % (stem, kind))
            partial.write_text(json.dumps(notice, ensure_ascii=False), encoding="utf-8")
            os.replace(str(partial), str(pending / ("%s.%s.json" % (stem, kind))))
        except OSError as exc:
            log.error("알림을 남기지 못했습니다: %s", exc)

    def convert(self, job: Path, received_at: Optional[datetime] = None) -> Optional[Path]:
        """받은 인쇄 데이터 하나를 PDF로 바꾼다. 실패하면 None.

        자동 저장이면 출력 폴더의 PDF 경로를, 물어보기면 저장 대기 중인 PDF 경로를 돌려준다.
        """
        received_at = received_at or datetime.now()
        with job.open("rb") as handle:
            head = handle.read(HEADER_SCAN_BYTES)
        title, user = parse_dsc(head, "Title"), parse_dsc(head, "For")
        if self.config.show_progress:
            self._notify("start", job.stem, title=title, user=user)
        saved = self._convert(job, head, title, user, received_at)
        if saved is None and self.config.show_progress:
            self._notify("failed", job.stem, title=title, user=user)
        return saved

    def _convert(self, job: Path, head: bytes, title: str, user: str, received_at: datetime) -> Optional[Path]:
        if find_postscript_start(head) != 0:
            self._fail(job, tr("PostScript 인쇄 데이터가 아닙니다. 프린터 드라이버 설정을 확인하세요."))
            return None
        if self.gs is None:
            self._fail(job, tr("Ghostscript 실행 파일을 찾을 수 없습니다."))
            return None

        # Ghostscript에는 영문 경로만 넘기고, 한글이 들어갈 수 있는 출력 폴더로는 직접 복사한다.
        work_pdf = job.with_suffix(".pdf")
        pdfa = PDF_FORMATS[self.config.pdf_format]
        source, pdfa_def = job, None
        try:
            if pdfa:
                source, pdfa_def = job.with_suffix(".utf8"), job.with_suffix(".pdfa")
                rewrite_font_names_as_utf8(job, source)
                pdfa_def.write_text(pdfa_definition(self.config.color_mode == "gray"), encoding="ascii")
            result = subprocess.run(
                ghostscript_args(self.gs, self.config.image_quality, work_pdf, source,
                                 self.config.page_rotation, self.config.color_mode, pdfa, pdfa_def),
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
                timeout=self.config.convert_timeout_seconds,
                creationflags=CREATE_NO_WINDOW if os.name == "nt" else 0,
            )
        except subprocess.TimeoutExpired:
            _remove_quietly(work_pdf)
            self._fail(job, tr("변환 제한 시간(%d초)을 넘었습니다.") % self.config.convert_timeout_seconds)
            return None
        except OSError as exc:
            self._fail(job, tr("Ghostscript를 실행하지 못했습니다: %s") % exc)
            return None
        finally:
            if pdfa:
                _remove_quietly(job.with_suffix(".utf8"))
                _remove_quietly(job.with_suffix(".pdfa"))
        if result.returncode != 0 or not work_pdf.exists() or work_pdf.stat().st_size == 0:
            detail = result.stdout.decode("utf-8", "replace").strip()[-2000:]
            _remove_quietly(work_pdf)
            self._fail(job, tr("Ghostscript 오류(코드 %d)\n%s") % (result.returncode, detail))
            return None

        values = token_values(title, user, received_at, next_counter(), self.config.printer_name)
        base = build_filename(self.config.name_rule, values)
        try:
            if self.config.save_mode == "ask":
                saved = self._hand_over(work_pdf, job.stem, base, title, user,
                                        render_folder(self.config.output_folder, values))
            else:
                saved = self._save(work_pdf, render_folder(self.config.output_folder, values), base)
                if self.config.notifies_helper:
                    self._notify("saved", job.stem, title=title, user=user, name=saved.name, path=str(saved))
        except OSError as exc:
            _remove_quietly(work_pdf)
            self._fail(job, tr("PDF를 저장하지 못했습니다: %s") % exc)
            return None
        _remove_quietly(work_pdf)
        _remove_quietly(job)
        return saved

    def _save(self, work_pdf: Path, folder: Path, base: str) -> Path:
        folder.mkdir(parents=True, exist_ok=True)
        final = unique_pdf_path(folder, base)
        # 폴더를 감시하는 프로그램이 쓰는 중인 파일을 집어가지 않도록 .tmp로 쓴 뒤 이름을 바꾼다.
        partial = final.with_name(final.name + ".tmp")
        shutil.copyfile(str(work_pdf), str(partial))
        os.replace(str(partial), str(final))
        log.info("PDF 저장: %s (%d bytes)", final, final.stat().st_size)
        return final

    def _hand_over(self, work_pdf: Path, stem: str, base: str, title: str, user: str, folder: Path) -> Path:
        """저장 위치를 물어볼 수 있도록 로그인한 사용자 쪽 도우미에게 넘긴다.

        folder는 저장 창이 처음 열릴 기본 폴더다. 서브 폴더 규칙이 있으면 여기서 미리 만들어 둔다.
        """
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            log.error("기본 저장 폴더를 만들지 못했습니다: %s", exc)
        pending = self.config.pending_dir
        pending.mkdir(parents=True, exist_ok=True)
        pdf = pending / (stem + ".pdf")
        shutil.copyfile(str(work_pdf), str(pdf))
        self._notify("ask", stem, title=title, user=user, name=base, pdf=str(pdf), folder=str(folder))
        log.info("저장 위치 확인 대기: %s (%s)", base, user or "사용자 미상")
        return pdf


# ---------------------------------------------------------------- 수신 서버

class PrintServer:
    def __init__(self, config: Config, config_path: Optional[Path] = None):
        self.config = config
        self.config_path = config_path
        self._config_mtime = self._read_config_mtime()
        self.converter = Converter(config)
        self.jobs: "queue.Queue[Optional[Tuple[Path, datetime]]]" = queue.Queue()
        self.stop_event = threading.Event()
        self._counter = 0
        self._lock = threading.Lock()
        self._socket: Optional[socket.socket] = None

    def _next_job_stem(self) -> str:
        with self._lock:
            self._counter += 1
            return "%s_%04d" % (datetime.now().strftime("%Y%m%d_%H%M%S"), self._counter)

    def _receive(self, conn: socket.socket) -> None:
        received_at = datetime.now()
        part = self.config.spool_dir / (self._next_job_stem() + ".part")
        total = 0
        try:
            conn.settimeout(RECEIVE_IDLE_TIMEOUT)
            head = b""
            started = False
            with part.open("wb") as handle:
                while True:
                    chunk = conn.recv(65536)
                    if not chunk:
                        break
                    total += len(chunk)
                    if started:
                        handle.write(chunk)
                        continue
                    head += chunk
                    offset = find_postscript_start(head)
                    if offset >= 0:
                        handle.write(head[offset:])
                        started = True
                    elif len(head) >= HEADER_SCAN_BYTES:
                        handle.write(head)
                        started = True
                if not started:
                    handle.write(head)
        except (OSError, socket.timeout) as exc:
            log.error("인쇄 데이터를 끝까지 받지 못했습니다: %s", exc)
            _remove_quietly(part)
            return
        finally:
            try:
                conn.close()
            except OSError:
                pass
        if total == 0:
            _remove_quietly(part)
            return
        trim_trailer(part)
        job = part.with_suffix(".job")
        os.replace(str(part), str(job))
        log.info("인쇄 데이터 수신: %s (%d bytes)", job.name, total)
        self.jobs.put((job, received_at))

    def _read_config_mtime(self) -> float:
        try:
            return self.config_path.stat().st_mtime if self.config_path else 0.0
        except OSError:
            return 0.0

    def _reload_config_if_changed(self) -> None:
        """설정 화면에서 저장한 내용을 다시 시작하지 않고 다음 인쇄부터 적용한다."""
        mtime = self._read_config_mtime()
        if mtime == self._config_mtime or self.config_path is None:
            return
        self._config_mtime = mtime
        try:
            updated = load_config(self.config_path)
        except (OSError, ValueError) as exc:
            log.error("설정 파일을 읽지 못해 이전 설정을 계속 씁니다: %s", exc)
            return
        # 포트는 프린터 등록과 묶여 있어 실행 중에는 바꾸지 않는다.
        updated.port = self.config.port
        self.config = updated
        self.converter = Converter(updated)
        set_language(updated.language)
        log.info("설정 변경 적용 — 저장 방식 %s, 폴더 %s", updated.save_mode, updated.output_folder)

    def _worker(self) -> None:
        while True:
            item = self.jobs.get()
            if item is None:
                return
            try:
                self._reload_config_if_changed()
                self.converter.convert(item[0], item[1])
            except Exception:
                log.exception("변환 중 예상하지 못한 오류")

    def _recover(self) -> None:
        """이전 실행에서 다 받고도 변환하지 못한 인쇄 데이터를 다시 처리한다."""
        for leftover in self.config.spool_dir.glob("*.part"):
            _remove_quietly(leftover)
        for pattern in ("*.pdf", "*.utf8", "*.pdfa"):
            for leftover in self.config.spool_dir.glob(pattern):
                _remove_quietly(leftover)
        for job in sorted(self.config.spool_dir.glob("*.job")):
            self.jobs.put((job, datetime.fromtimestamp(job.stat().st_mtime)))

    def serve(self) -> None:
        set_language(self.config.language)
        self.config.spool_dir.mkdir(parents=True, exist_ok=True)
        self._recover()
        worker = threading.Thread(target=self._worker, name="convert", daemon=True)
        worker.start()

        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        # 다른 PC에서 인쇄 데이터를 밀어 넣지 못하도록 이 PC 안에서만 연다.
        server.bind(("127.0.0.1", self.config.port))
        server.listen(16)
        server.settimeout(1.0)
        self._socket = server
        log.info("%s %s 시작 — 포트 %d, 저장 방식 %s, 폴더 %s",
                 APP_NAME, APP_VERSION, self.config.port, self.config.save_mode, self.config.output_folder)
        if self.converter.gs is None:
            log.error("Ghostscript 실행 파일이 없어 변환할 수 없습니다.")
        try:
            while not self.stop_event.is_set():
                try:
                    conn, _ = server.accept()
                except socket.timeout:
                    continue
                except OSError:
                    break
                threading.Thread(target=self._receive, args=(conn,), daemon=True).start()
        finally:
            server.close()
            self.jobs.put(None)
            worker.join(timeout=self.config.convert_timeout_seconds)

    def stop(self) -> None:
        self.stop_event.set()


# ---------------------------------------------------------------- 사용자 쪽 도우미

def claim_pending(notice_path: Path, username: str) -> Optional[Dict[str, Any]]:
    """대기 중인 알림 하나를 이 사용자 몫으로 잡는다. 다른 도우미가 먼저 잡았으면 None."""
    try:
        notice = json.loads(notice_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    owner = str(notice.get("user") or "")
    age = time.time() - float(notice.get("received_at") or 0)
    # 여러 사람이 동시에 로그인한 PC에서는 인쇄한 사람에게 먼저 기회를 준다.
    if owner and owner.casefold() != username.casefold() and age < OTHER_USER_GRACE_SECONDS:
        return None
    claimed = notice_path.with_name(notice_path.name + ".claimed")
    try:
        os.rename(str(notice_path), str(claimed))
    except OSError:
        return None
    notice["claimed"] = str(claimed)
    notice["age"] = age
    return notice


def finish_pending(notice: Dict[str, Any], destination: Optional[Path] = None) -> None:
    """알림 처리를 끝낸다. 저장 대기 PDF가 딸려 있으면 고른 위치에 저장하고(None이면 버리고) 치운다."""
    if notice.get("pdf"):
        if destination is not None:
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(notice["pdf"], str(destination))
        _remove_quietly(Path(notice["pdf"]))
    _remove_quietly(Path(notice["claimed"]))


def release_stale_claims(pending: Path, older_than_seconds: int = 600) -> None:
    """저장 창을 띄운 채 도우미가 죽어 묶여 버린 인쇄물을 다시 대기 상태로 돌린다."""
    for claimed in pending.glob("*.claimed"):
        try:
            if time.time() - claimed.stat().st_mtime > older_than_seconds:
                os.rename(str(claimed), str(claimed.with_name(claimed.name[:-len(".claimed")])))
        except OSError:
            pass


def _agent_state_path() -> Path:
    return Path(os.environ.get("APPDATA") or Path.home()) / APP_NAME / "agent.json"


def _last_save_folder() -> str:
    try:
        folder = json.loads(_agent_state_path().read_text(encoding="utf-8")).get("last_folder", "")
        if folder and os.path.isdir(folder):
            return folder
    except (OSError, ValueError):
        pass
    return str(Path.home() / "Documents")


def _remember_save_folder(folder: str) -> None:
    try:
        path = _agent_state_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps({"last_folder": folder}, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def after_save(config: Config, pdf: Path, report=None) -> None:
    """저장한 뒤에 하기로 한 일(열기, 실제 프린터로도 출력)을 한다."""
    if config.open_after_save:
        # 보기 프로그램이 뜨는 동안 몇 초씩 걸릴 수 있어, 진행 표시가 멈추지 않게 따로 돌린다.
        def open_pdf():
            try:
                os.startfile(str(pdf))
            except OSError as exc:
                if report:
                    report(tr("PDF를 열지 못했습니다"), str(exc))

        threading.Thread(target=open_pdf, daemon=True).start()
    if config.also_print and config.also_print_printer and config.also_print_printer != config.printer_name:
        gs = find_ghostscript(config)

        def work():
            try:
                if gs is None:
                    raise OSError(tr("Ghostscript 실행 파일을 찾을 수 없습니다."))
                print_pdf(gs, pdf, config.also_print_printer)
            except (OSError, subprocess.TimeoutExpired) as exc:
                if report:
                    report(tr("%s 프린터로 출력하지 못했습니다") % config.also_print_printer, str(exc))

        threading.Thread(target=work, daemon=True).start()


def run_agent(config_path: Path) -> int:
    # 한 사용자에게 창이 두 번 뜨지 않도록 로그인 세션마다 하나만 돈다.
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "Local\\MatrixPdfDriverAgent")
    if ctypes.windll.kernel32.GetLastError() == 183:
        return 0
    import tkinter as tk
    from tkinter import filedialog, messagebox

    _enable_dpi_awareness()
    root = tk.Tk()
    root.withdraw()
    _apply_window_icon(root)
    root.attributes("-topmost", True)
    username = os.environ.get("USERNAME") or ""
    pending = load_config(config_path).pending_dir
    scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
    card: Dict[str, Any] = {"window": None, "timer": None, "tick": None, "kind": "", "shown_at": 0.0}
    try:
        card_logo = tk.PhotoImage(file=str(resource_path("logo-32.png")))
    except Exception:
        card_logo = None
    failures: "queue.Queue[Tuple[str, str]]" = queue.Queue()

    def close_card() -> None:
        for key in ("timer", "tick"):
            if card[key] is not None:
                root.after_cancel(card[key])
                card[key] = None
        if card["window"] is not None:
            card["window"].destroy()
            card["window"] = None
        card["kind"] = ""

    def show_saved(config: Config, name: str) -> None:
        # 만든 PDF가 바로 열리면 그것이 곧 끝났다는 표시라, "저장했습니다"는 열지 않을 때만 보여 준다.
        if config.show_progress:
            show_card("close" if config.open_after_save else "done", tr("저장했습니다"), name)

    def show_card(kind: str, headline: str = "", detail: str = "") -> None:
        """화면 가운데에 뜨는 진행 표시. kind: busy(만드는 중), done(끝남), error(실패), close(닫기만)."""
        # 변환은 1초도 안 걸려서, 바로 바꾸면 진행 바가 보이지도 않고 지나간다. 잠깐은 보여 준다.
        remaining = 0.6 - (time.time() - card["shown_at"])
        if kind != "busy" and card["kind"] == "busy" and remaining > 0:
            if card["timer"] is not None:
                root.after_cancel(card["timer"])
            card["timer"] = root.after(int(remaining * 1000), lambda: show_card(kind, headline, detail))
            return
        close_card()
        if kind == "close":
            return
        width, pad, bar_height = int(380 * scale), int(22 * scale), max(4, int(6 * scale))
        window = tk.Toplevel(root)
        window.overrideredirect(True)
        window.attributes("-topmost", True)
        window.configure(bg=COLOR_LINE)
        body = tk.Frame(window, bg=COLOR_SURFACE)
        body.pack(fill="both", expand=True, padx=1, pady=1)
        top = tk.Frame(body, bg=COLOR_SURFACE)
        top.pack(fill="x", padx=pad, pady=(pad, 0))
        if card_logo is not None:
            tk.Label(top, image=card_logo, bg=COLOR_SURFACE).pack(side="left", padx=(0, int(10 * scale)))
        tk.Label(top, text=APP_NAME, bg=COLOR_SURFACE, fg=COLOR_MUTED, font=(FONT_FAMILY, 9)).pack(side="left")
        tk.Label(body, text=headline, bg=COLOR_SURFACE, fg=COLOR_ERROR if kind == "error" else COLOR_INK,
                 font=(FONT_FAMILY, 11, "bold"), anchor="w").pack(fill="x", padx=pad, pady=(int(12 * scale), 0))
        tk.Label(body, text=detail or " ", bg=COLOR_SURFACE, fg=COLOR_MUTED, font=(FONT_FAMILY, 9), anchor="w",
                 justify="left", wraplength=width - 2 * pad).pack(fill="x", padx=pad)
        inner = width - 2 * pad - 2
        bar = tk.Canvas(body, width=inner, height=bar_height, bg=COLOR_SURFACE_2, highlightthickness=0, bd=0)
        bar.pack(padx=pad, pady=(int(14 * scale), pad))
        window.update_idletasks()
        height = window.winfo_reqheight()
        window.geometry("%dx%d+%d+%d" % (
            width, height, (root.winfo_screenwidth() - width) // 2, (root.winfo_screenheight() - height) // 2))

        if kind == "busy":
            block, step = inner // 3, max(4, inner // 45)
            runner = bar.create_rectangle(-block, 0, 0, bar_height, fill=COLOR_PRIMARY, width=0)
            position = {"x": -block}

            def tick() -> None:
                position["x"] = -block if position["x"] > inner else position["x"] + step
                bar.coords(runner, position["x"], 0, position["x"] + block, bar_height)
                card["tick"] = root.after(30, tick)

            tick()
        else:
            bar.create_rectangle(0, 0, inner, bar_height, width=0,
                                 fill=COLOR_ERROR if kind == "error" else COLOR_PRIMARY)
        card["window"], card["kind"], card["shown_at"] = window, kind, time.time()
        # 끝났다는 알림을 놓치더라도 '만드는 중' 창이 화면에 계속 남지 않게 한다.
        card["timer"] = root.after({"busy": 15000, "done": 900, "error": 6000}[kind], close_card)

    def starting_folder(notice: Dict[str, Any]) -> str:
        folder = str(notice.get("folder") or "")
        return folder if folder and os.path.isdir(folder) else _last_save_folder()

    def ask_and_save(notice: Dict[str, Any], config: Config) -> None:
        while True:
            chosen = filedialog.asksaveasfilename(
                parent=root, title=tr("PDF 저장 — %s") % APP_NAME, initialdir=starting_folder(notice),
                initialfile=str(notice.get("name") or "print") + ".pdf", defaultextension=".pdf",
                filetypes=[(tr("PDF 문서"), "*.pdf")],
            )
            if not chosen:
                if messagebox.askyesno(APP_NAME, tr("저장하지 않고 이 인쇄물을 버릴까요?"), parent=root):
                    finish_pending(notice, None)
                    return
                continue
            try:
                finish_pending(notice, Path(chosen))
            except OSError as exc:
                messagebox.showerror(APP_NAME, tr("저장하지 못했습니다.\n\n%s\n\n다른 위치를 골라 주세요.") % exc, parent=root)
                continue
            _remember_save_folder(os.path.dirname(chosen))
            show_saved(config, os.path.basename(chosen))
            after_save(config, Path(chosen), lambda headline, detail: failures.put((headline, detail)))
            return

    def handle(notice: Dict[str, Any]) -> None:
        kind = notice.get("type")
        config = load_config(config_path)
        set_language(config.language)
        if kind == "ask":
            ask_and_save(notice, config)
            return
        finish_pending(notice)
        # 아무도 로그인하지 않은 동안 쌓인 알림은 조용히 버린다. 로그인하자마자 PDF가 줄줄이 열리면 곤란하다.
        if notice.get("age", 0) > STALE_NOTICE_SECONDS:
            return
        label = str(notice.get("title") or "")
        if kind == "start" and config.show_progress:
            show_card("busy", tr("PDF를 만드는 중입니다"), label)
        elif kind == "failed":
            show_card("error", tr("PDF를 만들지 못했습니다"), label)
        elif kind == "saved":
            show_saved(config, str(notice.get("name") or ""))
            after_save(config, Path(str(notice.get("path") or "")),
                       lambda headline, detail: failures.put((headline, detail)))

    def poll() -> None:
        try:
            while not failures.empty():
                show_card("error", *failures.get_nowait())
            if pending.is_dir():
                release_stale_claims(pending)
                for notice_path in sorted(pending.glob("*.json"), key=notice_order):
                    notice = claim_pending(notice_path, username)
                    if notice:
                        handle(notice)
        except Exception:
            pass
        root.after(700, poll)

    poll()
    root.mainloop()
    del mutex
    return 0


# ---------------------------------------------------------------- 설치와 제거

class SetupError(Exception):
    pass


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except (AttributeError, OSError):
        return False


def _windows_dir() -> Path:
    return Path(os.environ.get("SystemRoot") or r"C:\Windows")


def _system_dir() -> Path:
    # 32비트 프로세스가 64비트 Windows의 진짜 System32를 보려면 Sysnative를 거쳐야 한다.
    sysnative = _windows_dir() / "Sysnative"
    return sysnative if sysnative.exists() else _windows_dir() / "System32"


def _run(command: List[str], timeout: int = 180) -> Tuple[int, str]:
    try:
        result = subprocess.run(
            command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
            timeout=timeout, creationflags=CREATE_NO_WINDOW,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return 1, str(exc)
    return result.returncode, result.stdout.decode("mbcs", "replace").strip()


def _powershell(script: str) -> Tuple[int, str]:
    exe = _system_dir() / "WindowsPowerShell" / "v1.0" / "powershell.exe"
    return _run([str(exe), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script])


def _has_print_cmdlets() -> bool:
    version = sys.getwindowsversion()
    return (version.major, version.minor) >= (6, 2)


def _registry_key_exists(path: str) -> bool:
    import winreg
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY))
        return True
    except OSError:
        return False


def printer_exists(name: str) -> bool:
    return _registry_key_exists(r"SYSTEM\CurrentControlSet\Control\Print\Printers" + "\\" + name)


def port_exists(name: str) -> bool:
    return _registry_key_exists(
        r"SYSTEM\CurrentControlSet\Control\Print\Monitors\Standard TCP/IP Port\Ports" + "\\" + name
    )


def _prnport_script() -> Optional[Path]:
    matches = sorted((_system_dir() / "Printing_Admin_Scripts").glob("*/prnport.vbs"))
    return matches[0] if matches else None


def create_port(name: str, tcp_port: int) -> None:
    if _has_print_cmdlets():
        code, output = _powershell(
            "Add-PrinterPort -Name '%s' -PrinterHostAddress '127.0.0.1' -PortNumber %d" % (name, tcp_port)
        )
    else:
        script = _prnport_script()
        if script is None:
            raise SetupError(tr("프린터 포트 등록 도구(prnport.vbs)를 찾을 수 없습니다."))
        code, output = _run([
            str(_system_dir() / "cscript.exe"), "//nologo", str(script),
            "-a", "-r", name, "-h", "127.0.0.1", "-o", "raw", "-n", str(tcp_port), "-md",
        ])
    if not port_exists(name):
        raise SetupError(tr("프린터 포트를 등록하지 못했습니다.\n%s") % output)


def delete_port(name: str) -> None:
    if not port_exists(name):
        return
    if _has_print_cmdlets():
        _powershell("Remove-PrinterPort -Name '%s'" % name)
    else:
        script = _prnport_script()
        if script is not None:
            _run([str(_system_dir() / "cscript.exe"), "//nologo", str(script), "-d", "-r", name])


def find_driver() -> List[Tuple[str, Path]]:
    """이 Windows에 기본으로 들어 있는 PostScript 드라이버와 그 inf 위치."""
    found = []
    inf_dir = _windows_dir() / "inf"
    for model, inf_names in DRIVER_CANDIDATES:
        for inf_name in inf_names:
            inf = inf_dir / inf_name
            try:
                content = inf.read_bytes()
            except OSError:
                continue
            if model.encode("ascii") in content or model.encode("utf-16-le") in content:
                found.append((model, inf))
                break
    return found


def _wait_until(check, seconds: int) -> bool:
    deadline = time.time() + seconds
    while time.time() < deadline:
        if check():
            return True
        time.sleep(1)
    return check()


def create_printer(name: str, port_name: str) -> str:
    drivers = find_driver()
    if not drivers:
        raise SetupError(tr("Windows 기본 PostScript 프린터 드라이버를 찾을 수 없습니다."))
    rundll = str(_system_dir() / "rundll32.exe")
    output = ""
    for model, inf in drivers:
        _, output = _run([
            rundll, "printui.dll,PrintUIEntry", "/if", "/b", name, "/f", str(inf),
            "/r", port_name, "/m", model, "/z", "/q",
        ], timeout=300)
        if _wait_until(lambda: printer_exists(name), 30):
            return model
    raise SetupError(tr("프린터를 등록하지 못했습니다.\n%s") % output)


def delete_printer(name: str) -> None:
    if not printer_exists(name):
        return
    _run([str(_system_dir() / "rundll32.exe"), "printui.dll,PrintUIEntry", "/dl", "/n", name, "/q"])
    _wait_until(lambda: not printer_exists(name), 20)


def _task_xml(exe: Path) -> str:
    return """<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>%s</Author>
    <Description>Receives print jobs and saves them as PDF.</Description>
  </RegistrationInfo>
  <Triggers>
    <BootTrigger>
      <Enabled>true</Enabled>
    </BootTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <RestartOnFailure>
      <Interval>PT1M</Interval>
      <Count>999</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>%s</Command>
      <Arguments>--run</Arguments>
    </Exec>
  </Actions>
</Task>
""" % (PUBLISHER, str(exe).replace("&", "&amp;"))


def register_task(exe: Path) -> None:
    xml_path = data_dir() / "task.xml"
    xml_path.write_bytes(b"\xff\xfe" + _task_xml(exe).encode("utf-16-le"))
    schtasks = str(_system_dir() / "schtasks.exe")
    code, output = _run([schtasks, "/Create", "/TN", APP_NAME, "/XML", str(xml_path), "/F"])
    _remove_quietly(xml_path)
    if code != 0:
        raise SetupError(tr("자동 실행을 등록하지 못했습니다.\n%s") % output)
    _run([schtasks, "/Run", "/TN", APP_NAME])


def remove_task() -> None:
    schtasks = str(_system_dir() / "schtasks.exe")
    _run([schtasks, "/End", "/TN", APP_NAME])
    _run([schtasks, "/Delete", "/TN", APP_NAME, "/F"])
    stop_running_copies()


def stop_running_copies() -> None:
    # 단일 실행 파일은 풀어 주는 프로세스와 실제 작업 프로세스 둘로 돌고,
    # 예약 작업을 멈춰도 작업 프로세스가 남아 포트와 파일을 잡고 있다.
    # 지금 설치·제거를 수행 중인 자신과, 그 끝을 기다리는 프로세스들은 남긴다.
    # 기다리는 쪽이 먼저 죽으면 설치 파일이 임시 폴더를 지워 설치가 중간에 깨진다.
    command = [str(_system_dir() / "taskkill.exe"), "/F", "/IM", EXE_NAME]
    for pid in {os.getpid(), os.getppid()} | KEEP_PIDS:
        command += ["/FI", "PID ne %d" % pid]
    _run(command)


def _port_is_free(tcp_port: int) -> bool:
    probe = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        probe.bind(("127.0.0.1", tcp_port))
        return True
    except OSError:
        return False
    finally:
        probe.close()


def choose_tcp_port(preferred: int) -> int:
    for candidate in [preferred] + list(range(DEFAULT_TCP_PORT, DEFAULT_TCP_PORT + 20)):
        if _port_is_free(candidate):
            return candidate
    raise SetupError(tr("사용할 수 있는 포트 번호(9100~9119)가 없습니다."))


def _copy_with_retry(source: Path, target: Path) -> None:
    if source.resolve() == target.resolve():
        return
    for attempt in range(10):
        try:
            shutil.copy2(str(source), str(target))
            return
        except OSError:
            if attempt == 9:
                raise
            # 설치 도중에 누가 설정 창을 열면 파일이 다시 잠긴다. 닫고 다시 시도한다.
            stop_running_copies()
            time.sleep(1)


def _copy_program_files(target: Path) -> Path:
    source = app_dir()
    exe_source = Path(sys.executable).resolve()
    if not getattr(sys, "frozen", False):
        raise SetupError(tr("설치는 빌드된 실행 파일(%s)로만 할 수 있습니다.") % EXE_NAME)
    gs_source = find_ghostscript(Config({}))
    if gs_source is None:
        raise SetupError(tr("설치 파일에 Ghostscript가 들어 있지 않습니다."))
    (target / "gs").mkdir(parents=True, exist_ok=True)
    exe_target = target / EXE_NAME
    _copy_with_retry(exe_source, exe_target)
    for name in GS_FILES:
        _copy_with_retry(gs_source.parent / name, target / "gs" / name)
    for name in LICENSE_FILES:
        if (source / name).exists():
            _copy_with_retry(source / name, target / name)
    return exe_target


def _write_registry_entries(exe: Path, target: Path) -> None:
    import winreg
    access = winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY
    with winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, UNINSTALL_KEY, 0, access) as key:
        winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
        winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
        winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, PUBLISHER)
        winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, str(target))
        winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, str(exe))
        winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, '"%s" --uninstall --ui' % exe)
        winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)
    # 누가 로그인하든 저장 위치를 묻는 도우미가 함께 뜨게 한다.
    with winreg.CreateKeyEx(winreg.HKEY_LOCAL_MACHINE, RUN_KEY, 0, access) as key:
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, '"%s" --agent' % exe)


def _remove_registry_entries() -> None:
    import winreg
    try:
        winreg.DeleteKeyEx(winreg.HKEY_LOCAL_MACHINE, UNINSTALL_KEY, winreg.KEY_WOW64_64KEY, 0)
    except OSError:
        pass
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, RUN_KEY, 0, winreg.KEY_WRITE | winreg.KEY_WOW64_64KEY) as key:
            winreg.DeleteValue(key, APP_NAME)
    except OSError:
        pass


def _prepare_pending_folder(config: Config) -> None:
    # 변환은 SYSTEM이 하고 저장은 로그인한 사용자가 하므로, 사용자도 이 폴더의 파일을 옮기고 지울 수 있어야 한다.
    config.pending_dir.mkdir(parents=True, exist_ok=True)
    _run([str(_system_dir() / "icacls.exe"), str(config.pending_dir), "/grant", "*S-1-5-32-545:(OI)(CI)M"])


def _start_menu_folder() -> Path:
    return data_dir().parent / "Microsoft" / "Windows" / "Start Menu" / "Programs"


def _start_menu_shortcut() -> Path:
    return _start_menu_folder() / ("%s Setting.lnk" % APP_NAME)


def _remove_start_menu_shortcuts() -> None:
    # 예전에 한글 이름("… 설정")으로 만든 바로가기도 함께 지운다.
    for name in ("%s Setting.lnk" % APP_NAME, "%s 설정.lnk" % APP_NAME):
        _remove_quietly(_start_menu_folder() / name)


def create_start_menu_shortcut(exe: Path) -> None:
    script = (
        "$s=(New-Object -ComObject WScript.Shell).CreateShortcut('%s');"
        "$s.TargetPath='%s';$s.Arguments='--settings';$s.WorkingDirectory='%s';$s.Save()"
    ) % (_start_menu_shortcut(), exe, exe.parent)
    _powershell(script)


def start_agent() -> None:
    exe = installed_dir() / EXE_NAME
    if exe.exists():
        subprocess.Popen([str(exe), "--agent"], close_fds=True, creationflags=DETACHED_PROCESS)


def install(config_path: Path, output_folder: Optional[str] = None, save_mode: Optional[str] = None,
            filename_pattern: Optional[str] = None, target_dir: Optional[str] = None,
            language: Optional[str] = None) -> Config:
    if not is_admin():
        raise SetupError(tr("관리자 권한으로 실행해야 합니다."))
    raw = read_raw_config(config_path)
    if output_folder:
        raw["output_folder"] = os.path.abspath(output_folder)
    if save_mode in SAVE_MODES:
        raw["save_mode"] = save_mode
    if filename_pattern:
        check_template(filename_pattern)
        raw.pop("filename_prefix", None)
        raw["filename_pattern"] = filename_pattern
    if language in LANGUAGE_CODES:
        raw["language"] = language
    raw.setdefault("save_mode", "auto")
    raw.setdefault("output_folder", DEFAULT_OUTPUT_FOLDER)
    raw.setdefault("printer_name", DEFAULT_PRINTER_NAME)
    raw["filename_prefix"], raw["filename_pattern"] = split_name_rule(raw)
    raw.setdefault("image_quality", "lossless")

    # 다시 설치할 때 실행 중인 파일과 열려 있는 포트를 먼저 놓아 준다.
    remove_task()
    time.sleep(2)
    raw["port"] = choose_tcp_port(int(raw.get("port") or DEFAULT_TCP_PORT))
    config = Config(raw)

    previous = installed_dir()
    target = check_install_folder(target_dir) if target_dir else previous
    if previous != target:
        # 다른 위치에 다시 설치하면 예전 위치의 파일이 남지 않게 치운다.
        for leftover in program_files_in(previous):
            _remove_quietly(leftover)
        for folder in (previous / "gs", previous):
            try:
                folder.rmdir()
            except OSError:
                pass
    exe = _copy_program_files(target)
    Path(static_folder_root(config.output_folder)).mkdir(parents=True, exist_ok=True)
    data_dir().mkdir(parents=True, exist_ok=True)
    _prepare_pending_folder(config)
    config_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")

    delete_printer(config.printer_name)
    delete_port(PORT_NAME)
    create_port(PORT_NAME, config.port)
    create_printer(config.printer_name, PORT_NAME)
    register_task(exe)
    _write_registry_entries(exe, target)
    _remove_start_menu_shortcuts()
    create_start_menu_shortcut(exe)
    return config


def uninstall(config_path: Path) -> None:
    if not is_admin():
        raise SetupError(tr("관리자 권한으로 실행해야 합니다."))
    config = load_config(config_path)
    remove_task()
    delete_printer(config.printer_name)
    delete_port(PORT_NAME)
    _remove_start_menu_shortcuts()
    target = installed_dir()
    _remove_registry_entries()
    if target.exists():
        # 지금 실행 중인 파일은 스스로 지울 수 없어, 끝난 뒤 지우도록 넘긴다.
        # 설치 폴더를 사용자가 골랐을 수 있으므로 우리가 넣은 파일만 지우고, 폴더는 비었을 때만 없앤다.
        steps = ['del /f /q "%s"' % path for path in program_files_in(target)]
        steps += ['rmdir "%s"' % (target / "gs"), 'rmdir "%s"' % target]
        subprocess.Popen(
            "cmd.exe /c ping -n 4 127.0.0.1 >nul & " + " & ".join(steps),
            creationflags=CREATE_NO_WINDOW, close_fds=True, cwd=str(_windows_dir()),
        )


def relaunch_elevated(arguments: List[str]) -> int:
    """관리자 권한으로 자신을 다시 실행하고 끝날 때까지 기다린다."""
    from ctypes import wintypes

    class ShellExecuteInfo(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD), ("fMask", wintypes.ULONG), ("hwnd", wintypes.HWND),
            ("lpVerb", wintypes.LPCWSTR), ("lpFile", wintypes.LPCWSTR),
            ("lpParameters", wintypes.LPCWSTR), ("lpDirectory", wintypes.LPCWSTR),
            ("nShow", ctypes.c_int), ("hInstApp", wintypes.HINSTANCE),
            ("lpIDList", ctypes.c_void_p), ("lpClass", wintypes.LPCWSTR),
            ("hkeyClass", wintypes.HKEY), ("dwHotKey", wintypes.DWORD),
            ("hIcon", wintypes.HANDLE), ("hProcess", wintypes.HANDLE),
        ]

    info = ShellExecuteInfo()
    info.cbSize = ctypes.sizeof(info)
    info.fMask = 0x00000040  # 프로세스 핸들을 돌려받아 종료를 기다린다
    info.lpVerb = "runas"
    info.lpFile = sys.executable
    waiting = "%d,%d" % (os.getpid(), os.getppid())
    info.lpParameters = subprocess.list2cmdline(arguments + ["--waiting-pids", waiting])
    info.nShow = 1
    if not ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(info)):
        return 1
    kernel32 = ctypes.windll.kernel32
    kernel32.WaitForSingleObject(info.hProcess, 0xFFFFFFFF)
    code = wintypes.DWORD()
    kernel32.GetExitCodeProcess(info.hProcess, ctypes.byref(code))
    kernel32.CloseHandle(info.hProcess)
    return int(code.value)


# ---------------------------------------------------------------- 설정·설치 화면

COLOR_HEADER = "#10202B"
COLOR_HEADER_SUB = "#9FB3C1"
COLOR_SURFACE = "#FFFFFF"
COLOR_SURFACE_2 = "#F5F7F9"
COLOR_INK = "#10202B"
COLOR_MUTED = "#4E6072"
COLOR_LINE = "#D7DFE5"
COLOR_PRIMARY = "#0B62A0"
COLOR_PRIMARY_STRONG = "#084C7D"
COLOR_TINT = "#EEF5FB"
COLOR_OK = "#1B6E3C"
COLOR_ERROR = "#B3261E"
FONT_FAMILY = "맑은 고딕"
FONT_MONO = "Consolas"

MB_OK, MB_YESNO, MB_ICONERROR, MB_ICONQUESTION, MB_ICONINFORMATION, IDYES = 0x0, 0x4, 0x10, 0x20, 0x40, 6


def message_box(text: str, flags: int = MB_OK | MB_ICONINFORMATION) -> int:
    return ctypes.windll.user32.MessageBoxW(None, text, APP_NAME, flags | 0x10000)


def _enable_dpi_awareness() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass


def _apply_window_icon(window) -> None:
    try:
        window.iconbitmap(default=str(resource_path("hihon.ico")))
    except Exception:
        pass


def is_running(tcp_port: int) -> bool:
    try:
        socket.create_connection(("127.0.0.1", tcp_port), timeout=1).close()
        return True
    except OSError:
        return False


def with_trailing_separator(folder: str) -> str:
    folder = folder.strip()
    return folder if not folder or folder.endswith("\\") else folder + "\\"


def append_subfolder(folder: str, token: str) -> str:
    """저장 폴더 맨 뒤에 항목을 그대로 붙인다. 구분 기호(\\)는 넣지 않는다."""
    return (folder or with_trailing_separator(DEFAULT_OUTPUT_FOLDER)) + token


def split_full_rule(text: str) -> Tuple[str, str, str]:
    """`C:\\폴더\\Cube_{datetime}.pdf` 같은 한 줄을 (저장 폴더, 앞 이름, 추가 이름)으로 나눈다."""
    text = text.strip()
    if text.lower().endswith(".pdf"):
        text = text[:-4]
    cut = text.rfind("\\")
    folder, name = text[:cut + 1], text[cut + 1:]
    first_item = name.find("{")
    if first_item < 0:
        return folder, name, ""
    return folder, name[:first_item], name[first_item:]


def save_settings(config_path: Path, settings: Dict[str, Any]) -> None:
    """설정 화면의 값을 확인하고 저장한다. 화면에 없는 값(포트 등)은 그대로 둔다."""
    save_mode = settings.get("save_mode")
    output_folder = str(settings.get("output_folder") or "").strip()
    prefix = str(settings.get("filename_prefix") or "")
    pattern = str(settings.get("filename_pattern") or "").strip()
    if save_mode not in SAVE_MODES:
        raise ValueError(tr("저장 방식을 골라 주세요."))
    if save_mode == "auto" or output_folder:
        root = static_folder_root(output_folder)
        if not re.match(r"^[A-Za-z]:\\", root + "\\"):
            raise ValueError(tr("저장 폴더는 C:\\ 처럼 드라이브부터 시작하는 전체 경로로 적어 주세요."))
        check_template(output_folder)
        os.makedirs(root, exist_ok=True)
    if re.search(r'[\\/:*?"<>|]', prefix):
        raise ValueError(tr("앞 이름에는 \\ / : * ? \" < > | 를 쓸 수 없습니다."))
    if not prefix.strip() and not pattern:
        raise ValueError(tr("앞 이름이나 추가 이름 중 하나는 있어야 합니다."))
    check_template(pattern)
    also_print = bool(settings.get("also_print"))
    printer = str(settings.get("also_print_printer") or "")
    if also_print and not printer:
        raise ValueError(tr("동시에 출력할 프린터를 골라 주세요."))

    raw = read_raw_config(config_path)
    raw["save_mode"] = save_mode
    if output_folder:
        raw["output_folder"] = output_folder
    raw["filename_prefix"] = prefix
    raw["filename_pattern"] = pattern
    quality = settings.get("image_quality")
    raw["image_quality"] = quality if quality in ("lossless", "compact") else "lossless"
    rotation = settings.get("page_rotation")
    raw["page_rotation"] = rotation if rotation in PAGE_ROTATIONS else "keep"
    color = settings.get("color_mode")
    raw["color_mode"] = color if color in COLOR_MODES else "keep"
    pdf_format = settings.get("pdf_format")
    raw["pdf_format"] = pdf_format if pdf_format in PDF_FORMATS else "pdf"
    raw["show_progress"] = bool(settings.get("show_progress"))
    raw["open_after_save"] = bool(settings.get("open_after_save"))
    raw["also_print"] = also_print
    raw["also_print_printer"] = printer
    if settings.get("language") in LANGUAGE_CODES:
        raw["language"] = settings["language"]
    config_path.parent.mkdir(parents=True, exist_ok=True)
    config_path.write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")


def build_settings_window(config_path: Path, install_mode: bool = False):
    """설정 창(설치할 때는 설치 창)을 만들어 돌려준다.

    시험에서 창을 띄우지 않고 다룰 수 있게 실행(mainloop)과 분리했다.
    """
    import tkinter as tk
    from tkinter import filedialog, ttk
    from tkinter import font as tkfont

    _enable_dpi_awareness()
    raw = read_raw_config(config_path)
    config = Config(raw)
    root = tk.Tk()
    root.title("%s %s" % (APP_NAME, "Setup" if install_mode else "Setting"))
    root.resizable(False, False)
    root.configure(bg=COLOR_SURFACE)
    _apply_window_icon(root)
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
        tkfont.nametofont(name).configure(family=FONT_FAMILY, size=9)
    root.option_add("*TCombobox*Listbox.font", (FONT_FAMILY, 9))
    ttk.Style(root).configure("Plain.TCheckbutton", background=COLOR_SURFACE, font=(FONT_FAMILY, 9))
    scale = max(1.0, root.winfo_fpixels("1i") / 96.0)

    def px(value: int) -> int:
        return int(round(value * scale))

    body_font = (FONT_FAMILY, 9)
    strong_font = (FONT_FAMILY, 9, "bold")

    # ------------------------------------------------ 부품

    def make_entry(parent, variable, mono: bool = False):
        """1px 테두리 입력 칸. 입력 중에는 테두리가 강조색이 된다."""
        outer = tk.Frame(parent, bg=COLOR_LINE, padx=1, pady=1)
        inner = tk.Frame(outer, bg=COLOR_SURFACE)
        inner.pack(fill="both", expand=True)
        entry = tk.Entry(
            inner, textvariable=variable, relief="flat", bd=0, bg=COLOR_SURFACE, fg=COLOR_INK,
            disabledbackground=COLOR_SURFACE_2, disabledforeground=COLOR_MUTED, insertbackground=COLOR_INK,
            font=(FONT_MONO, 10) if mono else body_font,
        )
        entry.pack(fill="x", padx=px(9), ipady=px(6))
        entry.bind("<FocusIn>", lambda _e: outer.configure(bg=COLOR_PRIMARY), add="+")
        entry.bind("<FocusOut>", lambda _e: outer.configure(bg=COLOR_LINE), add="+")

        def set_enabled(enabled: bool):
            entry.configure(state="normal" if enabled else "disabled")
            inner.configure(bg=COLOR_SURFACE if enabled else COLOR_SURFACE_2)

        outer.entry, outer.set_enabled = entry, set_enabled
        return outer

    def make_button(parent, text: str, command, primary: bool = False, width: int = 11):
        """크기가 같은 평면 버튼. 주요 동작은 채움, 나머지는 테두리만. 글이 긴 버튼은 width를 넓힌다."""
        fill = COLOR_PRIMARY if primary else COLOR_SURFACE
        hover = COLOR_PRIMARY_STRONG if primary else COLOR_SURFACE_2
        outer = tk.Frame(parent, bg=COLOR_PRIMARY if primary else COLOR_LINE, padx=1, pady=1)
        face = tk.Label(outer, text=text, bg=fill, fg="#FFFFFF" if primary else COLOR_INK, font=body_font,
                        width=width, pady=px(7), cursor="hand2")
        face.pack()
        state = {"enabled": True, "command": command}

        def set_enabled(enabled: bool):
            state["enabled"] = enabled
            face.configure(cursor="hand2" if enabled else "arrow",
                           fg=("#FFFFFF" if primary else COLOR_INK) if enabled else ("#BFD6E6" if primary else "#A9B5BF"))

        face.bind("<Enter>", lambda _e: state["enabled"] and face.configure(bg=hover))
        face.bind("<Leave>", lambda _e: face.configure(bg=fill))
        face.bind("<Button-1>", lambda _e: state["enabled"] and state["command"]())
        outer.set_enabled, outer.state, outer.face = set_enabled, state, face
        return outer

    def make_option_cards(parent, variable, options):
        """설명이 붙은 선택 카드. options: (값, 제목, 설명)."""
        holder = tk.Frame(parent, bg=COLOR_SURFACE)
        cards = []

        def paint(*_):
            for value, outer, parts, dot in cards:
                selected = variable.get() == value
                outer.configure(bg=COLOR_PRIMARY if selected else COLOR_LINE)
                for part in parts:
                    part.configure(bg=COLOR_TINT if selected else COLOR_SURFACE)
                dot.configure(text="●" if selected else "○", fg=COLOR_PRIMARY if selected else "#9AA8B4")

        for index, (value, title, description) in enumerate(options):
            outer = tk.Frame(holder, bg=COLOR_LINE, padx=1, pady=1, cursor="hand2")
            outer.grid(row=0, column=index, sticky="nsew", padx=(0 if index == 0 else px(10), 0))
            holder.columnconfigure(index, weight=1, uniform="cards")
            inner = tk.Frame(outer, padx=px(12), pady=px(10), cursor="hand2")
            inner.pack(fill="both", expand=True)
            dot = tk.Label(inner, font=(FONT_FAMILY, 10), cursor="hand2")
            dot.grid(row=0, column=0, sticky="n", padx=(0, px(8)))
            head = tk.Label(inner, text=title, font=strong_font, fg=COLOR_INK, anchor="w", cursor="hand2")
            head.grid(row=0, column=1, sticky="w")
            note = tk.Label(inner, text=description, fg=COLOR_MUTED, font=body_font, justify="left",
                            anchor="w", wraplength=px(210), cursor="hand2")
            note.grid(row=1, column=1, sticky="w", pady=(px(2), 0))
            for part in (outer, inner, dot, head, note):
                part.bind("<Button-1>", lambda _e, chosen=value: variable.set(chosen))
            cards.append((value, outer, (inner, dot, head, note), dot))
        variable.trace_add("write", paint)
        paint()
        return holder

    def make_dropdown(parent, caption, labels, on_pick, width: int = 22):
        """입력 칸과 테두리·높이가 같은 목록 단추. 누르면 바로 아래로 목록이 펼쳐진다."""
        outer = tk.Frame(parent, bg=COLOR_LINE, padx=1, pady=1)
        face = tk.Frame(outer, bg=COLOR_SURFACE, cursor="hand2")
        face.pack(fill="both", expand=True)
        text = tk.Label(face, textvariable=caption, bg=COLOR_SURFACE, fg=COLOR_INK, font=body_font, anchor="w",
                        width=width, cursor="hand2")
        text.pack(side="left", padx=(px(9), 0), pady=px(6))
        arrow = tk.Label(face, text="▾", bg=COLOR_SURFACE, fg=COLOR_MUTED, font=body_font, cursor="hand2")
        arrow.pack(side="right", padx=px(8))
        menu = tk.Menu(outer, tearoff=0, font=body_font, bg=COLOR_SURFACE, fg=COLOR_INK, bd=0,
                       activebackground=COLOR_TINT, activeforeground=COLOR_INK)
        for index, label in enumerate(labels):
            menu.add_command(label=label, command=lambda chosen=index: on_pick(chosen))
        state = {"enabled": True}

        def open_menu(_event=None):
            if state["enabled"] and labels:
                menu.tk_popup(outer.winfo_rootx(), outer.winfo_rooty() + outer.winfo_height())

        def set_enabled(enabled: bool):
            state["enabled"] = enabled
            for part in (face, text, arrow):
                part.configure(bg=COLOR_SURFACE if enabled else COLOR_SURFACE_2, cursor="hand2" if enabled else "arrow")
            text.configure(fg=COLOR_INK if enabled else COLOR_MUTED)

        for part in (face, text, arrow):
            part.bind("<Button-1>", open_menu)
        outer.set_enabled = set_enabled
        return outer

    def make_picker(parent, placeholder: str, on_pick):
        """항목을 고르면 on_pick(항목)을 부르는 목록. 단추에는 늘 안내 문구가 보인다."""
        return make_dropdown(parent, tk.StringVar(value=placeholder),
                             ["%s   %s" % (tr(label), token) for token, label in TOKENS],
                             lambda index: on_pick(TOKENS[index][0]))

    def make_choice(parent, variable, options, width: int = 30):
        """값 하나를 고르는 목록. options: (값, 보여 줄 말). 단추에는 지금 고른 값이 보인다."""
        caption = tk.StringVar()

        def show(*_):
            current = [label for value, label in options if value == variable.get()]
            caption.set(current[0] if current else (options[0][1] if options else ""))

        variable.trace_add("write", show)
        show()
        return make_dropdown(parent, caption, [label for _value, label in options],
                             lambda index: variable.set(options[index][0]), width=width)

    def make_check(parent, text: str, variable):
        return ttk.Checkbutton(parent, text=text, variable=variable, style="Plain.TCheckbutton", cursor="hand2")

    def field_label(parent, text: str, row: int, top: int = 0):
        tk.Label(parent, text=text, bg=COLOR_SURFACE, fg=COLOR_MUTED, font=body_font, anchor="w").grid(
            row=row, column=0, sticky="nw", pady=(px(top + 7), 0))

    def note_label(parent, text: str, row: int):
        tk.Label(parent, text=text, bg=COLOR_SURFACE, fg=COLOR_MUTED, font=body_font, justify="left",
                 anchor="w", wraplength=px(480)).grid(row=row, column=1, sticky="w", pady=(px(6), 0))

    # ------------------------------------------------ 제품명 띠

    header = tk.Frame(root, bg=COLOR_HEADER)
    header.pack(fill="x")
    try:
        logo_name = "logo-64.png" if scale >= 1.75 else "logo-48.png" if scale >= 1.25 else "logo-32.png"
        root.logo_image = tk.PhotoImage(file=str(resource_path(logo_name)))
        tk.Label(header, image=root.logo_image, bg=COLOR_HEADER).pack(side="left", padx=(px(20), px(12)), pady=px(14))
    except Exception:
        tk.Frame(header, bg=COLOR_HEADER, width=px(20), height=px(60)).pack(side="left")
    tk.Label(header, text=APP_NAME, bg=COLOR_HEADER, fg="#FFFFFF", font=(FONT_FAMILY, 13, "bold")).pack(side="left")
    status_var = tk.StringVar()
    status_label = tk.Label(header, textvariable=status_var, bg=COLOR_HEADER, fg=COLOR_HEADER_SUB, font=body_font)
    status_label.pack(side="right", padx=px(20))

    # ------------------------------------------------ 가로 메뉴

    menu = tk.Frame(root, bg=COLOR_SURFACE)
    menu.pack(fill="x", padx=px(20))
    tk.Frame(root, bg=COLOR_LINE, height=1).pack(fill="x")
    pages = tk.Frame(root, bg=COLOR_SURFACE)
    pages.pack(fill="both", expand=True, padx=px(24), pady=(px(20), px(16)))
    pages.columnconfigure(0, weight=1, minsize=px(600))
    tabs: Dict[str, Any] = {}

    def show_page(name: str):
        for key, (label, underline, page) in tabs.items():
            active = key == name
            label.configure(fg=COLOR_INK if active else COLOR_MUTED, font=strong_font if active else body_font)
            underline.configure(bg=COLOR_PRIMARY if active else COLOR_SURFACE)
        tabs[name][2].tkraise()
        root.current_page = name

    def add_page(name: str):
        holder = tk.Frame(menu, bg=COLOR_SURFACE, cursor="hand2")
        holder.pack(side="left", padx=(0, px(22)))
        label = tk.Label(holder, text=tr(name), bg=COLOR_SURFACE, cursor="hand2", pady=px(10))
        label.pack()
        underline = tk.Frame(holder, height=px(2), bg=COLOR_SURFACE)
        underline.pack(fill="x")
        page = tk.Frame(pages, bg=COLOR_SURFACE)
        page.grid(row=0, column=0, sticky="nsew")
        page.columnconfigure(0, minsize=px(92))
        page.columnconfigure(1, weight=1)
        for part in (holder, label):
            part.bind("<Button-1>", lambda _e: show_page(name))
        tabs[name] = (label, underline, page)
        return page

    # 설치 직후에는 매번 묻는 방식이 무난하다. 이미 쓰던 설정이 있으면 그대로 둔다.
    initial_mode = config.save_mode if "save_mode" in raw or not install_mode else "ask"
    mode_var = tk.StringVar(value=initial_mode)
    # 하위 폴더 규칙이 없는 기본 폴더는 끝에 \\를 붙여 보여 준다. 고른 항목이 그 뒤에 바로 이어진다.
    shown_folder = config.output_folder if "{" in config.output_folder else with_trailing_separator(config.output_folder)
    folder_var = tk.StringVar(value=shown_folder)
    prefix_var = tk.StringVar(value=config.filename_prefix)
    pattern_var = tk.StringVar(value=config.filename_pattern)
    quality_var = tk.StringVar(value=config.image_quality)
    rotation_var = tk.StringVar(value=config.page_rotation)
    color_var = tk.StringVar(value=config.color_mode)
    format_var = tk.StringVar(value=config.pdf_format)
    language_var = tk.StringVar(value=config.language)
    progress_var = tk.BooleanVar(value=config.show_progress)
    open_var = tk.BooleanVar(value=config.open_after_save)
    also_print_var = tk.BooleanVar(value=config.also_print)
    printers = list_printers()
    if config.also_print_printer and config.also_print_printer not in printers:
        printers.append(config.also_print_printer)
    printer_var = tk.StringVar(value=config.also_print_printer or (printers[0] if printers else ""))
    preview_title_var = tk.StringVar()
    preview_var = tk.StringVar()
    message_var = tk.StringVar()

    # ------------------------------------------------ 저장

    save_page = add_page("저장")
    field_label(save_page, tr("저장 방식"), 0)
    make_option_cards(save_page, mode_var, (
        ("ask", tr("매번 묻기"), tr("저장 폴더에서 저장 창이 열립니다.")),
        ("auto", tr("자동 저장"), tr("저장 폴더에 묻지 않고 바로 저장합니다.")),
    )).grid(row=0, column=1, sticky="we")

    field_label(save_page, tr("저장 폴더"), 1, top=18)
    folder_row = tk.Frame(save_page, bg=COLOR_SURFACE)
    folder_row.grid(row=1, column=1, sticky="we", pady=(px(18), 0))
    folder_row.columnconfigure(0, weight=1)
    folder_field = make_entry(folder_row, folder_var, mono=True)
    folder_field.grid(row=0, column=0, sticky="nsew")
    browse_button = make_button(folder_row, tr("찾아보기"), lambda: choose_folder())
    browse_button.grid(row=0, column=1, padx=(px(8), 0))

    def add_subfolder(token: str):
        folder_var.set(append_subfolder(folder_var.get(), token))
        folder_field.entry.icursor("end")
        folder_field.entry.xview_moveto(1.0)

    folder_picker = make_picker(save_page, tr("서브 폴더 생성"), add_subfolder)
    folder_picker.grid(row=2, column=1, sticky="w", pady=(px(8), 0))
    note_label(save_page, tr("고른 항목이 저장 폴더 뒤에 이어 붙고, 그 이름의 폴더가 인쇄할 때 자동으로 만들어집니다. "
                          "폴더를 한 단계 더 나누려면 사이에 \\ 를 직접 넣으세요."), 3)

    # ------------------------------------------------ 파일 이름

    name_page = add_page("파일 이름")
    field_label(name_page, tr("앞 이름"), 0)
    prefix_field = make_entry(name_page, prefix_var, mono=True)
    prefix_field.grid(row=0, column=1, sticky="w")
    prefix_field.entry.configure(width=24)
    field_label(name_page, tr("추가 이름"), 1, top=12)
    name_row = tk.Frame(name_page, bg=COLOR_SURFACE)
    name_row.grid(row=1, column=1, sticky="we", pady=(px(12), 0))
    name_row.columnconfigure(0, weight=1)
    pattern_field = make_entry(name_row, pattern_var, mono=True)
    pattern_field.grid(row=0, column=0, sticky="nsew")
    pattern_touched = {"value": False}
    pattern_field.entry.bind("<FocusIn>", lambda _e: pattern_touched.update(value=True), add="+")

    def add_name_token(token: str):
        # 한 번도 누르지 않은 칸의 커서는 맨 앞에 있어, 그때는 맨 뒤에 붙인다.
        entry = pattern_field.entry
        entry.insert(entry.index("insert") if pattern_touched["value"] else "end", token)
        entry.focus_set()
        pattern_touched["value"] = True

    make_picker(name_row, tr("추가 이름 생성"), add_name_token).grid(row=0, column=1, sticky="ns", padx=(px(8), 0))

    # 폴더·앞 이름·추가 이름을 한 줄로 이어 보여 주는 칸. 여기서 고쳐도 위 칸들에 그대로 반영된다.
    field_label(name_page, tr("전체"), 2, top=12)
    full_var = tk.StringVar()
    full_field = make_entry(name_page, full_var, mono=True)
    full_field.grid(row=2, column=1, sticky="we", pady=(px(12), 0))
    syncing = {"on": False}

    def compose_full(*_):
        if syncing["on"]:
            return
        syncing["on"] = True
        full_var.set(with_trailing_separator(folder_var.get()) + prefix_var.get() + pattern_var.get() + ".pdf")
        syncing["on"] = False

    def split_full(*_):
        if syncing["on"]:
            return
        syncing["on"] = True
        folder, prefix, pattern = split_full_rule(full_var.get())
        if folder:
            folder_var.set(folder)
        prefix_var.set(prefix)
        pattern_var.set(pattern)
        syncing["on"] = False

    for part in (folder_var, prefix_var, pattern_var):
        part.trace_add("write", compose_full)
    full_var.trace_add("write", split_full)
    compose_full()
    note_label(name_page, tr("위 세 칸 중 어디를 고쳐도 서로 맞춰집니다. "
                          "같은 이름의 파일이 이미 있으면 뒤에 _2, _3이 붙습니다."), 3)

    # ------------------------------------------------ PDF

    pdf_page = add_page("품질")
    field_label(pdf_page, tr("형식"), 0)
    make_choice(pdf_page, format_var, (
        ("pdf", "PDF"),
        ("pdfa1", "PDF/A-1b"),
        ("pdfa2", tr("PDF/A-2b (권장)")),
        ("pdfa3", "PDF/A-3b"),
    )).grid(row=0, column=1, sticky="w")
    field_label(pdf_page, tr("그림 품질"), 1, top=14)
    make_option_cards(pdf_page, quality_var, (
        ("lossless", tr("원본 그대로"), tr("그림을 손대지 않습니다.")),
        ("compact", tr("파일 크기 줄이기"), tr("그림을 압축합니다. 조금 흐려질 수 있습니다.")),
    )).grid(row=1, column=1, sticky="we", pady=(px(14), 0))
    field_label(pdf_page, tr("페이지 방향"), 2, top=14)
    make_choice(pdf_page, rotation_var, (
        ("keep", tr("인쇄한 그대로")),
        ("auto", tr("글자 방향에 맞춰 자동으로 돌리기")),
    )).grid(row=2, column=1, sticky="w", pady=(px(14), 0))
    field_label(pdf_page, tr("색상"), 3, top=10)
    make_choice(pdf_page, color_var, (
        ("keep", tr("인쇄한 그대로 (컬러)")),
        ("gray", tr("흑백")),
    )).grid(row=3, column=1, sticky="w", pady=(px(10), 0))
    note_label(pdf_page, tr("세로·가로 용지 방향은 인쇄하는 프로그램의 인쇄 창에서 고릅니다."), 4)

    # ------------------------------------------------ 출력

    after_page = add_page("출력")
    field_label(after_page, tr("할 일"), 0)
    checks = tk.Frame(after_page, bg=COLOR_SURFACE)
    checks.grid(row=0, column=1, sticky="w")
    make_check(checks, tr("변환 과정 보여주기"), progress_var).grid(row=0, column=0, columnspan=2, sticky="w", pady=(px(4), 0))
    make_check(checks, tr("생성 후 PDF 열기"), open_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=(px(8), 0))
    make_check(checks, tr("동시 출력"), also_print_var).grid(row=2, column=0, sticky="w", pady=(px(8), 0))
    printer_choice = make_choice(checks, printer_var, [(name, name) for name in printers] or [("", tr("쓸 수 있는 프린터가 없습니다"))], width=34)
    printer_choice.grid(row=2, column=1, sticky="w", padx=(px(10), 0), pady=(px(8), 0))
    note_label(after_page, tr("변환 과정은 화면 가운데에 진행 바로 잠깐 표시됩니다. "
                           "동시 출력은 PDF를 만들면서 고른 프린터로도 인쇄합니다."), 1)

    # ------------------------------------------------ 정보

    info_page = add_page("정보")
    facts = (
        (tr("제품"), "%s %s" % (APP_NAME, APP_VERSION)),
        (tr("제작"), tr("주식회사 히온 (Hihon Inc.)")),
        (tr("프린터 이름"), config.printer_name),
        (tr("라이선스"), "GNU AGPL v3"),
    )
    for index, (label, value) in enumerate(facts):
        tk.Label(info_page, text=label, bg=COLOR_SURFACE, fg=COLOR_MUTED, font=body_font, anchor="w").grid(
            row=index, column=0, sticky="w", pady=px(4))
        tk.Label(info_page, text=value, bg=COLOR_SURFACE, fg=COLOR_INK, font=body_font, anchor="w").grid(
            row=index, column=1, sticky="w", pady=px(4))

    def open_folder(path: Path):
        try:
            path.mkdir(parents=True, exist_ok=True)
            os.startfile(str(path))
        except OSError as exc:
            show_message(str(exc), error=True)

    field_label(info_page, tr("언어"), len(facts), top=8)
    make_choice(info_page, language_var, LANGUAGES, width=16).grid(
        row=len(facts), column=1, sticky="w", pady=(px(8), 0))
    tools = tk.Frame(info_page, bg=COLOR_SURFACE)
    tools.grid(row=len(facts) + 1, column=1, sticky="w", pady=(px(14), 0))
    make_button(tools, tr("저장 폴더 열기"), lambda: open_folder(Path(static_folder_root(folder_var.get()))),
                width=20).pack(side="left")
    make_button(tools, tr("실패한 인쇄물 보기"), lambda: open_folder(config.failed_dir),
                width=20).pack(side="left", padx=(px(8), 0))

    # ------------------------------------------------ 다음 인쇄물 미리보기

    preview = tk.Frame(root, bg=COLOR_SURFACE_2)
    preview.pack(fill="x")
    tk.Frame(preview, bg=COLOR_LINE, height=1).pack(fill="x")
    tk.Label(preview, textvariable=preview_title_var, bg=COLOR_SURFACE_2, fg=COLOR_MUTED, font=body_font,
             anchor="w").pack(fill="x", padx=px(24), pady=(px(10), 0))
    tk.Label(preview, textvariable=preview_var, bg=COLOR_SURFACE_2, fg=COLOR_INK, font=(FONT_MONO, 10),
             anchor="w", justify="left", wraplength=px(600)).pack(fill="x", padx=px(24), pady=(px(2), px(11)))
    tk.Frame(preview, bg=COLOR_LINE, height=1).pack(fill="x")

    # ------------------------------------------------ 알림과 버튼

    footer = tk.Frame(root, bg=COLOR_SURFACE)
    footer.pack(fill="x", padx=px(24), pady=px(14))
    close_button = make_button(footer, tr("닫기"), root.destroy)
    close_button.pack(side="right")
    primary_button = make_button(footer, tr("설치") if install_mode else tr("저장"), lambda: None, primary=True)
    primary_button.pack(side="right", padx=(0, px(8)))
    message_label = tk.Label(footer, textvariable=message_var, bg=COLOR_SURFACE, fg=COLOR_OK, font=body_font,
                             anchor="w", justify="left", wraplength=px(360))
    message_label.pack(side="left", fill="x", expand=True)

    def show_message(text: str, error: bool = False):
        message_var.set(text)
        message_label.configure(fg=COLOR_ERROR if error else COLOR_OK)

    def refresh(*_):
        auto = mode_var.get() == "auto"
        printer_choice.set_enabled(bool(also_print_var.get() and printers))
        values = sample_values()
        rule = prefix_var.get().replace("{", "{{").replace("}", "}}") + pattern_var.get()
        name = build_filename(rule, values) + ".pdf"
        preview_title_var.set(tr("다음 인쇄물이 저장될 곳") if auto else tr("저장 창이 처음 열릴 폴더와 이름"))
        preview_var.set(str(render_folder(folder_var.get() or DEFAULT_OUTPUT_FOLDER, values) / name))

    def refresh_status():
        if install_mode and not root.installed:
            status_var.set("")
        else:
            running = is_running(config.port)
            status_var.set(tr("●  동작 중") if running else tr("●  멈춰 있음"))
            status_label.configure(fg="#6FD08C" if running else "#E8A0A0")
        root.after(5000, refresh_status)

    def choose_folder():
        start = static_folder_root(folder_var.get()) or "C:\\"
        chosen = filedialog.askdirectory(parent=root, initialdir=start, title=tr("PDF를 저장할 폴더"), mustexist=False)
        if chosen:
            folder_var.set(with_trailing_separator(os.path.normpath(chosen)))

    def collect() -> Dict[str, Any]:
        return {
            "save_mode": mode_var.get(), "output_folder": folder_var.get(),
            "filename_prefix": prefix_var.get(), "filename_pattern": pattern_var.get(),
            "image_quality": quality_var.get(), "page_rotation": rotation_var.get(), "color_mode": color_var.get(),
            "pdf_format": format_var.get(),
            "show_progress": progress_var.get(), "open_after_save": open_var.get(),
            "also_print": also_print_var.get(), "also_print_printer": printer_var.get(),
            "language": language_var.get(),
        }

    def save() -> bool:
        try:
            save_settings(config_path, collect())
        except (ValueError, OSError) as exc:
            show_message(str(exc), error=True)
            return False
        show_message(tr("저장했습니다. 다음 인쇄부터 적용됩니다."))
        if language_var.get() != config.language:
            # 언어를 바꿨으면 창을 닫고, 부른 쪽에서 새 언어로 다시 연다.
            root.reopen = True
            root.destroy()
        return True

    def run_install():
        if not save():
            return
        primary_button.set_enabled(False)
        close_button.set_enabled(False)
        show_message(tr("설치하는 중입니다. 잠시만 기다려 주세요."))
        outcome: Dict[str, Any] = {}

        def work():
            try:
                outcome["config"] = install(config_path)
            except (SetupError, OSError) as exc:
                outcome["error"] = str(exc)

        thread = threading.Thread(target=work, daemon=True)
        thread.start()

        def wait():
            if thread.is_alive():
                root.after(300, wait)
                return
            close_button.set_enabled(True)
            if "error" in outcome:
                primary_button.set_enabled(True)
                show_message(tr("설치하지 못했습니다. %s") % outcome["error"], error=True)
                return
            root.installed = True
            primary_button.pack_forget()
            show_message(tr("설치가 끝났습니다. 프린터를 \"%s\"로 골라 인쇄해 보세요.") % outcome["config"].printer_name)

        wait()

    primary_button.state["command"] = run_install if install_mode else save
    root.bind("<Escape>", lambda _e: close_button.state["enabled"] and root.destroy())
    for variable in (mode_var, folder_var, prefix_var, pattern_var, also_print_var):
        variable.trace_add("write", refresh)
    root.installed = False
    root.reopen = False
    show_page("저장")
    refresh()
    refresh_status()
    root.save_settings = save
    root.show_page = show_page
    root.add_subfolder = add_subfolder
    root.add_name_token = add_name_token
    root.variables = {
        "mode": mode_var, "folder": folder_var, "prefix": prefix_var, "pattern": pattern_var, "full": full_var,
        "quality": quality_var, "rotation": rotation_var, "color": color_var, "format": format_var,
        "progress": progress_var,
        "open": open_var, "also_print": also_print_var, "printer": printer_var,
        "language": language_var,
        "preview": preview_var, "preview_title": preview_title_var, "message": message_var,
    }
    return root


# ---------------------------------------------------------------- 설치 마법사

SETUP_STEPS = ("welcome", "license", "folder", "progress", "done")


def read_license_text() -> str:
    for folder in (app_dir(), Path(getattr(sys, "_MEIPASS", app_dir()))):
        try:
            return (folder / LICENSE_FILES[0]).read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
    return "GNU Affero General Public License v3\nhttps://www.gnu.org/licenses/agpl-3.0.html"


def build_setup_window(config_path: Path, language: Optional[str] = None):
    """설치 마법사 창을 만들어 돌려준다: 시작 → 라이선스 → 설치 폴더 → 설치 진행 → 완료."""
    import tkinter as tk
    from tkinter import filedialog, ttk
    from tkinter import font as tkfont

    _enable_dpi_awareness()
    raw = read_raw_config(config_path)
    chosen_language = language if language in LANGUAGE_CODES else Config(raw).language
    set_language(chosen_language)
    already_installed = (installed_dir() / EXE_NAME).exists()
    root = tk.Tk()
    root.title("%s Setup" % APP_NAME)
    root.resizable(False, False)
    root.configure(bg=COLOR_SURFACE)
    _apply_window_icon(root)
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
        tkfont.nametofont(name).configure(family=FONT_FAMILY, size=9)
    ttk.Style(root).configure("Plain.TCheckbutton", background=COLOR_SURFACE, font=(FONT_FAMILY, 9))
    scale = max(1.0, root.winfo_fpixels("1i") / 96.0)

    def px(value: int) -> int:
        return int(round(value * scale))

    body_font = (FONT_FAMILY, 9)

    def make_button(parent, text: str, command, primary: bool = False):
        fill = COLOR_PRIMARY if primary else COLOR_SURFACE
        hover = COLOR_PRIMARY_STRONG if primary else COLOR_SURFACE_2
        outer = tk.Frame(parent, bg=COLOR_PRIMARY if primary else COLOR_LINE, padx=1, pady=1)
        face = tk.Label(outer, text=text, bg=fill, fg="#FFFFFF" if primary else COLOR_INK, font=body_font,
                        width=11, pady=px(7), cursor="hand2")
        face.pack()
        state = {"enabled": True}

        def set_enabled(enabled: bool):
            state["enabled"] = enabled
            face.configure(cursor="hand2" if enabled else "arrow",
                           fg=("#FFFFFF" if primary else COLOR_INK) if enabled else ("#BFD6E6" if primary else "#A9B5BF"))

        face.bind("<Enter>", lambda _e: state["enabled"] and face.configure(bg=hover))
        face.bind("<Leave>", lambda _e: face.configure(bg=fill))
        face.bind("<Button-1>", lambda _e: state["enabled"] and command())
        outer.set_enabled, outer.face, outer.state = set_enabled, face, state
        return outer

    # --- 제품명 띠
    header = tk.Frame(root, bg=COLOR_HEADER)
    header.pack(fill="x")
    try:
        logo_name = "logo-64.png" if scale >= 1.75 else "logo-48.png" if scale >= 1.25 else "logo-32.png"
        root.logo_image = tk.PhotoImage(file=str(resource_path(logo_name)))
        tk.Label(header, image=root.logo_image, bg=COLOR_HEADER).pack(side="left", padx=(px(20), px(12)), pady=px(14))
    except Exception:
        tk.Frame(header, bg=COLOR_HEADER, width=px(20), height=px(60)).pack(side="left")
    tk.Label(header, text=APP_NAME, bg=COLOR_HEADER, fg="#FFFFFF", font=(FONT_FAMILY, 13, "bold")).pack(side="left")
    tk.Label(header, text="Setup", bg=COLOR_HEADER, fg=COLOR_HEADER_SUB, font=body_font).pack(side="right", padx=px(20))

    pages_holder = tk.Frame(root, bg=COLOR_SURFACE)
    pages_holder.pack(fill="both", expand=True, padx=px(28), pady=(px(22), px(12)))
    pages_holder.columnconfigure(0, weight=1, minsize=px(560))
    pages_holder.rowconfigure(0, weight=1, minsize=px(290))
    pages: Dict[str, Any] = {}

    def add_page(name: str, title: str):
        page = tk.Frame(pages_holder, bg=COLOR_SURFACE)
        page.grid(row=0, column=0, sticky="nsew")
        tk.Label(page, text=title, bg=COLOR_SURFACE, fg=COLOR_INK, font=(FONT_FAMILY, 12, "bold"), anchor="w").pack(
            fill="x", pady=(0, px(10)))
        pages[name] = page
        return page

    def paragraph(parent, text: str, muted: bool = False, variable=None):
        label = tk.Label(parent, text=text, textvariable=variable, bg=COLOR_SURFACE,
                         fg=COLOR_MUTED if muted else COLOR_INK, font=body_font, justify="left", anchor="w",
                         wraplength=px(550))
        label.pack(fill="x", pady=(0, px(8)))
        return label

    agree_var = tk.BooleanVar(value=False)
    folder_var = tk.StringVar(value=str(installed_dir()))
    open_settings_var = tk.BooleanVar(value=True)
    status_var = tk.StringVar()
    message_var = tk.StringVar()

    # --- 시작
    page = add_page("welcome", tr("%s 설치를 시작합니다") % APP_NAME)
    paragraph(page, tr("인쇄할 때 프린터를 \"%s\"로 고르면 PDF 파일이 만들어집니다.") % APP_NAME)
    paragraph(page, tr("버전 %s\n제작 주식회사 히온 (Hihon Inc.)") % APP_VERSION, muted=True)
    if already_installed:
        paragraph(page, tr("이 PC에 이미 설치되어 있습니다. 계속하면 새 파일로 바꾸고, 쓰던 설정은 그대로 둡니다."))
    paragraph(page, tr("계속하려면 \"다음\"을 누르세요."))
    language_row = tk.Frame(page, bg=COLOR_SURFACE)
    language_row.pack(anchor="w", pady=(px(16), 0))
    tk.Label(language_row, text=tr("언어"), bg=COLOR_SURFACE, fg=COLOR_MUTED, font=body_font).pack(
        side="left", padx=(0, px(10)))
    language_box = ttk.Combobox(language_row, values=[label for _code, label in LANGUAGES], state="readonly",
                                width=14, font=body_font)
    language_box.current(LANGUAGE_CODES.index(chosen_language))
    language_box.pack(side="left")

    def change_language(_event=None):
        picked = LANGUAGE_CODES[language_box.current()]
        if picked != chosen_language:
            # 창을 닫고, 부른 쪽에서 고른 언어로 처음부터 다시 연다.
            root.language, root.reopen = picked, True
            root.destroy()

    language_box.bind("<<ComboboxSelected>>", change_language)

    # --- 라이선스
    page = add_page("license", tr("라이선스"))
    paragraph(page, tr("이 프로그램은 GNU AGPL v3 조건으로 제공됩니다. 함께 설치되는 다른 프로그램의 라이선스는 "
                    "설치 폴더의 THIRD-PARTY.txt에 있습니다."), muted=True)
    license_frame = tk.Frame(page, bg=COLOR_LINE, padx=1, pady=1)
    license_frame.pack(fill="both", expand=True)
    scrollbar = ttk.Scrollbar(license_frame, orient="vertical")
    license_text = tk.Text(license_frame, height=10, wrap="word", relief="flat", bd=0, bg=COLOR_SURFACE_2,
                           fg=COLOR_INK, font=(FONT_MONO, 9), padx=px(10), pady=px(8), yscrollcommand=scrollbar.set)
    scrollbar.configure(command=license_text.yview)
    scrollbar.pack(side="right", fill="y")
    license_text.pack(side="left", fill="both", expand=True)
    license_text.insert("1.0", read_license_text())
    license_text.configure(state="disabled")
    ttk.Checkbutton(page, text=tr("위 라이선스 내용에 동의합니다"), variable=agree_var, style="Plain.TCheckbutton",
                    cursor="hand2").pack(anchor="w", pady=(px(10), 0))

    # --- 설치 폴더
    page = add_page("folder", tr("설치 폴더"))
    paragraph(page, tr("프로그램 파일을 넣을 폴더입니다. 바꾸지 않아도 됩니다."))
    folder_row = tk.Frame(page, bg=COLOR_SURFACE)
    folder_row.pack(fill="x", pady=(px(4), px(10)))
    entry_border = tk.Frame(folder_row, bg=COLOR_LINE, padx=1, pady=1)
    entry_border.pack(side="left", fill="both", expand=True)
    entry_inner = tk.Frame(entry_border, bg=COLOR_SURFACE)
    entry_inner.pack(fill="both", expand=True)
    folder_entry = tk.Entry(entry_inner, textvariable=folder_var, relief="flat", bd=0, bg=COLOR_SURFACE,
                            fg=COLOR_INK, insertbackground=COLOR_INK, font=(FONT_MONO, 10))
    folder_entry.pack(fill="x", expand=True, padx=px(9), ipady=px(6))
    folder_entry.bind("<FocusIn>", lambda _e: entry_border.configure(bg=COLOR_PRIMARY))
    folder_entry.bind("<FocusOut>", lambda _e: entry_border.configure(bg=COLOR_LINE))

    def choose_folder():
        chosen = filedialog.askdirectory(parent=root, title=tr("설치할 폴더"), mustexist=False,
                                         initialdir=os.path.dirname(folder_var.get()) or "C:\\")
        if chosen:
            folder_var.set(normalize_install_folder(chosen))

    make_button(folder_row, tr("찾아보기"), choose_folder).pack(side="left", padx=(px(8), 0))
    paragraph(page, tr("필요한 공간: 약 35MB"), muted=True)
    paragraph(page, tr("만들어진 PDF가 저장되는 폴더는 설치 뒤 설정에서 정합니다."), muted=True)

    # --- 설치 진행
    page = add_page("progress", tr("설치하는 중입니다"))
    paragraph(page, "", variable=status_var)
    bar_width, bar_height = px(550), max(4, px(6))
    bar = tk.Canvas(page, width=bar_width, height=bar_height, bg=COLOR_SURFACE_2, highlightthickness=0, bd=0)
    bar.pack(anchor="w", pady=(px(6), 0))
    runner = bar.create_rectangle(-bar_width // 3, 0, 0, bar_height, fill=COLOR_PRIMARY, width=0)
    motion = {"x": -bar_width // 3, "running": False}

    def animate():
        if not motion["running"]:
            return
        motion["x"] = -bar_width // 3 if motion["x"] > bar_width else motion["x"] + max(4, bar_width // 60)
        bar.coords(runner, motion["x"], 0, motion["x"] + bar_width // 3, bar_height)
        root.after(30, animate)

    # --- 완료
    page = add_page("done", tr("설치가 끝났습니다"))
    paragraph(page, tr("프린터 목록에 \"%s\"가 추가됐습니다. 아무 프로그램에서나 이 프린터로 인쇄해 보세요.") % APP_NAME)
    paragraph(page, tr("저장 폴더, 파일 이름, 품질은 시작 메뉴의 \"%s Setting\"에서 바꿉니다.") % APP_NAME, muted=True)
    ttk.Checkbutton(page, text=tr("지금 설정 열기"), variable=open_settings_var, style="Plain.TCheckbutton",
                    cursor="hand2").pack(anchor="w", pady=(px(6), 0))

    # --- 알림과 버튼
    tk.Frame(root, bg=COLOR_LINE, height=1).pack(fill="x")
    footer = tk.Frame(root, bg=COLOR_SURFACE)
    footer.pack(fill="x", padx=px(24), pady=px(14))
    cancel_button = make_button(footer, tr("취소"), lambda: close())
    cancel_button.pack(side="right")
    next_button = make_button(footer, tr("다음"), lambda: go_next(), primary=True)
    next_button.pack(side="right", padx=(0, px(8)))
    back_button = make_button(footer, tr("이전"), lambda: go_back())
    back_button.pack(side="right", padx=(0, px(8)))
    message_label = tk.Label(footer, textvariable=message_var, bg=COLOR_SURFACE, fg=COLOR_ERROR, font=body_font,
                             anchor="w", justify="left", wraplength=px(260))
    message_label.pack(side="left", fill="x", expand=True)

    def show(step: str):
        root.step = step
        pages[step].tkraise()
        message_var.set("")
        back_button.set_enabled(step in ("license", "folder"))
        next_button.face.configure(text={"folder": tr("설치"), "done": tr("마침")}.get(step, tr("다음")))
        next_button.set_enabled(step != "progress")
        cancel_button.set_enabled(step not in ("progress", "done"))

    def close():
        if root.step != "progress":
            root.destroy()

    def go_back():
        if root.step in ("license", "folder"):
            show(SETUP_STEPS[SETUP_STEPS.index(root.step) - 1])

    def go_next():
        step = root.step
        if step == "welcome":
            show("license")
        elif step == "license":
            if not agree_var.get():
                message_var.set(tr("라이선스에 동의해야 설치할 수 있습니다."))
                return
            show("folder")
        elif step == "folder":
            try:
                chosen = check_install_folder(folder_var.get())
            except ValueError as exc:
                message_var.set(str(exc))
                return
            start_install(chosen)
        elif step == "done":
            if open_settings_var.get():
                exe = installed_dir() / EXE_NAME
                if exe.exists():
                    subprocess.Popen([str(exe), "--settings"], close_fds=True, creationflags=DETACHED_PROCESS)
            root.destroy()

    def start_install(chosen: Path):
        show("progress")
        status_var.set(tr("프린터를 등록하고 파일을 복사하고 있습니다. 10~20초쯤 걸립니다."))
        motion["running"] = True
        animate()
        outcome: Dict[str, Any] = {}

        def work():
            try:
                # 처음 설치할 때는 저장할 때마다 위치를 묻는 방식으로 시작한다. 쓰던 설정이 있으면 건드리지 않는다.
                outcome["config"] = install(config_path, save_mode=None if "save_mode" in raw else "ask",
                                            target_dir=str(chosen), language=chosen_language)
            except Exception as exc:  # 어떤 이유로든 실패하면 진행 화면에 갇히지 않고 사유를 보여 준다
                outcome["error"] = str(exc)

        thread = threading.Thread(target=work, daemon=True)
        thread.start()

        def wait():
            if thread.is_alive():
                root.after(300, wait)
                return
            motion["running"] = False
            if "error" in outcome:
                show("folder")
                message_var.set(tr("설치하지 못했습니다. %s") % outcome["error"])
                return
            root.installed = True
            show("done")

        wait()

    root.protocol("WM_DELETE_WINDOW", close)
    root.bind("<Escape>", lambda _e: close())
    root.installed = False
    root.reopen = False
    root.go_next, root.go_back = go_next, go_back
    root.language, root.change_language, root.language_box = chosen_language, change_language, language_box
    root.variables = {"agree": agree_var, "folder": folder_var, "open_settings": open_settings_var,
                      "message": message_var, "status": status_var}
    show("welcome")
    return root


# ---------------------------------------------------------------- 실행

SETTINGS_PAGES = ("저장", "파일 이름", "품질", "출력", "정보")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(prog=EXE_NAME, description=APP_NAME)
    parser.add_argument("--config", default=str(default_config_path()))
    parser.add_argument("--run", action="store_true", help="인쇄 데이터를 받아 PDF로 바꾼다 (기본 동작)")
    parser.add_argument("--agent", action="store_true", help="저장 위치를 묻는 창을 띄우는 도우미")
    parser.add_argument("--convert", metavar="FILE", help="PostScript 파일 하나를 PDF로 변환한다")
    parser.add_argument("--settings", action="store_true", help="설정 화면을 연다")
    parser.add_argument("--setup", action="store_true", help="설치 화면을 연다")
    parser.add_argument("--install", action="store_true", help="화면 없이 설치한다")
    parser.add_argument("--uninstall", action="store_true", help="제거한다")
    parser.add_argument("--output", help="PDF를 저장할 폴더 (--install과 함께)")
    parser.add_argument("--mode", choices=SAVE_MODES, help="저장 방식: auto 묻지 않고 저장, ask 매번 묻기 (--install과 함께)")
    parser.add_argument("--pattern", help="파일 이름 규칙 (--install과 함께)")
    parser.add_argument("--dir", help="설치할 폴더 (--install과 함께)")
    parser.add_argument("--language", choices=LANGUAGE_CODES, help="화면 언어 (--install과 함께)")
    parser.add_argument("--ui", action="store_true", help="결과를 알림창으로 보여 준다")
    parser.add_argument("--page", type=int, choices=range(1, len(SETTINGS_PAGES) + 1), help=argparse.SUPPRESS)
    parser.add_argument("--waiting-pids", default="", help=argparse.SUPPRESS)
    parser.add_argument("--version", action="version", version="%s %s" % (APP_NAME, APP_VERSION))
    # 창 없는 실행 파일로 빌드하면 표준 출력이 없어 argparse와 print가 실패한다.
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")
    args = parser.parse_args(argv)
    config_path = Path(args.config)
    arguments = list(sys.argv[1:] if argv is None else argv)
    KEEP_PIDS.update(int(pid) for pid in args.waiting_pids.split(",") if pid.strip().isdigit())
    try:
        set_language(load_config(config_path).language)
    except (OSError, ValueError):
        pass

    if args.agent:
        return run_agent(config_path)

    if args.settings:
        # 설정 파일이 모든 사용자 공용 위치에 있어 고치려면 관리자 권한이 필요하다.
        if config_path == default_config_path() and not is_admin():
            return relaunch_elevated(arguments)
        page = SETTINGS_PAGES[args.page - 1] if args.page else None
        while True:
            set_language(load_config(config_path).language)
            window = build_settings_window(config_path)
            if page:
                window.show_page(page)
            window.mainloop()
            if not window.reopen:
                return 0
            # 언어를 바꿔 다시 여는 경우에는 바꾸던 자리(정보)로 돌아온다.
            page = SETTINGS_PAGES[-1]

    if args.setup or args.install or args.uninstall:
        if os.name != "nt":
            print(tr("Windows에서만 설치할 수 있습니다."))
            return 1
        if not is_admin():
            code = relaunch_elevated(arguments)
            # 저장 창은 로그인한 사용자 권한으로 떠야 하므로, 권한을 올리기 전의 이쪽에서 도우미를 띄운다.
            if code == 0 and not args.uninstall:
                start_agent()
            return code
        setup_logging(data_dir() / "logs")
        if args.setup:
            language = None
            while True:
                window = build_setup_window(config_path, language)
                window.mainloop()
                if not window.reopen:
                    return 0 if window.installed else 1
                language = window.language
        try:
            if args.install:
                config = install(config_path, args.output, args.mode, args.pattern, args.dir, args.language)
                log.info("설치 완료 — 프린터 %s, 포트 %d, 저장 방식 %s, 폴더 %s",
                         config.printer_name, config.port, config.save_mode, config.output_folder)
                if args.ui:
                    message_box(tr("설치가 끝났습니다."))
            else:
                if args.ui and message_box(
                    tr("%s를 제거할까요?\n저장된 PDF와 설정은 지우지 않습니다.") % APP_NAME, MB_YESNO | MB_ICONQUESTION
                ) != IDYES:
                    return 1
                uninstall(config_path)
                log.info("제거 완료")
                if args.ui:
                    message_box(tr("제거했습니다."))
        except (SetupError, OSError, ValueError) as exc:
            log.error("%s", exc)
            if args.ui:
                message_box(str(exc), MB_OK | MB_ICONERROR)
            return 1
        return 0

    config = load_config(config_path)
    setup_logging(config.log_dir)
    if args.convert:
        config.spool_dir.mkdir(parents=True, exist_ok=True)
        job = config.spool_dir / (datetime.now().strftime("%Y%m%d_%H%M%S") + "_manual.job")
        shutil.copyfile(args.convert, str(job))
        trim_trailer(job)
        return 0 if Converter(config).convert(job) else 1

    server = PrintServer(config, config_path)
    try:
        server.serve()
    except KeyboardInterrupt:
        server.stop()
    except OSError as exc:
        log.error("포트 %d를 열지 못했습니다: %s", config.port, exc)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
