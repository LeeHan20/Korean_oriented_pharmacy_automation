"""
한약국 자동화 프로그램 - 공통 유틸리티 함수
"""

import os
import re
import glob
import time
import queue
import threading
import tkinter as tk
from tkinter import messagebox
from pathlib import Path
from datetime import datetime, timedelta

import config


# ─── Chrome / Selenium ───────────────────────────────────────────────────────

def connect_chrome():
    """
    원격 디버깅 모드로 실행 중인 Chrome에 연결합니다.
    '크롬_디버깅모드_실행.bat'으로 Chrome을 먼저 실행해야 합니다.
    """
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    opts = Options()
    opts.add_experimental_option(
        "debuggerAddress", f"localhost:{config.CHROME_DEBUG_PORT}"
    )
    try:
        driver = webdriver.Chrome(options=opts)
        driver.implicitly_wait(5)
        return driver
    except Exception as e:
        raise ConnectionError(
            f"Chrome 연결 실패.\n"
            f"'크롬_디버깅모드_실행.bat'으로 Chrome을 실행한 뒤 다시 시도하세요.\n"
            f"상세: {e}"
        )


# ─── 파일 다운로드 대기 ───────────────────────────────────────────────────────

def wait_for_new_file(directory: str, pattern: str, before_mtime: float,
                      timeout: int = None) -> str:
    """
    directory 안에서 pattern 에 맞고 before_mtime 이후에 생성된
    최신 파일이 다운로드 완료될 때까지 기다립니다.
    """
    if timeout is None:
        timeout = config.DOWNLOAD_TIMEOUT

    deadline = time.time() + timeout
    while time.time() < deadline:
        candidates = [
            f for f in glob.glob(os.path.join(directory, pattern))
            if os.path.getmtime(f) >= before_mtime
            and not f.endswith(".crdownload")
            and not f.endswith(".tmp")
        ]
        if candidates:
            return max(candidates, key=os.path.getmtime)
        time.sleep(0.5)
    raise TimeoutError(f"파일 다운로드 타임아웃 ({timeout}초). 패턴: {pattern}")


def get_latest_file(directory: str, pattern: str) -> str:
    """directory 에서 pattern 에 맞는 가장 최신 파일 경로 반환."""
    files = glob.glob(os.path.join(directory, pattern))
    if not files:
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {os.path.join(directory, pattern)}")
    return max(files, key=os.path.getmtime)


# ─── Excel 변환 ──────────────────────────────────────────────────────────────

def xls_to_xlsx(xls_path: str) -> str:
    """
    XLS 파일을 XLSX 로 안전하게 변환합니다 (win32com 사용).
    원본 XLS 파일은 그대로 유지됩니다.
    """
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()

    xls_path = os.path.abspath(xls_path)
    xlsx_path = os.path.splitext(xls_path)[0] + ".xlsx"

    xl_app = win32com.client.Dispatch("Excel.Application")
    xl_app.Visible = False
    xl_app.DisplayAlerts = False

    try:
        wb = xl_app.Workbooks.Open(xls_path)
        wb.SaveAs(xlsx_path, FileFormat=51)   # 51 = xlOpenXMLWorkbook
        wb.Close(SaveChanges=False)
    finally:
        xl_app.Quit()
        pythoncom.CoUninitialize()

    return xlsx_path


# ─── OKOSC 통합문서 Excel 찾기 ───────────────────────────────────────────────

def get_okosc_workbook(wait_new: bool = False, before_names: set = None):
    """
    현재 열려 있는 Excel 통합 문서 중 '통합 문서 N' 형식에서
    N이 가장 큰 워크북을 반환합니다.

    wait_new=True 이면 before_names 에 없는 새 통합문서가
    나타날 때까지 최대 OKOSC_WAIT_TIMEOUT 초 대기합니다.
    """
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    pattern = re.compile(r'^통합 문서\s*(\d+)$')
    deadline = time.time() + config.OKOSC_WAIT_TIMEOUT

    while True:
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            if time.time() < deadline:
                time.sleep(0.5)
                continue
            raise RuntimeError("Excel 응용 프로그램을 찾을 수 없습니다.")

        max_n = -1
        target_wb = None
        for wb in xl.Workbooks:
            m = pattern.match(wb.Name)
            if m:
                n = int(m.group(1))
                if wait_new and before_names and wb.Name in before_names:
                    continue
                if n > max_n:
                    max_n = n
                    target_wb = wb

        if target_wb is not None:
            return target_wb

        if not wait_new or time.time() >= deadline:
            raise RuntimeError(
                "'통합 문서 N' 형식의 Excel 창을 찾을 수 없습니다.\n"
                "OKOSC에서 택배목록을 먼저 열어주세요."
            )
        time.sleep(0.5)


def list_excel_workbook_names() -> set:
    """현재 열려 있는 모든 Excel 워크북 이름을 반환합니다."""
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        return {wb.Name for wb in xl.Workbooks}
    except Exception:
        return set()


# ─── 셀 색상 판별 ─────────────────────────────────────────────────────────────

def is_light_green(cell) -> bool:
    """
    openpyxl 셀의 배경색이 '연한 녹색'인지 판별합니다.
    RGB 기준: Green 채널이 Red/Blue 보다 크고 밝은 경우.
    """
    fill = cell.fill
    if fill is None or fill.fill_type in (None, "none"):
        return False
    try:
        fg = fill.fgColor
        if fg.type == "rgb":
            rgb = fg.rgb  # "AARRGGBB" 형식
            if len(rgb) == 8:
                r = int(rgb[2:4], 16)
                g = int(rgb[4:6], 16)
                b = int(rgb[6:8], 16)
                # 녹색 채널이 지배적이고 밝은 색
                return g > r and g > b and g > 100
        elif fg.type == "theme":
            # 테마 색상 9번(Accent 6)이 엑셀 기본 '연한 녹색'에 해당하는 경우가 많음
            return fg.theme == 9
    except Exception:
        pass
    return False


# ─── 익산대장 파일 찾기 ───────────────────────────────────────────────────────

def find_iksan_file(directory: str = None) -> str:
    """'익산대장'으로 시작하는 Excel 파일 경로를 반환합니다."""
    if directory is None:
        directory = config.IKSAN_FILE_DIR

    for ext in ("*.xlsx", "*.xls", "*.xlsm"):
        matches = glob.glob(os.path.join(directory, "익산대장*" + ext.lstrip("*")))
        if matches:
            return matches[0]
    raise FileNotFoundError(
        f"익산대장 파일을 찾을 수 없습니다.\n경로: {directory}"
    )


# ─── Excel 마지막 데이터 행 ───────────────────────────────────────────────────

def get_last_data_row(ws, check_cols=(1,)) -> int:
    """
    지정한 열(check_cols, 1-based) 중 하나라도 값이 있는
    가장 마지막 행 번호를 반환합니다. 데이터가 없으면 0.
    """
    for row in range(ws.max_row, 0, -1):
        if any(ws.cell(row=row, column=c).value not in (None, "") for c in check_cols):
            return row
    return 0


# ─── 날짜 유틸 ───────────────────────────────────────────────────────────────

def get_search_dates():
    """(시작일 문자열, 오늘 문자열) 반환. 형식: YYYY-MM-DD"""
    today = datetime.now()
    start = today - timedelta(days=config.DATE_RANGE_DAYS)
    return start.strftime("%Y-%m-%d"), today.strftime("%Y-%m-%d")


# ─── 인간 검토 다이얼로그 ─────────────────────────────────────────────────────

def human_review_dialog(title: str, message: str,
                        ok_text="계속 진행", cancel_text="중단") -> bool:
    """
    tkinter 다이얼로그를 메인 스레드에서 표시합니다.
    OK → True, Cancel → False
    """
    result_holder = [None]
    event = threading.Event()

    def _show():
        win = tk.Toplevel()
        win.title(title)
        win.grab_set()
        win.resizable(False, False)

        tk.Label(win, text=message, font=("맑은 고딕", 10),
                 wraplength=400, justify="left", padx=20, pady=20).pack()

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=(0, 15))

        def on_ok():
            result_holder[0] = True
            win.destroy()
            event.set()

        def on_cancel():
            result_holder[0] = False
            win.destroy()
            event.set()

        tk.Button(btn_frame, text=ok_text, command=on_ok,
                  bg="#4CAF50", fg="white",
                  font=("맑은 고딕", 10, "bold"), padx=15, pady=5).pack(side="left", padx=5)
        tk.Button(btn_frame, text=cancel_text, command=on_cancel,
                  font=("맑은 고딕", 10), padx=15, pady=5).pack(side="left", padx=5)

        win.protocol("WM_DELETE_WINDOW", on_cancel)

    # 메인 스레드에서 창 표시 (caller가 _ROOT 를 전달해야 함)
    _ROOT.after(0, _show)
    event.wait()
    return result_holder[0]


# ─── 공유 tkinter 루트 참조 ──────────────────────────────────────────────────
# 각 auto*.py 에서 실행 시 _ROOT 를 실제 root 로 교체해야 합니다.
_ROOT: tk.Tk = None


def set_root(root: tk.Tk):
    global _ROOT
    _ROOT = root


# ─── OKOSC 창 찾기 ───────────────────────────────────────────────────────────

def find_okosc_app():
    """
    pywinauto를 사용해 OKOSC 애플리케이션 창을 찾습니다.
    반환: pywinauto Application 인스턴스
    """
    from pywinauto import Desktop

    desktop = Desktop(backend="uia")
    for w in desktop.windows():
        title = w.window_text()
        if any(kw in title for kw in config.OKOSC_WINDOW_KEYWORDS):
            return w
    raise RuntimeError(
        "OKOSC 프로그램 창을 찾을 수 없습니다.\n"
        "OKOSC를 먼저 실행해 주세요.\n"
        f"찾는 키워드: {config.OKOSC_WINDOW_KEYWORDS}\n"
        "config.py 의 OKOSC_WINDOW_KEYWORDS를 실제 창 제목으로 수정하세요."
    )


def print_okosc_controls():
    """
    OKOSC 창의 컨트롤 식별자를 출력합니다 (개발/디버깅용).
    터미널에서 실행: python -c "from utils import print_okosc_controls; print_okosc_controls()"
    """
    w = find_okosc_app()
    w.print_control_identifiers()
