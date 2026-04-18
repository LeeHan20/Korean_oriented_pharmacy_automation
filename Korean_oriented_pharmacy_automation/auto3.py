"""
자동화 3번 - 배송메모 입력
===========================================
흐름:
  1. 주문등록_출력(복수건)_출력완료 파일에서
     G열(이름)을 키로, D열(운송장번호)을 값으로 매핑
     → "로젠 000000" 형식 문자열 생성
  2. OKOSC 처방전검색(출력) → 대기처방 → 날짜(7일 전~오늘, 조제) 검색
  3. 각 행의 이름과 매핑하여 메모수정 → 배송메모 → "로젠 000000" 저장
"""

import os
import re
import time
import queue
import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox
from datetime import datetime

import openpyxl

import config
import utils


# ═══════════════════════════════════════════════════════════════════════════════
#  자동화 로직 함수들
# ═══════════════════════════════════════════════════════════════════════════════

def step1_build_tracking_map() -> dict:
    """
    Downloads 폴더에서 최신 '주문등록_출력(복수건)_출력완료' 파일을 읽어
    {이름: "로젠 XXXXXX"} 딕셔너리 반환.

    D열: 운송장번호 (연속된 숫자열)
    G열: 이름 (primary key)
    """
    excel_path = utils.get_latest_file(
        config.DOWNLOAD_DIR,
        "주문등록_출력*복수건*_출력완료*.xls*"
    )

    # xls → xlsx 변환 필요 시
    if excel_path.lower().endswith(".xls"):
        excel_path = utils.xls_to_xlsx(excel_path)

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    tracking_map = {}
    # 숫자로만 이루어진 운송장번호 패턴
    num_pattern = re.compile(r'\d+')

    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 7:
            continue
        d_val = row[3]   # D열 (0-indexed: 3)
        g_val = row[6]   # G열 (0-indexed: 6)

        if not g_val or not d_val:
            continue

        name = str(g_val).strip()
        d_str = str(d_val).strip()

        # 운송장번호: 연속된 숫자 추출
        nums = num_pattern.findall(d_str)
        if nums:
            tracking_no = max(nums, key=len)   # 가장 긴 숫자열 선택
            tracking_map[name] = f"로젠 {tracking_no}"

    wb.close()

    if not tracking_map:
        raise ValueError(
            "주문등록 파일에서 운송장번호를 찾을 수 없습니다.\n"
            f"파일: {excel_path}"
        )

    return tracking_map, excel_path


def step2_3_enter_delivery_memos(tracking_map: dict, log_fn=None):
    """
    OKOSC 처방전검색 결과 창에서
    각 행의 이름을 키로 배송메모를 입력합니다.

    OKOSC UI 상세 사항은 실제 프로그램을 보고 아래 TODO 부분을 수정하세요.
    """
    from pywinauto import Desktop
    import pyautogui

    def log(msg):
        if log_fn:
            log_fn(msg)

    start_date, end_date = utils.get_search_dates()

    # ── OKOSC 창 찾기 ──────────────────────────────────────────────────────────
    okosc_win = utils.find_okosc_app()
    okosc_win.set_focus()
    time.sleep(0.3)

    # ── 처방전검색(출력) 클릭 ──────────────────────────────────────────────────
    # TODO: 실제 컨트롤 이름으로 수정 필요 (utils.print_okosc_controls() 참조)
    try:
        okosc_win.child_window(title_re=".*처방전검색.*출력.*",
                               control_type="Button").click_input()
    except Exception:
        try:
            okosc_win.child_window(title_re=".*처방전검색.*").click_input()
        except Exception as e:
            raise RuntimeError(
                f"OKOSC '처방전검색(출력)' 버튼을 찾지 못했습니다.\n상세: {e}"
            )
    time.sleep(0.5)

    # ── 검색 다이얼로그 ────────────────────────────────────────────────────────
    try:
        search_dlg = Desktop(backend="uia").window(title_re=".*처방.*")
        search_dlg.wait("visible", timeout=5)
    except Exception:
        search_dlg = okosc_win

    # ── 대기처방 선택 ──────────────────────────────────────────────────────────
    try:
        search_dlg.child_window(title_re=".*대기처방.*").click_input()
    except Exception:
        pass

    # ── 날짜 설정 ─────────────────────────────────────────────────────────────
    _set_date_field_safe(search_dlg, "시작일", start_date)
    _set_date_field_safe(search_dlg, "종료일", end_date)

    # ── 조제 옵션 ─────────────────────────────────────────────────────────────
    try:
        search_dlg.child_window(title_re=".*조제.*",
                                 control_type="RadioButton").click_input()
    except Exception:
        pass

    # ── 검색 클릭 ─────────────────────────────────────────────────────────────
    try:
        search_dlg.child_window(title_re="검색|조회",
                                 control_type="Button").click_input()
    except Exception:
        search_dlg.child_window(title_re="검색|조회").click_input()
    time.sleep(1)

    # ── 검색 결과 창 / 그리드에서 각 행 처리 ──────────────────────────────────
    # TODO: 아래는 리스트 컨트롤 이름과 열 인덱스를 실제 값으로 수정해야 합니다.
    #       utils.print_okosc_controls() 로 컨트롤 목록을 먼저 확인하세요.

    result_rows = _get_okosc_result_rows(search_dlg)
    matched = 0

    for row_idx, row_info in enumerate(result_rows):
        name = row_info.get("name", "")
        if name not in tracking_map:
            log(f"  매핑 없음: {name}")
            continue

        tracking_str = tracking_map[name]
        log(f"  [{row_idx+1}] {name} → {tracking_str}")

        # 해당 행 선택
        _select_okosc_row(search_dlg, row_idx)

        # 메모수정 → 배송메모 입력 → 저장
        _enter_delivery_memo(search_dlg, tracking_str)
        matched += 1
        time.sleep(0.3)

    log(f"배송메모 입력 완료: {matched}/{len(result_rows)}건")


def _get_okosc_result_rows(dlg) -> list:
    """
    OKOSC 검색 결과 창에서 행 목록 추출.
    TODO: 실제 그리드 컨트롤 이름으로 수정 필요.
    """
    rows = []
    try:
        # DataGrid / ListView 형태
        grid = dlg.child_window(control_type="DataGrid")
        for i, item in enumerate(grid.children(control_type="DataItem")):
            name_cell = item.children()[0]   # TODO: 이름 열 인덱스 확인
            rows.append({"name": name_cell.window_text().strip(), "_idx": i})
    except Exception:
        try:
            list_ctrl = dlg.child_window(control_type="List")
            for i, item in enumerate(list_ctrl.items()):
                rows.append({"name": item.window_text().strip(), "_idx": i})
        except Exception:
            pass
    return rows


def _select_okosc_row(dlg, row_idx: int):
    """OKOSC 검색 결과에서 row_idx 번째 행을 선택."""
    try:
        grid = dlg.child_window(control_type="DataGrid")
        item = grid.children(control_type="DataItem")[row_idx]
        item.click_input()
    except Exception:
        try:
            list_ctrl = dlg.child_window(control_type="List")
            list_ctrl.items()[row_idx].click_input()
        except Exception:
            pass


def _enter_delivery_memo(dlg, memo_text: str):
    """
    OKOSC 메모수정 창을 열고 배송메모 필드에 memo_text 를 입력 후 저장.
    TODO: 실제 컨트롤 이름으로 수정 필요.
    """
    from pywinauto import Desktop

    # 메모수정 버튼/링크 클릭
    try:
        dlg.child_window(title_re=".*메모수정.*|.*메모 수정.*").click_input()
    except Exception:
        try:
            dlg.child_window(title_re=".*메모.*", control_type="Button").click_input()
        except Exception as e:
            raise RuntimeError(f"메모수정 버튼을 찾지 못했습니다: {e}")
    time.sleep(0.4)

    # 메모수정 팝업/다이얼로그
    try:
        memo_dlg = Desktop(backend="uia").window(title_re=".*메모.*")
        memo_dlg.wait("visible", timeout=5)
    except Exception:
        memo_dlg = dlg

    # 배송메모 입력 필드 찾기
    try:
        field = memo_dlg.child_window(title_re=".*배송메모.*", control_type="Edit")
    except Exception:
        # 배송메모 레이블 옆의 Edit 컨트롤 찾기
        try:
            fields = memo_dlg.children(control_type="Edit")
            # TODO: 배송메모 필드가 몇 번째인지 확인 후 인덱스 수정
            field = fields[-1]
        except Exception as e:
            raise RuntimeError(f"배송메모 입력 필드를 찾지 못했습니다: {e}")

    field.click_input()
    field.set_text(memo_text)
    time.sleep(0.2)

    # 저장 버튼
    try:
        memo_dlg.child_window(title_re="저장|확인",
                               control_type="Button").click_input()
    except Exception:
        memo_dlg.child_window(title_re="저장|확인").click_input()
    time.sleep(0.3)


def _set_date_field_safe(dlg, hint: str, date_str: str):
    """날짜 필드 안전하게 설정 (예외 무시)."""
    try:
        field = dlg.child_window(title_re=f".*{hint}.*", control_type="Edit")
        field.set_text(date_str)
    except Exception:
        try:
            field = dlg.child_window(title_re=f".*{hint}.*",
                                      control_type="DateTimePicker")
            import win32com.client
            import pyautogui
            rect = field.rectangle()
            x = (rect.left + rect.right) // 2
            y = (rect.top + rect.bottom) // 2
            pyautogui.click(x, y)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.write(date_str.replace("-", ""), interval=0.05)
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════════════════════
#  tkinter GUI 앱
# ═══════════════════════════════════════════════════════════════════════════════

class Auto3App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("자동화 3번 - 배송메모 입력")
        self.root.geometry("480x480")
        self.root.resizable(True, True)

        utils.set_root(root)

        self._q: queue.Queue = queue.Queue()
        self._running = False

        self._build_ui()
        self._poll_queue()

    def _build_ui(self):
        top = tk.Frame(self.root, bg="#6C3483", padx=10, pady=8)
        top.pack(fill="x")
        tk.Label(top, text="자동화 3번 - 배송메모 입력",
                 font=("맑은 고딕", 13, "bold"),
                 bg="#6C3483", fg="white").pack()

        self._status_var = tk.StringVar(value="■ 대기 중")
        tk.Label(self.root, textvariable=self._status_var,
                 font=("맑은 고딕", 10), fg="#7D3C98",
                 anchor="w", padx=10).pack(fill="x", pady=(6, 0))

        log_frame = tk.Frame(self.root)
        log_frame.pack(fill="both", expand=True, padx=10, pady=4)
        self._log = scrolledtext.ScrolledText(
            log_frame, height=16, font=("Consolas", 9),
            state="disabled", bg="#FAFAFA"
        )
        self._log.pack(fill="both", expand=True)

        btn_frame = tk.Frame(self.root, pady=8)
        btn_frame.pack()
        self._run_btn = tk.Button(
            btn_frame, text="  시 작  ",
            command=self._start,
            font=("맑은 고딕", 11, "bold"),
            bg="#6C3483", fg="white", padx=18, pady=6,
            relief="flat", cursor="hand2"
        )
        self._run_btn.pack(side="left", padx=6)
        tk.Button(
            btn_frame, text="  닫 기  ",
            command=self.root.destroy,
            font=("맑은 고딕", 11),
            padx=18, pady=6,
            relief="flat", cursor="hand2"
        ).pack(side="left", padx=6)

    def _log_write(self, msg: str):
        self._log.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.insert("end", f"[{ts}] {msg}\n")
        self._log.see("end")
        self._log.config(state="disabled")

    def _poll_queue(self):
        while not self._q.empty():
            kind, payload = self._q.get_nowait()
            if kind == "log":
                self._log_write(payload)
            elif kind == "status":
                self._status_var.set(payload)
            elif kind == "done":
                self._running = False
                self._run_btn.config(state="normal")
        self.root.after(80, self._poll_queue)

    def _put(self, kind: str, payload: str):
        self._q.put((kind, payload))

    def _log_msg(self, msg: str):
        self._put("log", msg)
        self._put("status", f"▶ {msg}")

    def _start(self):
        if self._running:
            return
        self._running = True
        self._run_btn.config(state="disabled")
        threading.Thread(target=self._run_all, daemon=True).start()

    def _run_all(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            # 1. 운송장 매핑 생성
            self._log_msg("1단계: 운송장 매핑 생성 중...")
            tracking_map, excel_path = step1_build_tracking_map()
            self._log_msg(f"  ✓ {len(tracking_map)}건 로드: {os.path.basename(excel_path)}")
            for name, code in list(tracking_map.items())[:5]:
                self._log_msg(f"    {name} → {code}")
            if len(tracking_map) > 5:
                self._log_msg(f"    ... 외 {len(tracking_map)-5}건")

            # 2-3. OKOSC 배송메모 입력
            self._log_msg("2~3단계: OKOSC 배송메모 입력 중...")
            step2_3_enter_delivery_memos(
                tracking_map,
                log_fn=lambda m: self._log_msg(m)
            )

            self._log_msg("━━━ 자동화 3번 완료 ━━━")
            self._put("status", "✅ 완료")

        except Exception as e:
            self._log_msg(f"오류 발생: {e}")
            self._put("status", "❌ 오류 발생")
            self.root.after(0, lambda e=e: messagebox.showerror(
                "오류", str(e), parent=self.root
            ))
        finally:
            pythoncom.CoUninitialize()
            self._put("done", "")


# ═══════════════════════════════════════════════════════════════════════════════
#  진입점
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = Auto3App(root)
    root.mainloop()
