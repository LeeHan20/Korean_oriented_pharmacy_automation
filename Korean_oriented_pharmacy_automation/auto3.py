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

def step1_build_tracking_map() -> tuple:
    """
    Downloads 폴더에서 최신 '주문등록_출력(복수건)_출력완료' 파일을 읽어
    {처방번호: "로젠 XXXXXX"} 딕셔너리 반환.

    D열: 운송장번호 (연속된 숫자열)
    S열: autocode (4자리, 0-indexed: 18) → DB에서 처방번호 역조회
    """
    import json

    excel_path = utils.get_latest_file(
        config.DOWNLOAD_DIR,
        "주문등록_출력*복수건*_출력완료*.xls*"
    )

    # xls → xlsx 변환 필요 시
    if excel_path.lower().endswith(".xls"):
        excel_path = utils.xls_to_xlsx(excel_path)

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    # S열(index 18) autocode → 운송장번호 매핑
    autocode_to_tracking = {}
    num_pattern = re.compile(r'\d+')

    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 19:
            continue
        d_val = row[3]    # D열: 운송장번호
        s_val = row[18]   # S열: autocode (4자리)

        if not d_val or not s_val:
            continue

        d_str = str(d_val).strip()
        s_str = str(s_val).strip()

        nums = num_pattern.findall(d_str)
        if not nums:
            continue
        tracking_no = max(nums, key=len)

        ac_nums = num_pattern.findall(s_str)
        if not ac_nums:
            continue
        autocode = ac_nums[0]

        autocode_to_tracking[autocode] = f"로젠 {tracking_no}"

    if not autocode_to_tracking:
        raise ValueError(
            "주문등록 파일에서 autocode/운송장번호를 찾을 수 없습니다.\n"
            f"파일: {excel_path}"
        )

    # DB 로드: {처방번호: autocode} → 역변환 {autocode: 처방번호}
    if not os.path.exists(config.AUTOCODE_DB_PATH):
        raise FileNotFoundError(
            f"autocode DB 파일이 없습니다: {config.AUTOCODE_DB_PATH}\n"
            "자동화 1번을 먼저 실행하세요."
        )
    with open(config.AUTOCODE_DB_PATH, 'r', encoding='utf-8') as f:
        db = json.load(f)  # {처방번호: autocode}

    ac_to_presc = {v: k for k, v in db.items()}  # {autocode: 처방번호}

    # 최종 매핑: {처방번호: "로젠 XXXXXX"}
    tracking_map = {}
    unmatched = []
    for autocode, tracking_str in autocode_to_tracking.items():
        presc_no = ac_to_presc.get(autocode)
        if presc_no:
            tracking_map[presc_no] = tracking_str
        else:
            unmatched.append(autocode)

    if unmatched:
        # 미매칭은 경고만 (일부 항목은 익산대장 등 DB 외 항목일 수 있음)
        print(f"[경고] autocode DB 미매칭: {unmatched}")

    if not tracking_map:
        raise ValueError(
            "주문등록 파일의 autocode와 DB가 일치하는 항목이 없습니다.\n"
            f"파일: {excel_path}\n미매칭 autocode: {unmatched}"
        )

    return tracking_map, excel_path


def step2_3_4_enter_memo_and_complete(tracking_map: dict, log_fn=None):
    """
    처방번호별로 배송메모 입력 → 완료상태 처리를 순차 실행합니다.
    흐름: (1회) 검색 설정·실행 → 항목별 메모 입력 → 완료 처리 반복
    """
    import json
    import tempfile

    def log(msg):
        if log_fn:
            log_fn(msg)

    # ── OKOSC 검색 설정 및 실행 (1회) ───────────────────────────────────────
    log("OKOSC 검색 설정 중 (진행상태=조제, 7일)...")
    search_result = utils.call_okosc_worker("setup_search", timeout=30)
    if search_result.get("status") != "ok":
        raise RuntimeError(f"OKOSC 검색 설정 실패: {search_result.get('message')}")
    log("  ✓ 검색 완료")

    total = len(tracking_map)
    success = 0
    failed = []

    for idx, (presc_no, tracking_str) in enumerate(tracking_map.items(), 1):
        log(f"[{idx}/{total}] {presc_no} → {tracking_str}")

        # ── 배송메모 입력 ────────────────────────────────────────────────────
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8',
                                         suffix='.json', delete=False) as f:
            json.dump({presc_no: tracking_str}, f, ensure_ascii=False)
            memo_json = f.name
        try:
            memo_result = utils.call_okosc_worker(
                "enter_delivery_memos",
                extra_args=[memo_json],
                timeout=60,
            )
        finally:
            try:
                os.remove(memo_json)
            except Exception:
                pass

        if memo_result.get("status") != "ok":
            log(f"  ✗ 배송메모 실패: {memo_result.get('message')}")
            failed.append(presc_no)
            continue

        for r in memo_result.get("results", []):
            mark = "✓" if r["status"] == "ok" else "✗"
            err_txt = f" ({r['error']})" if r.get("error") else ""
            log(f"  [{mark}] 메모: {r.get('memo', '')}{err_txt}")

        # ── 완료상태 처리 ────────────────────────────────────────────────────
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8',
                                         suffix='.json', delete=False) as f:
            json.dump({"presc_nos": [presc_no]}, f, ensure_ascii=False)
            complete_json = f.name
        try:
            complete_result = utils.call_okosc_worker(
                "check_and_complete",
                extra_args=[complete_json],
                timeout=60,
            )
        finally:
            try:
                os.remove(complete_json)
            except Exception:
                pass

        if complete_result.get("status") != "ok":
            log(f"  ✗ 완료상태 실패: {complete_result.get('message')}")
            failed.append(presc_no)
            continue

        checked = complete_result.get("checked", 0)
        clicked = complete_result.get("complete_clicked", False)
        log(f"  ✓ 완료: {checked}건 체크, 클릭={'성공' if clicked else '실패'}")
        success += 1

    log(f"전체 처리 완료: {success}/{total}건" + (f", 실패: {failed}" if failed else ""))


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

            # 2-3-4. OKOSC 배송메모 입력 → 완료상태 (항목별 순차)
            self._log_msg("2~4단계: 배송메모 입력 및 완료상태 처리 중...")
            step2_3_4_enter_memo_and_complete(
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
