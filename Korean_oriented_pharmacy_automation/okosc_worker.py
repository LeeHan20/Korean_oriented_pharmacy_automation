#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
okosc_worker.py  ← 반드시 32비트 Python으로 실행
==========================================
OKOSC 창 자동화 헬퍼 (insurance_med.py에서 subprocess로 호출)

사용법:
  python32.exe okosc_worker.py <command>

commands:
  step1       - 처방전송일자 기준 검색 → 보험 행 선택
  get_herbs   - 선택된 행의 우측 패널 약재 목록 반환
  get_patient - 환자명 / 연락처 / 주소 반환
  get_dosage  - 복용법 텍스트 반환

결과: stdout에 JSON (ensure_ascii=False)
오류 시: {"status": "error", "message": "..."}
"""

import sys
import json
import time
import re as _re

_PAREN_RE = _re.compile(r'\s*[\(（][^\)）]*[\)）]')


def _out(obj: dict):
    sys.stdout.buffer.write(json.dumps(obj, ensure_ascii=False).encode("utf-8"))
    sys.stdout.buffer.write(b"\n")
    sys.stdout.buffer.flush()


def _okosc_win():
    from pywinauto import Application
    app = Application(backend="uia").connect(title_re=".*OKOCSTJS.*", timeout=10)
    return app.window(title_re=".*OKOCSTJS.*")


# ═══════════════════════════════════════════════════════════════════════════════
#  step1: 검색 설정 → 검색 → 보험 행 클릭
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_step1():
    import pyautogui
    from pywinauto.keyboard import send_keys

    win = _okosc_win()
    win.set_focus()
    time.sleep(0.5)

    # ulCboSearch 클릭 → HOME → ENTER (처방전송일자 선택)
    try:
        ctrl = win.child_window(auto_id="ulCboSearch")
        r = ctrl.rectangle()
        pyautogui.click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
        time.sleep(0.2)
        send_keys('{HOME}')
        time.sleep(0.1)
        send_keys('{ENTER}')
        time.sleep(0.3)
    except Exception:
        pass

    # TAB × 6 → ENTER (검색 실행)
    for _ in range(6):
        send_keys('{TAB}')
        time.sleep(0.1)
    send_keys('{ENTER}')
    time.sleep(1.2)

    # 32비트 Python + UIA → DataItem 열거 정상 작동
    all_items = win.descendants(control_type="DataItem")
    for i, item in enumerate(all_items):
        try:
            cells = item.children()
            row_text = " ".join(c.window_text() for c in cells)
        except Exception:
            row_text = item.window_text()
        if "보험" in row_text:
            item.click_input()
            time.sleep(0.5)
            return {"status": "ok", "row_idx": i, "row_text": row_text}

    return {"status": "error", "message": "대기처방 목록에서 '보험' 행을 찾을 수 없습니다"}


# ═══════════════════════════════════════════════════════════════════════════════
#  get_herbs: 우측 약재 패널 목록 반환
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_get_herbs():
    def _rm(t):
        return _PAREN_RE.sub("", t).strip()

    win = _okosc_win()
    herbs = []

    # 그리드가 여러 개일 때 첫 번째가 처방 목록, 그 이후가 약재 패널
    grids = win.descendants(control_type="DataGrid")
    target_grids = grids[1:] if len(grids) > 1 else grids
    for grid in target_grids:
        for item in grid.children(control_type="DataItem"):
            cells = item.children()
            if len(cells) >= 2:
                name = _rm(cells[0].window_text())
                dose = _rm(cells[1].window_text())
                if name:
                    herbs.append([name, dose])
        if herbs:
            break

    return {"status": "ok", "herbs": herbs}


# ═══════════════════════════════════════════════════════════════════════════════
#  get_patient: 환자명 / 연락처 / 주소 반환
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_get_patient():
    win = _okosc_win()
    result = {"status": "ok", "name": "", "contact": "", "address": ""}

    for ctrl in win.descendants(control_type="Edit"):
        try:
            aid = ctrl.automation_id().lower()
            val = ctrl.window_text().strip()
            if not val:
                continue
            if any(k in aid for k in ("patient", "name", "환자", "성명")):
                result["name"] = val
            elif any(k in aid for k in ("tel", "phone", "연락", "전화")):
                result["contact"] = val
            elif any(k in aid for k in ("addr", "주소")):
                result["address"] = val
        except Exception:
            pass

    return result


# ═══════════════════════════════════════════════════════════════════════════════
#  get_dosage: 복용법 텍스트 반환
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_get_dosage():
    win = _okosc_win()
    dosage = ""

    for ctrl in win.descendants(control_type="Edit"):
        try:
            aid = ctrl.automation_id().lower()
            if any(k in aid for k in ("dosage", "복용", "복약")):
                dosage = ctrl.window_text().strip()
                if dosage:
                    break
        except Exception:
            pass

    return {"status": "ok", "dosage": dosage}


# ═══════════════════════════════════════════════════════════════════════════════
#  진입점
# ═══════════════════════════════════════════════════════════════════════════════

COMMANDS = {
    "step1":       cmd_step1,
    "get_herbs":   cmd_get_herbs,
    "get_patient": cmd_get_patient,
    "get_dosage":  cmd_get_dosage,
}

if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    fn = COMMANDS.get(cmd)
    if fn is None:
        _out({"status": "error", "message": f"unknown command: {cmd}"})
        sys.exit(1)
    try:
        _out(fn())
    except Exception as e:
        _out({"status": "error", "message": str(e)})
        sys.exit(1)
