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
    """
    OKOSC 창을 hwnd 기반으로 찾아 uia backend로 연결합니다.
    utils.find_okosc_app()과 동일한 키워드/클래스 조건을 사용하되,
    DataItem 열거를 위해 uia backend를 사용합니다.
    """
    import win32gui
    import win32con
    from pywinauto import Application

    KEYWORDS = ["OKOCSTJS", "탕전실", "OKOSC", "OK처방", "처방프로그램"]

    found_hwnd = None

    def _enum(hwnd, _):
        nonlocal found_hwnd
        title = win32gui.GetWindowText(hwnd)
        cls = win32gui.GetClassName(hwnd)
        if any(kw in title for kw in KEYWORDS) and 'WindowsForms' in cls:
            found_hwnd = hwnd

    win32gui.EnumWindows(_enum, None)

    if not found_hwnd:
        raise RuntimeError(
            "OKOSC 창을 찾을 수 없습니다.\n"
            f"찾는 키워드: {KEYWORDS}"
        )

    win32gui.ShowWindow(found_hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(found_hwnd)
    time.sleep(0.3)

    app = Application(backend="uia").connect(handle=found_hwnd)
    return app.window(handle=found_hwnd)


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

# ═══════════════════════════════════════════════════════════════════════════════
#  setup_search: OKOSC 검색 설정 및 실행 (auto3.py 루프 전 1회 호출)
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_setup_search():
    """
    OKOSC 검색 조건 설정 및 검색 실행.
    - 검색기준: 진행상태
    - 날짜: 7일 전 ~ 오늘
    - 진행상태 필터: 조제
    auto3.py에서 enter_delivery_memos / check_and_complete 루프 전에 1회 호출.
    """
    import pyautogui
    from pywinauto.keyboard import send_keys
    from datetime import datetime, timedelta

    win = _okosc_win()
    win.set_focus()
    time.sleep(0.5)

    # ── 검색 기준 → 진행상태 ─────────────────────────────────────────────────
    try:
        ctrl = win.child_window(auto_id="ulCboSearch")
        r = ctrl.rectangle()
        pyautogui.click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
        time.sleep(0.2)
        send_keys('^a')
        send_keys('진행상태')
        time.sleep(0.1)
        send_keys('{ENTER}')
        time.sleep(0.3)
    except Exception:
        pass

    # ── 날짜 설정 ────────────────────────────────────────────────────────────
    today = datetime.now()
    start_date = (today - timedelta(days=30)).strftime("%Y-%m-%d")
    end_date = today.strftime("%Y-%m-%d")

    def _set_ultra_date(auto_id, date_str):
        y, m, d = date_str.split("-")
        yy = y[2:]
        ctrl = win.child_window(auto_id=auto_id)
        rect = ctrl.rectangle()
        cx = (rect.left + rect.right) // 2
        cy = (rect.top + rect.bottom) // 2
        pyautogui.click(cx, cy)
        time.sleep(0.2)
        send_keys('{END}')
        time.sleep(0.1)
        for _ in range(8):
            send_keys('{BACKSPACE}')
            time.sleep(0.05)
        time.sleep(0.1)
        for ch in yy + m + d:
            send_keys(ch)
            time.sleep(0.05)
        send_keys('{TAB}')
        time.sleep(0.3)

    _set_ultra_date("ulDteSearchStart", start_date)
    _set_ultra_date("ulDteSearchEnd", end_date)

    send_keys('{ESC}')
    time.sleep(0.2)
    win.set_focus()
    time.sleep(0.2)

    # ── 진행상태 → 조제 선택 ────────────────────────────────────────────────
    try:
        state_ctrl = win.child_window(auto_id="ulCboSearchCBJState")
        rect = state_ctrl.rectangle()
        pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
    except Exception:
        try:
            state_ctrl = win.child_window(title_re="조제|대기|완료|취소",
                                          class_name_re=".*UltraComboEditor.*")
            rect = state_ctrl.rectangle()
            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
        except Exception:
            pass
    time.sleep(0.15)
    send_keys('^a')
    send_keys('조제')
    send_keys('{ENTER}')
    time.sleep(0.3)

    # ── 검색 실행 ────────────────────────────────────────────────────────────
    win.child_window(auto_id="ulBtnSearchCBJ").click_input()
    time.sleep(2)

    return {"status": "ok"}


# ═══════════════════════════════════════════════════════════════════════════════
#  enter_delivery_memos: 배송메모 입력 (auto3.py에서 호출)
#  사전에 cmd_setup_search()가 호출되어 검색 결과가 표시되어 있어야 합니다.
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_enter_delivery_memos():
    """
    배송메모 일괄 입력.
    sys.argv[2]: {처방번호: "로젠 XXXXXX"} 형태의 tracking_map JSON 파일 경로

    사전에 cmd_setup_search()가 호출되어 검색 결과가 표시되어 있어야 합니다.
    """
    import pyautogui
    from pywinauto.keyboard import send_keys
    from pywinauto import Desktop
    import json

    if len(sys.argv) < 3:
        return {"status": "error", "message": "tracking_map JSON 파일 경로 필요 (argv[2])"}

    with open(sys.argv[2], 'r', encoding='utf-8') as f:
        tracking_map = json.load(f)

    if not tracking_map:
        return {"status": "error", "message": "tracking_map이 비어 있습니다"}

    win = _okosc_win()
    win.set_focus()
    time.sleep(0.3)

    # ── 검색 결과 DataItem 열거 → 처방번호 기준으로 행 클릭 ─────────────────
    import collections
    _row_re_local = _re.compile(r'\s+Row(\d+)$')

    all_items = win.descendants(control_type="DataItem")

    # 1) row_idx → {col_name: (val, item)} 맵 구성
    rows_dict = collections.defaultdict(dict)
    for item in all_items:
        name = item.element_info.name
        m = _row_re_local.search(name)
        if not m:
            continue
        row_idx = int(m.group(1))
        col_name = name[:m.start()]
        try:
            val = item.iface_value.CurrentValue or ""
        except Exception:
            val = ""
        rows_dict[row_idx][col_name] = {"val": val, "item": item}

    results = []
    not_matched = []

    for row_idx in sorted(rows_dict.keys()):
        row = rows_dict[row_idx]
        presc_no = row.get("처방번호", {}).get("val", "").strip()
        if not presc_no:
            continue

        if presc_no not in tracking_map:
            not_matched.append(presc_no)
            continue

        tracking_str = tracking_map[presc_no]

        # 행의 임의 셀 클릭으로 행 선택
        click_item = row.get("처방번호", {}).get("item") or row.get("환자명", {}).get("item")
        try:
            if click_item:
                click_item.click_input()
            time.sleep(0.3)
        except Exception:
            pass

        # 메모수정 버튼 클릭 (우측 패널)
        try:
            win.child_window(title_re="메모수정").click_input()
        except Exception:
            try:
                win.child_window(auto_id="ulBtnMemoCBJ").click_input()
            except Exception:
                pass
        time.sleep(0.6)

        # ── 처방전메모 다이얼로그를 win32gui로 탐색 후 UIA 연결 ──────────────
        import win32gui as _wg
        dlg_hwnd = None
        deadline2 = time.time() + 6
        while time.time() < deadline2:
            def _find_dlg(hwnd, _):
                nonlocal dlg_hwnd
                t = _wg.GetWindowText(hwnd)
                if '처방전메모' in t and _wg.IsWindowVisible(hwnd):
                    dlg_hwnd = hwnd
            _wg.EnumWindows(_find_dlg, None)
            if dlg_hwnd:
                break
            time.sleep(0.3)

        if not dlg_hwnd:
            results.append({"presc_no": presc_no, "status": "error",
                             "error": "처방전메모 다이얼로그를 찾을 수 없음"})
            continue

        from pywinauto import Application as _App
        memo_dlg = _App(backend="uia").connect(handle=dlg_hwnd).window(handle=dlg_hwnd)

        # auto_id로 직접 배송메모 필드 탐색 (ulTxtMemoBesong)
        try:
            edits = memo_dlg.descendants(control_type="Edit")
            field = None
            for ed in edits:
                if ed.element_info.automation_id == "ulTxtMemoBesong":
                    field = ed
                    break
            # EmbeddableTextBox(실제 입력 자식)가 있으면 그쪽으로
            if field is not None:
                children = field.children(control_type="Edit")
                if children:
                    field = children[0]
        except Exception:
            field = None

        if field is None:
            results.append({"presc_no": presc_no, "status": "error",
                             "error": "배송메모 Edit 필드를 찾을 수 없음"})
            try:
                memo_dlg.child_window(title="닫기", control_type="Button").click_input()
            except Exception:
                pass
            continue

        # 배송메모 입력
        field.click_input()
        time.sleep(0.1)
        field.set_text(tracking_str)
        time.sleep(0.2)

        # 저장 버튼: auto_id에 "Besong" 포함된 버튼 우선, 없으면 index 1
        try:
            all_btns = memo_dlg.descendants(control_type="Button")
            save_btn = None
            for b in all_btns:
                aid_b = b.element_info.automation_id
                if "Besong" in aid_b and b.window_text().strip() == "저장":
                    save_btn = b
                    break
            if save_btn is None:
                save_btns = [b for b in all_btns if b.window_text().strip() == "저장"]
                save_btn = save_btns[1] if len(save_btns) > 1 else (save_btns[0] if save_btns else None)
            if save_btn:
                save_btn.click_input()
        except Exception as e:
            results.append({"presc_no": presc_no, "status": "error",
                             "error": f"저장 버튼 클릭 실패: {e}"})
            try:
                memo_dlg.child_window(title="닫기", control_type="Button").click_input()
            except Exception:
                pass
            continue
        time.sleep(0.3)

        # 닫기
        try:
            memo_dlg.child_window(title="닫기", control_type="Button").click_input()
        except Exception:
            try:
                memo_dlg.child_window(title_re="닫기|Close",
                                       control_type="Button").click_input()
            except Exception:
                pass
        time.sleep(0.4)

        results.append({"presc_no": presc_no, "memo": tracking_str, "status": "ok"})


    return {
        "status": "ok",
        "matched": sum(1 for r in results if r["status"] == "ok"),
        "not_matched": not_matched,
        "results": results,
    }


# ═══════════════════════════════════════════════════════════════════════════════
#  get_presc_numbers: 중디 검색결과에서 처방번호 목록 반환 (auto1.py step5에서 호출)
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_get_presc_numbers():
    """
    OKOSC 검색 결과 DataItem을 파싱합니다.
    각 DataItem name = "{컬럼명} Row{N}" 형식, iface_value.CurrentValue = 셀 값.
    "처방번호 RowN" / "환자명 RowN" 을 row 인덱스로 그룹핑해 반환합니다.
    """
    import collections
    _row_re = _re.compile(r'\s+Row(\d+)$')

    win = _okosc_win()
    all_items = win.descendants(control_type="DataItem")

    rows_dict = collections.defaultdict(dict)  # {row_idx: {col_name: val}}
    for item in all_items:
        name = item.element_info.name
        m = _row_re.search(name)
        if not m:
            continue
        row_idx = int(m.group(1))
        col_name = name[:m.start()]
        try:
            val = item.iface_value.CurrentValue or ""
        except Exception:
            val = ""
        rows_dict[row_idx][col_name] = val

    items_data = []
    for row_idx in sorted(rows_dict.keys()):
        row = rows_dict[row_idx]
        presc_no = row.get("처방번호", "").strip()
        patient  = row.get("환자명", "").strip()
        if presc_no:
            items_data.append({"presc_no": presc_no, "patient": patient})

    return {"status": "ok", "items": items_data}


# ═══════════════════════════════════════════════════════════════════════════════
#  check_and_complete: 체크박스 선택 후 상태변경→완료상태 (auto3.py step4)
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_check_and_complete():
    """
    argv[2]: {"presc_nos": ["8295", ...]} JSON 파일 경로
    1) 스크린샷 그리드에서 해당 처방번호 행의 체크박스 클릭 (선택)
    2) 상태변경 버튼 클릭 → 완료상태 뮩뉴 클릭
    """
    import collections
    import json
    import pyautogui
    from pywinauto.keyboard import send_keys

    if len(sys.argv) < 3:
        return {"status": "error", "message": "presc_nos JSON 파일 경로 필요 (argv[2])"}

    with open(sys.argv[2], 'r', encoding='utf-8') as f:
        presc_nos_set = set(json.load(f).get("presc_nos", []))

    if not presc_nos_set:
        return {"status": "error", "message": "presc_nos가 비어 있습니다"}

    win = _okosc_win()
    _row_re = _re.compile(r'\s+Row(\d+)$')

    all_items = win.descendants(control_type="DataItem")
    rows_dict = collections.defaultdict(dict)  # {row_idx: {col_name: item}}
    for item in all_items:
        name = item.element_info.name
        m = _row_re.search(name)
        if not m:
            continue
        row_idx = int(m.group(1))
        col_name = name[:m.start()]
        try:
            val = item.iface_value.CurrentValue or ""
        except Exception:
            val = ""
        rows_dict[row_idx][col_name] = {"val": val, "item": item}

    checked_count = 0
    not_found = []

    for row_idx in sorted(rows_dict.keys()):
        row = rows_dict[row_idx]
        presc_no = row.get("처방번호", {}).get("val", "").strip()
        if not presc_no:
            continue
        if presc_no not in presc_nos_set:
            continue

        # 체크박스 셀 탐색: Toggle 인터페이스 지원 항목 우선
        toggled = False
        for col_name, cell in row.items():
            item = cell["item"]
            try:
                state = item.iface_toggle.CurrentToggleState
                if state == 0:  # off → on
                    item.iface_toggle.Toggle()
                toggled = True
                break
            except Exception:
                pass

        if not toggled:
            # fallback: 좌측 기준 두 번째 셀(체크박스 열)
            items_in_row = sorted(row.values(), key=lambda c: c["item"].rectangle().left)
            if len(items_in_row) > 1:
                cb_item = items_in_row[1]["item"]
                cb_item.click_input()
                toggled = True

        if toggled:
            checked_count += 1
        else:
            not_found.append(presc_no)

        time.sleep(0.1)

    if checked_count == 0:
        return {"status": "error", "message": "체크할 항목을 찾을 수 없습니다", "not_found": not_found}

    time.sleep(0.3)
    win.set_focus()

    # 상태변경 버튼 클릭
    try:
        win.child_window(title="상태변경", control_type="Button").click_input()
    except Exception:
        try:
            win.child_window(auto_id="ulBtnChgState").click_input()
        except Exception:
            # 마지막 fallback: 첣아서 클릭
            for ctrl in win.descendants(control_type="Button"):
                if "상태변경" in ctrl.window_text():
                    ctrl.click_input()
                    break
    time.sleep(0.5)

    # 완료상태 메뉴 항목 클릭
    from pywinauto import Desktop as _Desktop
    clicked_complete = False
    # 메뉴가 팝업되면 MenuItem 또는 Button으로 나탈 수 있음
    for ctrl_type in ("MenuItem", "ListItem", "Button"):
        try:
            items = _Desktop(backend="uia").windows()
            for w in items:
                for ci in w.descendants(control_type=ctrl_type):
                    if "완료" in ci.window_text() and "상태" in ci.window_text():
                        ci.click_input()
                        clicked_complete = True
                        break
                if clicked_complete:
                    break
            if clicked_complete:
                break
        except Exception:
            pass

    if not clicked_complete:
        # fallback: win 안에서 메뉴 항목 탐색
        try:
            win.child_window(title_re=".*완료.*상태.*|.*상태.*완료.*").click_input()
            clicked_complete = True
        except Exception:
            pass

    time.sleep(0.5)

    return {
        "status": "ok",
        "checked": checked_count,
        "not_found": not_found,
        "complete_clicked": clicked_complete,
    }


COMMANDS = {
    "step1":                cmd_step1,
    "get_herbs":            cmd_get_herbs,
    "get_patient":          cmd_get_patient,
    "get_dosage":           cmd_get_dosage,
    "setup_search":         cmd_setup_search,
    "enter_delivery_memos": cmd_enter_delivery_memos,
    "get_presc_numbers":    cmd_get_presc_numbers,
    "check_and_complete":   cmd_check_and_complete,
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
