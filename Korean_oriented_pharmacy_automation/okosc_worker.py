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
    import sys as _sys
    all_items = win.descendants(control_type="DataItem")
    print(f"[DEBUG step1] 검색 후 DataItem 개수={len(all_items)}", file=_sys.stderr)

    def _ct(item):
        t = item.window_text().strip()
        if t:
            return t
        try:
            ch = item.children()
            if ch:
                return ch[0].window_text().strip()
        except Exception:
            pass
        return ""

    # 앞 40개 DataItem 덤프 (main grid 구조 파악)
    for dbg_i, dbg_it in enumerate(all_items[:40]):
        try:
            r = dbg_it.rectangle()
            t = _ct(dbg_it)
            print(f"[DEBUG step1] item[{dbg_i:03d}] x={r.left:5d} y={r.top:5d} "
                  f"text={t!r}", file=_sys.stderr)
        except Exception:
            pass

    _NAME_RE  = _re.compile(r'^[가-힣]{2,5}$')
    _PHONE_RE = _re.compile(r'01[0-9][-\s]?\d{3,4}[-\s]?\d{4}')
    _ROW_RE   = _re.compile(r'\s+Row(\d+)$')

    # ── "컬럼명 RowN" 형식으로 DataItem 파싱 (enter_delivery_memos와 동일 방식) ──
    import collections as _col
    rows_dict = _col.defaultdict(dict)
    for item in all_items:
        name_txt = item.element_info.name
        m = _ROW_RE.search(name_txt)
        if not m:
            continue  # 헤더 셀 등 "RowN" 없는 항목 제외
        row_idx = int(m.group(1))
        col_name = name_txt[:m.start()].strip()
        try:
            val = item.iface_value.CurrentValue or ""
        except Exception:
            val = item.window_text().strip()
        rows_dict[row_idx][col_name] = {"val": val, "item": item}

    print(f"[DEBUG step1] rows_dict 행 수={len(rows_dict)}", file=_sys.stderr)
    for ri in sorted(rows_dict.keys()):
        row_cols = {c: d["val"] for c, d in rows_dict[ri].items()}
        print(f"[DEBUG step1] Row{ri}: {row_cols}", file=_sys.stderr)

    # 보험 컬럼 값이 정확히 "보험"인 행 찾기
    target_row_idx = None
    for row_idx in sorted(rows_dict.keys()):
        row = rows_dict[row_idx]
        if row.get("보험", {}).get("val", "").strip() == "보험":
            target_row_idx = row_idx
            break

    if target_row_idx is None:
        # fallback: window_text 방식 (구조가 다른 버전 대비)
        for i, item in enumerate(all_items):
            row_text = _ct(item)
            if row_text.strip() == "보험":
                print(f"[DEBUG step1] fallback 보험 행 idx={i}", file=_sys.stderr)
                item.click_input()
                time.sleep(0.5)
                return {
                    "status": "ok", "row_idx": i,
                    "row_text": row_text, "patient_name": "", "patient_contact": "",
                }
        return {"status": "error", "message": "대기처방 목록에서 '보험' 행을 찾을 수 없습니다"}

    target_row = rows_dict[target_row_idx]

    # 환자명 / 연락처 추출
    patient_name    = target_row.get("환자명",   {}).get("val", "").strip()
    patient_contact = (target_row.get("핸드폰",  {}).get("val", "") or
                       target_row.get("전화번호", {}).get("val", "")).strip()
    if not patient_contact:
        for data in target_row.values():
            m2 = _PHONE_RE.search(data.get("val", ""))
            if m2:
                patient_contact = m2.group()
                break
    row_text = target_row.get("보험", {}).get("val", "")

    print(f"[DEBUG step1] 보험 행 Row{target_row_idx}  환자명={patient_name!r}"
          f"  전화={patient_contact!r}", file=_sys.stderr)

    # 처방번호 셀(또는 첫 번째 셀) 클릭으로 행 선택
    click_item = (target_row.get("처방번호", {}).get("item") or
                  target_row.get("No",       {}).get("item") or
                  next((v["item"] for v in target_row.values()), None))
    if click_item is None:
        return {"status": "error", "message": "클릭할 셀을 찾을 수 없습니다"}

    click_item.click_input()
    time.sleep(0.5)
    return {
        "status":          "ok",
        "row_idx":         target_row_idx,
        "row_text":        row_text,
        "patient_name":    patient_name,
        "patient_contact": patient_contact,
    }


# ═══════════════════════════════════════════════════════════════════════════════
#  get_herbs: 우측 약재 패널 목록 반환
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_get_herbs():
    def _rm(t):
        return _PAREN_RE.sub("", t).strip()

    def _cell_text(item):
        """DataItem 셀의 텍스트를 반환합니다."""
        t = item.window_text().strip()
        if t:
            return t
        try:
            ch = item.children()
            if ch:
                return ch[0].window_text().strip()
        except Exception:
            pass
        return ""

    win = _okosc_win()
    herbs = []

    # OKOSC 약재 목록은 DataGrid가 아닌 Table 타입으로 노출됨
    # 각 DataItem = 1개 셀 (행이 아닌 셀 단위 flat 구조)
    tables = win.descendants(control_type="Table")

    # 가장 많은 DataItem을 가진 테이블 = 약재 목록 테이블
    herb_table = None
    max_items = 0
    for table in tables:
        try:
            items = table.children(control_type="DataItem")
            if len(items) > max_items:
                max_items = len(items)
                herb_table = table
        except Exception:
            pass

    if herb_table is None or max_items < 2:
        return {"status": "ok", "herbs": herbs}

    all_items = herb_table.children(control_type="DataItem")
    all_texts = [_cell_text(item) for item in all_items]

    # 연속된 정수 '1' → '2' 패턴으로 데이터 시작 위치와 열 개수 자동 감지
    start_idx = None
    col_count = None
    for i, t in enumerate(all_texts):
        if t.strip() == "1":
            for j in range(i + 1, min(i + 15, len(all_texts))):
                if all_texts[j].strip() == "2":
                    start_idx = i
                    col_count = j - i
                    break
            if start_idx is not None:
                break

    # '1'→'2' 패턴을 못 찾은 경우: 순번이 연속된 숫자인 첫 구간으로 추론
    if start_idx is None or col_count is None:
        # 숫자 '1'이 있는 위치를 찾고, 다음 숫자 '2'까지 거리를 열 개수로 추정
        for i, t in enumerate(all_texts):
            if t.strip() == "1":
                # 뒤에서 다음 숫자를 찾아 열 개수 추정
                for j in range(i + 1, min(i + 20, len(all_texts))):
                    if all_texts[j].strip().isdigit() and int(all_texts[j].strip()) == 2:
                        start_idx = i
                        col_count = j - i
                        break
                if start_idx is not None:
                    break

    if start_idx is None or col_count is None or col_count < 2:
        return {"status": "ok", "herbs": herbs}

    # 열 구조: [순번, 약재명, 1회투약량, 총용량, ?, 비고]
    # 약재명 = col 1, 1회투약량 = col 2
    for i in range(start_idx, len(all_texts), col_count):
        row = all_texts[i: i + col_count]
        if len(row) < 2:
            break
        row_num = row[0].strip()
        if not row_num.isdigit():
            break  # 데이터 행 종료
        name = row[1].strip()
        dose = row[2].strip() if len(row) > 2 else ""
        if name:
            herbs.append([_rm(name), _rm(dose)])

    # ── 디버그 출력 ──────────────────────────────────────────────────────
    import sys as _sys
    print(f"[DEBUG get_herbs] Table 개수={len(tables)}, 최대DataItem={max_items}",
          file=_sys.stderr)
    print(f"[DEBUG get_herbs] start_idx={start_idx}, col_count={col_count}",
          file=_sys.stderr)
    print(f"[DEBUG get_herbs] 약재 {len(herbs)}개:", file=_sys.stderr)
    for _n, _d in herbs:
        print(f"  약재명={_n!r}  1회투약량={_d!r}", file=_sys.stderr)

    return {"status": "ok", "herbs": herbs}


# ═══════════════════════════════════════════════════════════════════════════════
#  get_patient: 환자명 / 연락처 / 주소 반환
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_get_patient():
    """
    OKOSC 하단 EDIT 컨트롤(WindowsForms10.EDIT...)에서
    환자명 / 핸드폰 / 주소 / 한의원명 / 한의원전화를 win32gui로 읽어옵니다.

    컨트롤 텍스트 예시:
      배송방식: 택배     배송옵션:      발송일자:
      환자명: 최해경     전화번호:      핸드폰: 010-6291-1476
         주소: 서울시 영등포구 신길로 28길 25 101동 1301호 (신길동, 신길파크자이)
      한의원명: 오초한의원     전화번호: 02-429-8775
           주소: 서울 강동구 천호대로 1024 ...
    """
    import sys as _sys
    import ctypes
    import win32gui
    import win32con

    KEYWORDS = ["OKOCSTJS", "탕전실", "OKOSC", "OK처방", "처방프로그램"]

    # ── OKOSC 루트 hwnd 찾기 ─────────────────────────────────────────────────
    root_hwnd = None
    def _find_root(hwnd, _):
        nonlocal root_hwnd
        title = win32gui.GetWindowText(hwnd)
        cls   = win32gui.GetClassName(hwnd)
        if any(kw in title for kw in KEYWORDS) and "WindowsForms" in cls:
            root_hwnd = hwnd
    win32gui.EnumWindows(_find_root, None)
    if not root_hwnd:
        return {"status": "error", "message": "OKOSC 창을 찾을 수 없습니다"}

    # ── 자식창에서 환자정보 EDIT 컨트롤 찾기 ────────────────────────────────
    def _wm_text(hwnd):
        try:
            n = win32gui.SendMessage(hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
            if n == 0:
                return ""
            buf = ctypes.create_unicode_buffer(n + 4)
            win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, n + 4, buf)
            return buf.value
        except Exception:
            return ""

    info_text = ""
    def _find_edit(hwnd, _):
        nonlocal info_text
        if info_text:            # 이미 찾았으면 skip
            return
        cls = win32gui.GetClassName(hwnd)
        if "EDIT" not in cls.upper():
            return
        t = _wm_text(hwnd)
        if "환자명" in t and "핸드폰" in t:
            info_text = t
            print(f"[DEBUG get_patient] EDIT hwnd={hwnd}  cls={cls!r}  len={len(t)}",
                  file=_sys.stderr)
            print(f"[DEBUG get_patient] text={t!r}", file=_sys.stderr)

    win32gui.EnumChildWindows(root_hwnd, _find_edit, None)

    if not info_text:
        print("[DEBUG get_patient] 환자정보 EDIT 컨트롤 찾지 못함", file=_sys.stderr)
        return {"status": "ok", "name": "", "contact": "", "address": "",
                "clinic_name": "", "clinic_contact": ""}

    # ── 파싱 ─────────────────────────────────────────────────────────────────
    _NAME_RE    = _re.compile(r'환자명:\s*([가-힣]{2,10})')
    _PHONE_RE   = _re.compile(r'핸드폰:\s*(01[0-9][\s\-]?\d{3,4}[\s\-]?\d{4})')
    _ADDR_RE    = _re.compile(r'환자명:.*?\n\s*주소:\s*(.+?)(?:\n|$)', _re.DOTALL)
    _CLINIC_RE  = _re.compile(r'한의원명:\s*([^\s]+)')
    _CLINIC_TEL = _re.compile(r'한의원명:.*?전화번호:\s*([\d\-]+)', _re.DOTALL)

    m = _NAME_RE.search(info_text)
    name = m.group(1).strip() if m else ""

    m = _PHONE_RE.search(info_text)
    contact = m.group(1).strip() if m else ""

    m = _ADDR_RE.search(info_text)
    address = m.group(1).strip() if m else ""

    m = _CLINIC_RE.search(info_text)
    clinic_name = m.group(1).strip() if m else ""

    m = _CLINIC_TEL.search(info_text)
    clinic_contact = m.group(1).strip() if m else ""

    print(f"[DEBUG get_patient] 환자명={name!r}  핸드폰={contact!r}", file=_sys.stderr)
    print(f"[DEBUG get_patient] 주소={address!r}", file=_sys.stderr)
    print(f"[DEBUG get_patient] 한의원={clinic_name!r}  한의원전화={clinic_contact!r}",
          file=_sys.stderr)

    return {"status": "ok", "name": name, "contact": contact, "address": address,
            "clinic_name": clinic_name, "clinic_contact": clinic_contact}


# ═══════════════════════════════════════════════════════════════════════════════
#  get_dosage: 복용법 텍스트 반환
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_get_dosage():
    """
    Table[0] (DataItem 5개)의 item[4]에서 복용법 텍스트를 반환합니다.
    예: '20 첩 30 팩 (1일 2팩/ 10일/ 1팩 110cc)'
    Text descendants 순회는 C1TrueDBGrid1 때문에 무한대기를 일으키므로 사용하지 않습니다.
    """
    import sys as _sys

    def _cell_text(item):
        t = item.window_text().strip()
        if t:
            return t
        try:
            ch = item.children()
            if ch:
                return ch[0].window_text().strip()
        except Exception:
            pass
        return ""

    win = _okosc_win()
    dosage = ""

    tables = win.descendants(control_type="Table")
    print(f"[DEBUG get_dosage] Table 개수={len(tables)}", file=_sys.stderr)
    for ti, table in enumerate(tables):
        items = table.children(control_type="DataItem")
        texts = [_cell_text(it) for it in items]
        print(f"[DEBUG get_dosage] Table[{ti}] DataItem={len(items)}개  texts={texts}",
              file=_sys.stderr)
        if len(items) == 5 and texts[4]:
            dosage = texts[4]
            print(f"[DEBUG get_dosage] 복용법 발견: {dosage!r}", file=_sys.stderr)
            break

    return {"status": "ok", "dosage": dosage}


# ═══════════════════════════════════════════════════════════════════════════════
#  save_pdf: 출력 → PDF저장 클릭 → 파일 저장 대화상자 처리 → PDF 경로 반환
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_save_pdf():
    """
    OKOSC 출력 메뉴 → PDF저장을 클릭하고, 저장 대화상자에서
    argv[2]로 전달된 경로에 PDF를 저장합니다.

    argv[2]: 저장할 PDF 경로 (예: C:\\Users\\COM\\Downloads\\presc.pdf)
    반환: {"status": "ok", "pdf_path": "<경로>"}
    """
    import sys as _sys
    import os as _os
    import pyautogui
    from pywinauto.keyboard import send_keys
    from pywinauto import Desktop as _Desktop

    if len(sys.argv) < 3:
        return {"status": "error", "message": "PDF 저장 경로를 argv[2]로 전달하세요"}

    save_path = sys.argv[2]
    # 이미 존재하면 삭제 (이전 실행 잔여물)
    if _os.path.exists(save_path):
        _os.remove(save_path)

    win = _okosc_win()
    win.set_focus()
    time.sleep(0.3)

    # ── 출력 버튼 클릭 ───────────────────────────────────────────────────────
    clicked_menu = False
    for try_title in ["출력", "출  력"]:
        try:
            btn = win.child_window(title=try_title, control_type="Button")
            btn.click_input()
            clicked_menu = True
            break
        except Exception:
            pass
    if not clicked_menu:
        # pyautogui fallback: 화면 좌표 탐색
        try:
            all_btns = win.descendants(control_type="Button")
            for b in all_btns:
                if "출력" in b.window_text():
                    b.click_input()
                    clicked_menu = True
                    break
        except Exception:
            pass
    if not clicked_menu:
        return {"status": "error", "message": "출력 버튼을 찾을 수 없습니다"}

    time.sleep(0.5)
    print("[DEBUG save_pdf] 출력 버튼 클릭 완료", file=_sys.stderr)

    # ── 첩약보험(처방전) 메뉴 항목 클릭 ────────────────────────────────────
    # 팝업 메뉴는 최상위 창으로 나타남
    clicked_pdf = False
    for _ in range(10):
        try:
            desktop = _Desktop(backend="uia")
            for ctrl_type in ("MenuItem", "ListItem", "Button", "Text"):
                for w in desktop.windows():
                    for ci in w.descendants(control_type=ctrl_type):
                        txt = ci.window_text()
                        if "첩약보험" in txt and "처방전" in txt:
                            ci.click_input()
                            clicked_pdf = True
                            break
                    if clicked_pdf:
                        break
                if clicked_pdf:
                    break
            if clicked_pdf:
                break
        except Exception:
            pass
        time.sleep(0.2)

    if not clicked_pdf:
        send_keys("{ESCAPE}")
        return {"status": "error", "message": "첩약보험(처방전) 메뉴 항목을 찾을 수 없습니다"}

    time.sleep(1.0)
    print("[DEBUG save_pdf] 첩약보험(처방전) 메뉴 클릭 완료", file=_sys.stderr)

    # ── 파일 저장 대화상자 처리 ──────────────────────────────────────────────
    # Windows 표준 저장 대화상자 (제목: "다른 이름으로 저장" 또는 "Save As")
    dlg = None
    for _ in range(20):
        try:
            dlg = _Desktop(backend="uia").window(
                title_re=r"다른 이름으로 저장|Save As|저장"
            )
            dlg.wait("visible", timeout=2)
            break
        except Exception:
            pass
        time.sleep(0.3)

    if dlg is None:
        return {"status": "error", "message": "파일 저장 대화상자가 나타나지 않았습니다"}

    print("[DEBUG save_pdf] 저장 대화상자 발견", file=_sys.stderr)

    # 파일명 입력란에 경로 입력
    try:
        # 파일명 콤보/에디트 컨트롤 찾기
        name_edit = dlg.child_window(control_type="Edit", found_index=0)
        name_edit.set_text(save_path)
    except Exception:
        # pyautogui 직접 타이핑 fallback
        send_keys("^a")
        time.sleep(0.1)
        pyautogui.write(save_path, interval=0.02)

    time.sleep(0.3)
    send_keys("{ENTER}")
    time.sleep(0.5)

    # 덮어쓰기 확인 대화상자가 나타날 경우 Enter
    try:
        confirm = _Desktop(backend="uia").window(title_re=r".*확인.*|.*Confirm.*")
        confirm.wait("visible", timeout=2)
        send_keys("{ENTER}")
        time.sleep(0.5)
    except Exception:
        pass

    # ── PDF 생성 대기 ────────────────────────────────────────────────────────
    for _ in range(30):
        if _os.path.exists(save_path) and _os.path.getsize(save_path) > 1024:
            print(f"[DEBUG save_pdf] PDF 저장 완료: {save_path}", file=_sys.stderr)
            return {"status": "ok", "pdf_path": save_path}
        time.sleep(0.5)

    # 경로에 없어도 다운로드 폴더에서 최근 PDF 탐색 (앱이 경로 무시하는 경우 대비)
    import glob as _glob
    dl_dir = _os.path.dirname(save_path)
    pdfs = sorted(
        _glob.glob(_os.path.join(dl_dir, "*.pdf")),
        key=_os.path.getmtime,
        reverse=True
    )
    if pdfs and (time.time() - _os.path.getmtime(pdfs[0])) < 30:
        actual = pdfs[0]
        print(f"[DEBUG save_pdf] 대체 PDF 발견: {actual}", file=_sys.stderr)
        return {"status": "ok", "pdf_path": actual}

    return {"status": "error", "message": f"PDF 파일이 생성되지 않았습니다: {save_path}"}


# ═══════════════════════════════════════════════════════════════════════════════
#  get_presc_screenshot: 출력→첩약보험(처방전) → 처방전출력 창 캡처 → 닫기
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_get_presc_screenshot():
    """
    1. 출력 버튼 클릭
    2. 첩약보험(처방전) 메뉴 클릭
    3. 처방전출력 창이 열릴 때까지 대기 (최대 15초)
    4. 스크린샷을 argv[2] 경로에 PNG로 저장
    5. 창 닫기
    반환: {"status": "ok", "img_path": "..."}
    """
    import win32gui
    import win32con
    from PIL import ImageGrab
    from pywinauto.keyboard import send_keys
    from pywinauto import Desktop as _Desktop

    if len(sys.argv) < 3:
        return {"status": "error", "message": "이미지 저장 경로를 argv[2]로 전달하세요"}

    save_path = sys.argv[2]

    win = _okosc_win()
    win.set_focus()
    time.sleep(0.3)

    # ── 출력 버튼 클릭 ───────────────────────────────────────────────────────
    clicked_menu = False
    for try_title in ["출력", "출  력"]:
        try:
            btn = win.child_window(title=try_title, control_type="Button")
            btn.click_input()
            clicked_menu = True
            break
        except Exception:
            pass
    if not clicked_menu:
        try:
            for b in win.descendants(control_type="Button"):
                if "출력" in b.window_text():
                    b.click_input()
                    clicked_menu = True
                    break
        except Exception:
            pass
    if not clicked_menu:
        return {"status": "error", "message": "출력 버튼을 찾을 수 없습니다"}

    time.sleep(0.5)
    print("[DEBUG get_presc_screenshot] 출력 버튼 클릭 완료", file=sys.stderr)

    # ── 첩약보험(처방전) 메뉴 클릭 ──────────────────────────────────────────
    clicked_presc = False
    for _ in range(10):
        try:
            for ctrl_type in ("MenuItem", "ListItem", "Button", "Text"):
                for w in _Desktop(backend="uia").windows():
                    for ci in w.descendants(control_type=ctrl_type):
                        txt = ci.window_text()
                        if "첩약보험" in txt and "처방전" in txt:
                            ci.click_input()
                            clicked_presc = True
                            break
                    if clicked_presc:
                        break
                if clicked_presc:
                    break
            if clicked_presc:
                break
        except Exception:
            pass
        time.sleep(0.2)

    if not clicked_presc:
        send_keys("{ESCAPE}")
        return {"status": "error", "message": "첩약보험(처방전) 메뉴 항목을 찾을 수 없습니다"}

    print("[DEBUG get_presc_screenshot] 첩약보험(처방전) 클릭 완료", file=sys.stderr)

    # ── 처방전출력 창 대기 ───────────────────────────────────────────────────
    presc_hwnd = None
    for _ in range(30):
        def _find_presc(h, _):
            nonlocal presc_hwnd
            t = win32gui.GetWindowText(h)
            if "처방전출력" in t and win32gui.IsWindowVisible(h):
                presc_hwnd = h
        win32gui.EnumWindows(_find_presc, None)
        if presc_hwnd:
            break
        time.sleep(0.5)

    if not presc_hwnd:
        return {"status": "error", "message": "처방전출력 창이 열리지 않았습니다 (15초 대기)"}

    print(f"[DEBUG get_presc_screenshot] 처방전출력 창 발견 hwnd={presc_hwnd}", file=sys.stderr)

    win32gui.ShowWindow(presc_hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(presc_hwnd)
    time.sleep(0.8)

    # ── 스크린샷 저장 ────────────────────────────────────────────────────────
    left, top, right, bottom = win32gui.GetWindowRect(presc_hwnd)
    img = ImageGrab.grab(bbox=(left, top, right, bottom))
    img.save(save_path)
    print(f"[DEBUG get_presc_screenshot] 저장 완료: {save_path}", file=sys.stderr)

    # ── 창 닫기 ──────────────────────────────────────────────────────────────
    win32gui.PostMessage(presc_hwnd, win32con.WM_CLOSE, 0, 0)
    time.sleep(0.5)

    return {"status": "ok", "img_path": save_path}


# ═══════════════════════════════════════════════════════════════════════════════
#  setup_search: OKOSC 검색 설정 및 실행 (auto3.py 루프 전 1회 호출)
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_setup_search():
    """
    OKOSC 검색 조건 설정 및 검색 실행.
    - 검색기준: 진행상태
    - 날짜: 30일 전 ~ 오늘
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
    "save_pdf":             cmd_save_pdf,
    "get_presc_screenshot": cmd_get_presc_screenshot,
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
