"""
한약국 자동화 프로그램 - 공통 설정 파일
실제 환경에 맞게 아래 값들을 수정하세요.
"""

import os
from pathlib import Path

# ─── 로컬 경로 ───────────────────────────────────────────────────────────────

# 다운로드 폴더 (Windows 기본값)
DOWNLOAD_DIR = str(Path.home() / "Downloads")

# 익산대장 파일이 있는 폴더 경로 (실제 경로로 수정 필요)
IKSAN_FILE_DIR = r"E:\한약국\옹기한약서류"

# ─── Chrome 원격 디버깅 설정 ──────────────────────────────────────────────────
# Chrome을 '크롬_디버깅모드_실행.bat'으로 실행해야 합니다.
CHROME_DEBUG_PORT = 9222

# ─── 사이트 URL ──────────────────────────────────────────────────────────────
HOMEPAGE_URL = "https://ongkihanyak.co.kr/RAD/rankup_index/main.html"
ROSEN_URL = "https://logis.ilogen.com/common/html/main.html"

# ─── 지점 / 주문 설정 ────────────────────────────────────────────────────────
BRANCH_NAME = "익산점"
ORDER_STATUS = "조제중"

# ─── OKOSC 프로그램 설정 ─────────────────────────────────────────────────────
# OKOSC 창 제목에 포함된 키워드 (inspect.exe로 실제 창 제목 확인 후 수정)
OKOSC_WINDOW_KEYWORDS = ["OKOCSTJS", "탕전실", "OKOSC", "OK처방", "처방프로그램"]

# 처방전송일자 검색 기간 (일)
DATE_RANGE_DAYS = 7

# ─── 택배 고정값 ─────────────────────────────────────────────────────────────
DELIVERY_F_VALUE = 1         # F열: 수량
DELIVERY_G_VALUE = "4400"    # G열: 금액
DELIVERY_H_VALUE = "한약"    # H열: 품목

# ─── 보험한약 파일 설정 ───────────────────────────────────────────────────────
# 첩약보험약값대장.xlsx 파일 경로 (실제 경로로 수정 필요)
INSURANCE_WORKBOOK_PATH = r"E:\한약국\첩약시범사업\첩약보험약값대장_테스트사본.xlsx"

# ─── 32비트 Python 경로 ───────────────────────────────────────────────────────
# OKOSC(32비트 앱) 자동화에 사용. 설치 후 실제 경로로 수정하세요.
# 예: C:\Python313-32\python.exe  또는
#     C:\Users\COM\AppData\Local\Programs\Python\Python313-32\python.exe
PYTHON32_PATH = r"C:\Users\COM\AppData\Local\Programs\Python\Python313-32\python.exe"

# ─── autocode DB ────────────────────────────────────────────────────────────
# 처방번호 ↔ autocode(4자리) 매핑 DB (auto1에서 생성, auto3에서 참조)
AUTOCODE_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "autocode_db.json")

# 동일주소 병합 DB: {대표 처방번호: [합쳐진 처방번호들]}
# auto1에서 생성, auto2에서 참조 (대표 1건 발송 → 나머지도 택배발송 처리)
SAME_ADDRESS_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "same_address_db.json")

# ─── 타임아웃 설정 (초) ──────────────────────────────────────────────────────
DOWNLOAD_TIMEOUT = 60        # 파일 다운로드 최대 대기 시간
PAGE_LOAD_TIMEOUT = 30       # 페이지 로딩 최대 대기 시간
OKOSC_WAIT_TIMEOUT = 30      # OKOSC 응답 최대 대기 시간
