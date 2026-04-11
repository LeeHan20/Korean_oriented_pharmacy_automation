@echo off
chcp 65001 > nul
title 자동화 2번 - 운송장 업로드
cd /d %~dp0

python --version > nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    pause
    exit /b 1
)

python auto2.py
pause
