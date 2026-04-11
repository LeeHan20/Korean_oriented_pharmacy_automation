@echo off
chcp 65001 > nul
title 패키지 설치
cd /d %~dp0

echo 필요한 패키지를 설치합니다...
echo.

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo 설치가 완료되었습니다.
pause
