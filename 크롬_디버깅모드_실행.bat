@echo off
chcp 65001 > nul
title 크롬 디버깅모드 실행

echo Chrome을 원격 디버깅 모드로 실행합니다...
echo 이 창은 닫지 마세요.
echo.

set CHROME_PATH=C:\Program Files\Google\Chrome\Application\chrome.exe
set PROFILE_DIR=%~dp0chrome_profile

if not exist "%CHROME_PATH%" (
    set CHROME_PATH=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe
)

if not exist "%CHROME_PATH%" (
    echo [오류] Chrome을 찾을 수 없습니다.
    echo Chrome 경로를 직접 수정하세요: %CHROME_PATH%
    pause
    exit /b 1
)

start "" "%CHROME_PATH%" --remote-debugging-port=9222 --user-data-dir="%PROFILE_DIR%"

echo Chrome이 실행되었습니다.
echo 로그인이 필요한 사이트에 로그인 후 자동화를 실행하세요.
echo.
echo 로그인 완료 후 이 창은 그대로 두세요.
pause
