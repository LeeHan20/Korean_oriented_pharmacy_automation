@echo off
title Chrome Debug Mode

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="%~dp0chrome_profile"

echo Chrome started on port 9222.
echo Do NOT close this window.
pause
