@echo off
title Auto test - just test
cd /d %~dp0
python autotest.py
pause
