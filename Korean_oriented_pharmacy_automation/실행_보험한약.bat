@echo off
chcp 65001 > nul
title 보험한약 자동화
cd /d %~dp0
python -X utf8 insurance_med.py
pause
