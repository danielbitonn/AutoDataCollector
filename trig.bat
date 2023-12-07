@echo off
cd %~dp0
powershell -ExecutionPolicy Bypass -File "%~dp0SendWeeklyEmail.ps1"
