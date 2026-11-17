@echo off
chcp 65001 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Iniciar_Examen.ps1" -Examen reto2 %*
if errorlevel 1 pause
