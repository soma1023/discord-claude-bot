@echo off
REM デスクトップのショートカットを、分かりやすい名前とアイコンに整える。
REM 中身は make_shortcut.ps1。

cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_shortcut.ps1"
echo.
pause
