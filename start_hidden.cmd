@echo off
REM 黒い画面を出さずに起動する。
REM
REM pythonw.exe はコンソールを持たない Python なので、黒い画面が一瞬も出ない。
REM VBScript も PowerShell も経由しないため、それらが無効化されても影響が無い。
REM
REM 更新・ライブラリの導入・起動は run_app.py が全部やる。
REM 記録は stream_highlight\app.log。失敗したらメモ帳で勝手に開く。

cd /d "%~dp0"

where pythonw >nul 2>&1 && goto PYTHONW
where pyw >nul 2>&1 && goto PYW

echo ============================================================
echo  pythonw が見つかりませんでした。
echo  黒い画面が出る形で起動します（これはそのまま使えます）。
echo ============================================================
echo.
call "%~dp0start_highlight.bat"
exit /b

:PYTHONW
start "" pythonw.exe "%~dp0run_app.py" --update
exit /b

:PYW
start "" pyw.exe "%~dp0run_app.py" --update
exit /b
