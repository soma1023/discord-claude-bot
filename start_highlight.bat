@echo off
chcp 65001 > nul
cd /d "%~dp0"
title 配信ハイライト抽出ツール

REM --silent を付けて呼ばれたときは、一時停止せずに終わる（非表示起動用）
set "QUIET="
if /i "%~1"=="--silent" set "QUIET=1"

REM 非表示で動かすと画面に何も出ないので、自分でログを取る。
REM 呼び出し側に書き出しを任せると、起動方法ごとに書き方が増えて壊れやすい。
REM SH_LOGGING は二重に潜らないための目印。
if defined QUIET if not defined SH_LOGGING (
    set "SH_LOGGING=1"
    if not exist "%~dp0stream_highlight" mkdir "%~dp0stream_highlight"
    call "%~f0" --silent >> "%~dp0stream_highlight\app.log" 2>&1
    exit /b
)

REM git が認証を聞いてくると、非表示のまま永久に待ち続けて「開かない」になる。
REM 聞かずに失敗させて、ログに残す。
set "GIT_TERMINAL_PROMPT=0"

if not defined QUIET echo 配信ハイライト抽出ツールを起動します...
if not defined QUIET echo.

REM --- Python を探す（python → py の順で試す） ---
set "PY="
python --version >nul 2>&1 && set "PY=python"
if not defined PY py --version >nul 2>&1 && set "PY=py"
if not defined PY goto NOPYTHON

REM --- 更新があれば取り込む ---
git rev-parse --git-dir >nul 2>&1
if errorlevel 1 goto SKIPPULL
git diff --quiet
if errorlevel 1 goto DIRTY
echo 更新を確認しています...
git pull --ff-only
goto SKIPPULL

:DIRTY
echo 手元に変更があるため、更新の取り込みは行いません。
echo （元に戻すには git stash を実行してください）

:SKIPPULL

REM --- 必要なファイルが揃っているか ---
if not exist "stream_highlight\requirements.txt" goto NOFILES

REM --- pip が使えるか ---
%PY% -m pip --version >nul 2>&1
if errorlevel 1 goto NOPIP

REM --- ライブラリが揃っていれば、そのまま起動する ---
%PY% -c "import fastapi, uvicorn, yt_dlp" >nul 2>&1
if not errorlevel 1 goto RUN

echo 必要なライブラリをインストールします（初回のみ・数分かかります）
echo.
%PY% -m pip install -r stream_highlight\requirements.txt
if errorlevel 1 goto PIPFAIL

%PY% -c "import fastapi, uvicorn, yt_dlp" >nul 2>&1
if errorlevel 1 goto IMPORTFAIL

:RUN
if not defined QUIET echo.
%PY% -m stream_highlight.server
if errorlevel 1 (
    echo.
    echo サーバーが異常終了しました。上のエラーを確認してください。
    if defined QUIET call :ALERT
)
if not defined QUIET pause
exit /b 0


:NOPYTHON
echo ============================================================
echo  Python が見つかりませんでした。
echo.
echo  コマンドプロンプトで次を試してください:
echo.
echo      python --version
echo      py --version
echo.
echo  どちらもエラーになる場合は https://www.python.org/downloads/
echo  からインストールし、"Add python.exe to PATH" に必ずチェックを入れてください。
echo ============================================================
if defined QUIET call :ALERT
if not defined QUIET pause
exit /b 1


:NOFILES
echo ============================================================
echo  stream_highlight フォルダが見つかりません。
echo.
echo  ブランチの取得がまだ済んでいないようです。
echo  このフォルダで次を実行してください:
echo.
echo      git fetch origin
echo      git checkout claude/stream-highlight-extractor-lzeupt
echo ============================================================
if defined QUIET call :ALERT
if not defined QUIET pause
exit /b 1


:NOPIP
echo ============================================================
echo  pip が使えません。次を実行してから、もう一度お試しください:
echo.
echo      %PY% -m ensurepip --upgrade
echo ============================================================
if defined QUIET call :ALERT
if not defined QUIET pause
exit /b 1


:PIPFAIL
echo.
echo ============================================================
echo  ライブラリのインストールに失敗しました。
echo  原因は上に表示されているエラーに出ています。
echo.
echo  詳しいログを install_error.log に保存します...
%PY% -m pip install -r stream_highlight\requirements.txt > install_error.log 2>&1
echo.
echo  保存先: %CD%\install_error.log
echo  このファイルの中身を貼ってもらえれば原因が分かります。
echo ============================================================
if defined QUIET call :ALERT
if not defined QUIET pause
exit /b 1


:IMPORTFAIL
echo.
echo ============================================================
echo  インストールは終わりましたが、ライブラリを読み込めませんでした。
echo  複数のPythonが入っていて、別の環境に入った可能性があります。
echo.
echo      %PY% -c "import sys; print(sys.executable)"
echo      %PY% -m pip list
echo ============================================================
if defined QUIET call :ALERT
if not defined QUIET pause
exit /b 1

:ALERT
REM 非表示起動では画面に何も出ないため、ログをメモ帳で開いて気づけるようにする。
start "" notepad "%~dp0stream_highlight\app.log"
exit /b 0
