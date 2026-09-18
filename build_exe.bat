@echo off
chcp 65001 > nul
cd /d "%~dp0"
title 配信ハイライト抽出ツール ビルド

echo 配信ハイライト抽出ツールを exe にまとめます。
echo 初回は10分ほどかかります。
echo.

set "PY="
python --version >nul 2>&1 && set "PY=python"
if not defined PY py --version >nul 2>&1 && set "PY=py"
if not defined PY (
    echo Python が見つかりません。https://www.python.org/downloads/ から入れてください。
    pause
    exit /b 1
)

echo 必要なものを用意します...
%PY% -m pip install -q --upgrade pyinstaller
if errorlevel 1 goto FAIL
%PY% -m pip install -q -r stream_highlight\requirements.txt
if errorlevel 1 goto FAIL

REM 配布物でもバージョンが分かるように、コミットIDを埋め込む
git rev-parse --short HEAD > stream_highlight\_build_id.txt 2>nul
if errorlevel 1 echo unknown > stream_highlight\_build_id.txt
set /p BUILT=<stream_highlight\_build_id.txt
echo このコードでまとめます: %BUILT%

echo.
echo まとめています...
%PY% -m PyInstaller --noconfirm --clean StreamHighlight.spec
if errorlevel 1 goto FAIL

echo.
echo ============================================================
echo  できあがりました（コード: %BUILT%）
echo.
echo   %CD%\dist\StreamHighlight\StreamHighlight.exe
echo.
echo  ※ 以前に別の場所へコピーしていた場合は、そちらは古いままです。
echo     この dist\StreamHighlight で置き換えてください。
echo     アプリの「解析の設定」に出るコードが %BUILT% なら最新です。
echo.
echo  dist\StreamHighlight フォルダごとどこへ移してもそのまま動きます。
echo  ショートカットをデスクトップに作っておくと楽です。
echo  チャットの保存先は、そのフォルダの中の data です。
echo ============================================================
start "" "%CD%\dist\StreamHighlight"
pause
exit /b 0

:FAIL
echo.
echo ビルドに失敗しました。上のエラーを確認してください。
pause
exit /b 1
