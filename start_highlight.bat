@echo off
chcp 65001 > nul
cd /d "%~dp0"
title 配信ハイライト抽出ツール

echo 配信ハイライト抽出ツールを起動します...
echo.

REM --- Python を探す（python → py の順で試す） ---
set "PY="
python --version >nul 2>&1 && set "PY=python"
if not defined PY py --version >nul 2>&1 && set "PY=py"
if not defined PY goto NOPYTHON

REM --- 必要なファイルが揃っているか ---
if not exist "stream_highlight\requirements.txt" goto NOFILES

REM --- pip が使えるか ---
%PY% -m pip --version >nul 2>&1
if errorlevel 1 goto NOPIP

REM --- ライブラリが揃っていれば、そのまま起動する ---
%PY% -c "import fastapi, uvicorn, yt_dlp, imageio_ffmpeg" >nul 2>&1
if not errorlevel 1 goto RUN

echo 必要なライブラリをインストールします（初回のみ・数分かかります）
echo.
%PY% -m pip install -r stream_highlight\requirements.txt
if errorlevel 1 goto PIPFAIL

REM --- インストール直後にもう一度確認する ---
%PY% -c "import fastapi, uvicorn, yt_dlp, imageio_ffmpeg" >nul 2>&1
if errorlevel 1 goto IMPORTFAIL

:RUN
echo.
%PY% -m stream_highlight.server
if errorlevel 1 (
    echo.
    echo サーバーが異常終了しました。上のエラーを確認してください。
)
pause
exit /b 0


:NOPYTHON
echo ============================================================
echo  Python が見つかりませんでした。
echo.
echo  このPCにはボット用のPythonが入っているはずなので、
echo  PATHが通っていない可能性があります。
echo  コマンドプロンプトで次を試してください:
echo.
echo      python --version
echo      py --version
echo.
echo  どちらもエラーになる場合は https://www.python.org/downloads/
echo  からインストールし、"Add python.exe to PATH" に必ずチェックを入れてください。
echo ============================================================
pause
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
pause
exit /b 1


:NOPIP
echo ============================================================
echo  pip が使えません。次を実行してから、もう一度お試しください:
echo.
echo      %PY% -m ensurepip --upgrade
echo ============================================================
pause
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
echo.
echo  よくある原因:
echo   - ネットワーク / プロキシで接続がブロックされている
echo   - pip が古い（%PY% -m pip install --upgrade pip で更新できます）
echo ============================================================
pause
exit /b 1


:IMPORTFAIL
echo.
echo ============================================================
echo  インストールは終わりましたが、ライブラリを読み込めませんでした。
echo  複数のPythonが入っていて、別の環境に入った可能性があります。
echo.
echo  次の出力を貼ってもらえれば分かります:
echo.
echo      %PY% -c "import sys; print(sys.executable)"
echo      %PY% -m pip list
echo ============================================================
pause
exit /b 1
