@echo off
REM 黒い画面を出さずに起動する（VBScript を使わない版）。
REM
REM Windows 11 では VBScript が段階的に無効化されており、
REM start_highlight_silent.vbs をダブルクリックしても何も起きなくなる。
REM こちらは PowerShell 経由で隠して起動するので、その影響を受けない。
REM
REM 出力は stream_highlight\app.log に残る（bat 側が自分で書き出す）。
REM 終了するときは、画面右上の「終了」ボタンを押す。

cd /d "%~dp0"

REM -WindowStyle Hidden が「ウィンドウを出さない」指定。
REM 引数は --silent だけ。リダイレクトを渡さないので、入れ子の引用符が無く壊れにくい。
powershell -NoProfile -Command "Start-Process -FilePath 'start_highlight.bat' -ArgumentList '--silent' -WindowStyle Hidden -WorkingDirectory '%~dp0'"
if errorlevel 1 (
    echo.
    echo PowerShell で起動できませんでした。
    echo start_highlight.bat を直接ダブルクリックしてください（黒い画面は出ます）。
    pause
)
