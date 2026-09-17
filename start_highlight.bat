@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo 配信ハイライト抽出ツールを起動します...

python -c "import fastapi, uvicorn, yt_dlp" 2>nul
if errorlevel 1 (
    echo 必要なライブラリをインストールします...
    python -m pip install -r stream_highlight\requirements.txt
    if errorlevel 1 (
        echo.
        echo インストールに失敗しました。Pythonが入っているか確認してください。
        pause
        exit /b 1
    )
)

python -m stream_highlight.server
pause
