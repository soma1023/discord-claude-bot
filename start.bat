@echo off
chcp 65001 > nul
cd /d %~dp0
echo Discord Claude Bot を起動します...

:: 依存パッケージのインストール
pip install -r requirements.txt -q

:: Bot起動
python bot.py
pause
