Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d C:\Users\somas\discord-claude-bot && python bot.py >> C:\Users\somas\discord-claude-bot\bot.log 2>&1", 0, False
