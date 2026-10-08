' 配信ハイライト抽出ツールを、黒い画面を出さずに起動する。
'
' 起動のたびに git pull で最新に更新されるので、作り直す手間がいらない。
' 出力は stream_highlight\app.log に残る。
' 終了するときは、画面右上の「終了」ボタンを押す。

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' このファイルがある場所で動かす。以降は相対パスで済むので、
' フォルダ名に空白が入っていても壊れない。
shell.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)

' 第2引数の 0 が「ウィンドウを出さない」指定
shell.Run "cmd /c start_highlight.bat --silent >> stream_highlight\app.log 2>&1", 0, False
