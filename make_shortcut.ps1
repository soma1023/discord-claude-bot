# デスクトップのショートカットを、分かりやすい名前とアイコンに整える。
#
#   make_shortcut.cmd をダブルクリックすると、これが走る。
#
# 既にあるショートカット（start_hidden.cmd を指しているもの）は、名前を
# 変えて使い回す。消さずに残すのは、取り違えて別のものを消さないため。

$ErrorActionPreference = 'Stop'

$root   = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = Join-Path $root 'start_hidden.cmd'
$icon   = Join-Path $root 'stream_highlight\static\icon.ico'
$name   = '配信ハイライト抽出'
$note   = '配信のチャットから見せ場の候補を探す'

if (-not (Test-Path $target)) {
    Write-Host "start_hidden.cmd が見つかりません: $target"
    Write-Host "git pull で更新を取り込んでから、もう一度実行してください。"
    exit 1
}
if (-not (Test-Path $icon)) {
    Write-Host "アイコンが見つかりません: $icon"
    Write-Host "git pull で更新を取り込んでから、もう一度実行してください。"
    exit 1
}

# デスクトップの場所は .NET から取る。OneDrive に移していても正しく出る。
$desktop = [Environment]::GetFolderPath('Desktop')
$link    = Join-Path $desktop ($name + '.lnk')

try {
    $shell = New-Object -ComObject WScript.Shell
} catch {
    Write-Host "ショートカットを作る仕組みが使えませんでした。"
    Write-Host "手で変える場合:"
    Write-Host "  1. ショートカットを選んで F2 → 「$name」"
    Write-Host "  2. 右クリック → プロパティ → アイコンの変更 → 次を貼る:"
    Write-Host "     $icon"
    exit 1
}

# 同じものを指している古いショートカットがあれば、名前を変えて引き継ぐ
if (-not (Test-Path $link)) {
    foreach ($file in Get-ChildItem $desktop -Filter *.lnk) {
        if ($file.FullName -eq $link) { continue }
        try { $found = $shell.CreateShortcut($file.FullName) } catch { continue }
        if ($found.TargetPath -eq $target) {
            Rename-Item $file.FullName -NewName ($name + '.lnk')
            Write-Host "名前を変えました: $($file.Name) → $name"
            break
        }
    }
}

$sc = $shell.CreateShortcut($link)
$sc.TargetPath       = $target
$sc.WorkingDirectory = $root
$sc.IconLocation     = "$icon,0"
$sc.Description      = $note
$sc.Save()

Write-Host ""
Write-Host "できました: $link"
Write-Host "  名前      : $name"
Write-Host "  アイコン  : コメント密度のグラフ（見せ場の棒が緑）"
Write-Host ""

# 動かなくなった VBScript 版のショートカットが残っていたら知らせる（消さない）
$old = Join-Path $root 'start_highlight_silent.vbs'
foreach ($file in Get-ChildItem $desktop -Filter *.lnk) {
    try { $found = $shell.CreateShortcut($file.FullName) } catch { continue }
    if ($found.TargetPath -eq $old) {
        Write-Host "※ 古い VBScript 版のショートカットが残っています: $($file.Name)"
        Write-Host "  Windows 11 では動かないので、消して構いません。"
    }
}
