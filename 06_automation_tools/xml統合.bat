@echo off
setlocal

set "PS1=%TEMP%\xml_merge_gui_%RANDOM%.ps1"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$lines = [System.IO.File]::ReadAllLines('%~f0', [System.Text.Encoding]::UTF8);" ^
    "$found = $false; $buf = New-Object System.Collections.ArrayList;" ^
    "foreach ($l in $lines) { if ($found) { [void]$buf.Add($l) } elseif ($l -match '^#PSSTART') { $found = $true } };" ^
    "[System.IO.File]::WriteAllLines('%PS1%', $buf.ToArray(), [System.Text.UTF8Encoding]::new($true))"

if not exist "%PS1%" (
    echo [ERROR] スクリプトの抽出に失敗しました。
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%"

if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] PowerShell がエラーコード %ERRORLEVEL% で終了しました。
    pause
)

del "%PS1%" >nul 2>&1
exit /b

#PSSTART
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:allFiles = @()

# ============================================================
# Form
# ============================================================
$f = New-Object System.Windows.Forms.Form
$f.Text = 'XML Merge Tool'
$f.Size = New-Object System.Drawing.Size(570, 720)
$f.StartPosition = 'CenterScreen'
$f.FormBorderStyle = 'FixedDialog'
$f.MaximizeBox = $false
$f.Font = New-Object System.Drawing.Font('Meiryo UI', 9)

$y = 15
$boldFont = New-Object System.Drawing.Font('Meiryo UI', 9, [System.Drawing.FontStyle]::Bold)

# ============================================================
# Step 1: Input folder
# ============================================================
$lbl1 = New-Object System.Windows.Forms.Label
$lbl1.Text = '1  入力フォルダ（XMLファイルが入ったフォルダ）'
$lbl1.Location = New-Object System.Drawing.Point(15, $y)
$lbl1.AutoSize = $true
$lbl1.Font = $boldFont
$f.Controls.Add($lbl1)
$y += 24

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(15, $y)
$txtFolder.Size = New-Object System.Drawing.Size(420, 25)
$txtFolder.ReadOnly = $true
$f.Controls.Add($txtFolder)

$btnFolder = New-Object System.Windows.Forms.Button
$btnFolder.Text = '参照...'
$btnFolder.Location = New-Object System.Drawing.Point(445, $y)
$btnFolder.Size = New-Object System.Drawing.Size(90, 25)
$f.Controls.Add($btnFolder)
$y += 30

$lblFiles = New-Object System.Windows.Forms.Label
$lblFiles.Text = '検出ファイル: なし'
$lblFiles.Location = New-Object System.Drawing.Point(15, $y)
$lblFiles.AutoSize = $true
$lblFiles.ForeColor = 'Blue'
$f.Controls.Add($lblFiles)
$y += 22

# --- CheckedListBox（チェックで選択） ---
$chkList = New-Object System.Windows.Forms.CheckedListBox
$chkList.Location = New-Object System.Drawing.Point(15, $y)
$chkList.Size = New-Object System.Drawing.Size(520, 110)
$chkList.Font = New-Object System.Drawing.Font('Consolas', 8.5)
$chkList.CheckOnClick = $true
$f.Controls.Add($chkList)
$y += 115

# --- 全選択 / 全解除 ボタン ---
$btnAll = New-Object System.Windows.Forms.Button
$btnAll.Text = '全選択'
$btnAll.Location = New-Object System.Drawing.Point(15, $y)
$btnAll.Size = New-Object System.Drawing.Size(80, 25)
$f.Controls.Add($btnAll)

$btnNone = New-Object System.Windows.Forms.Button
$btnNone.Text = '全解除'
$btnNone.Location = New-Object System.Drawing.Point(100, $y)
$btnNone.Size = New-Object System.Drawing.Size(80, 25)
$f.Controls.Add($btnNone)

$lblChecked = New-Object System.Windows.Forms.Label
$lblChecked.Text = '選択中: 0 個'
$lblChecked.Location = New-Object System.Drawing.Point(200, ($y + 4))
$lblChecked.AutoSize = $true
$lblChecked.ForeColor = [System.Drawing.Color]::FromArgb(0, 100, 0)
$f.Controls.Add($lblChecked)
$y += 35

# ============================================================
# Step 2: Mode
# ============================================================
$lbl2 = New-Object System.Windows.Forms.Label
$lbl2.Text = '2  マージモード'
$lbl2.Location = New-Object System.Drawing.Point(15, $y)
$lbl2.AutoSize = $true
$lbl2.Font = $boldFont
$f.Controls.Add($lbl2)
$y += 24

$cmbMode = New-Object System.Windows.Forms.ComboBox
$cmbMode.Location = New-Object System.Drawing.Point(15, $y)
$cmbMode.Size = New-Object System.Drawing.Size(520, 25)
$cmbMode.DropDownStyle = 'DropDownList'
[void]$cmbMode.Items.Add('モード1: 子ノードを共通ルート下に結合（同構造向け）')
[void]$cmbMode.Items.Add('モード2: File要素で個別ラップ（異構造でも安全）')
$cmbMode.SelectedIndex = 0
$f.Controls.Add($cmbMode)
$y += 28

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = '※ 同構造→モード1 / 異構造→モード2 を推奨'
$lblHint.Location = New-Object System.Drawing.Point(15, $y)
$lblHint.AutoSize = $true
$lblHint.ForeColor = 'Gray'
$f.Controls.Add($lblHint)
$y += 28

# ============================================================
# Step 3: Root tag
# ============================================================
$lbl3 = New-Object System.Windows.Forms.Label
$lbl3.Text = '3  出力ルートタグ名'
$lbl3.Location = New-Object System.Drawing.Point(15, $y)
$lbl3.AutoSize = $true
$lbl3.Font = $boldFont
$f.Controls.Add($lbl3)
$y += 24

$txtRoot = New-Object System.Windows.Forms.TextBox
$txtRoot.Location = New-Object System.Drawing.Point(15, $y)
$txtRoot.Size = New-Object System.Drawing.Size(520, 25)
$txtRoot.Text = 'root'
$f.Controls.Add($txtRoot)
$y += 28

$lblRoot = New-Object System.Windows.Forms.Label
$lblRoot.Text = '※ フォルダ選択時に最初のXMLから自動検出（手動変更可）'
$lblRoot.Location = New-Object System.Drawing.Point(15, $y)
$lblRoot.AutoSize = $true
$lblRoot.ForeColor = 'Gray'
$f.Controls.Add($lblRoot)
$y += 28

# ============================================================
# Step 4: Output file
# ============================================================
$lbl4 = New-Object System.Windows.Forms.Label
$lbl4.Text = '4  出力先ファイル'
$lbl4.Location = New-Object System.Drawing.Point(15, $y)
$lbl4.AutoSize = $true
$lbl4.Font = $boldFont
$f.Controls.Add($lbl4)
$y += 24

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(15, $y)
$txtOut.Size = New-Object System.Drawing.Size(420, 25)
$f.Controls.Add($txtOut)

$btnOut = New-Object System.Windows.Forms.Button
$btnOut.Text = '参照...'
$btnOut.Location = New-Object System.Drawing.Point(445, $y)
$btnOut.Size = New-Object System.Drawing.Size(90, 25)
$f.Controls.Add($btnOut)
$y += 42

# ============================================================
# Merge button
# ============================================================
$btnMerge = New-Object System.Windows.Forms.Button
$btnMerge.Text = [char]0x25B6 + ' マージ実行'
$btnMerge.Location = New-Object System.Drawing.Point(15, $y)
$btnMerge.Size = New-Object System.Drawing.Size(520, 42)
$btnMerge.BackColor = [System.Drawing.Color]::FromArgb(67, 160, 71)
$btnMerge.ForeColor = 'White'
$btnMerge.FlatStyle = 'Flat'
$btnMerge.Font = New-Object System.Drawing.Font('Meiryo UI', 12, [System.Drawing.FontStyle]::Bold)
$f.Controls.Add($btnMerge)
$y += 52

# ============================================================
# Status label
# ============================================================
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = ''
$lblStatus.Location = New-Object System.Drawing.Point(15, $y)
$lblStatus.Size = New-Object System.Drawing.Size(520, 50)
$f.Controls.Add($lblStatus)

# ============================================================
# Helper: update checked count
# ============================================================
function Update-CheckedCount {
    $cnt = $chkList.CheckedItems.Count
    $lblChecked.Text = "選択中: $cnt 個"
}

# ============================================================
# Event: Browse input folder
# ============================================================
$btnFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'XMLファイルが入ったフォルダを選択してください'
    if ($dlg.ShowDialog() -eq 'OK') {
        $txtFolder.Text = $dlg.SelectedPath
        $chkList.Items.Clear()
        $files = Get-ChildItem -Path $dlg.SelectedPath -Filter '*.xml' -File | Sort-Object Name
        $script:allFiles = @($files)
        foreach ($fi in $files) {
            [void]$chkList.Items.Add($fi.Name, $false)
        }
        $lblFiles.Text = "検出ファイル: $($files.Count) 個"
        Update-CheckedCount
        if ($files.Count -gt 0) {
            try {
                [xml]$x = Get-Content -Path $files[0].FullName -Encoding UTF8
                $txtRoot.Text = $x.DocumentElement.Name
            } catch {}
        }
    }
})

# ============================================================
# Event: Check changed -> update count
# ============================================================
$chkList.Add_ItemCheck({
    # ItemCheck fires before state changes, so adjust manually
    $cnt = $chkList.CheckedItems.Count
    if ($_.NewValue -eq 'Checked') { $cnt++ } else { $cnt-- }
    $lblChecked.Text = "選択中: $cnt 個"
})

# ============================================================
# Event: Select All / Deselect All
# ============================================================
$btnAll.Add_Click({
    for ($i = 0; $i -lt $chkList.Items.Count; $i++) {
        $chkList.SetItemChecked($i, $true)
    }
    Update-CheckedCount
})

$btnNone.Add_Click({
    for ($i = 0; $i -lt $chkList.Items.Count; $i++) {
        $chkList.SetItemChecked($i, $false)
    }
    Update-CheckedCount
})

# ============================================================
# Event: Browse output file
# ============================================================
$btnOut.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'XML Files (*.xml)|*.xml|All Files (*.*)|*.*'
    $dlg.FileName = 'combined.xml'
    $dlg.Title = '出力先を選択してください'
    if ($dlg.ShowDialog() -eq 'OK') {
        $txtOut.Text = $dlg.FileName
    }
})

# ============================================================
# Event: Merge execution
# ============================================================
$btnMerge.Add_Click({
    $lblStatus.ForeColor = 'Black'
    $lblStatus.Text = '処理中...'
    $f.Refresh()

    # Get checked file names
    $checkedNames = @()
    for ($i = 0; $i -lt $chkList.Items.Count; $i++) {
        if ($chkList.GetItemChecked($i)) {
            $checkedNames += $chkList.Items[$i]
        }
    }

    if ($checkedNames.Count -eq 0) {
        $lblStatus.ForeColor = 'Red'
        $lblStatus.Text = 'エラー: マージするファイルにチェックを入れてください'
        return
    }

    if ($checkedNames.Count -lt 2) {
        $lblStatus.ForeColor = 'Red'
        $lblStatus.Text = 'エラー: 2個以上のファイルにチェックを入れてください'
        return
    }

    $outPath = $txtOut.Text.Trim()
    if (-not $outPath) {
        $lblStatus.ForeColor = 'Red'
        $lblStatus.Text = 'エラー: 出力先を指定してください'
        return
    }

    # Build selected file list (full path)
    $selectedFiles = @()
    foreach ($fi in $script:allFiles) {
        if ($checkedNames -contains $fi.Name) {
            $selectedFiles += $fi
        }
    }

    $modeIdx = $cmbMode.SelectedIndex
    $rootTag = $txtRoot.Text.Trim()
    if (-not $rootTag) { $rootTag = 'root' }

    try {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine('<?xml version="1.0" encoding="utf-8"?>')
        [void]$sb.AppendLine("<$rootTag>")

        foreach ($fi in $selectedFiles) {
            [xml]$doc = Get-Content -Path $fi.FullName -Encoding UTF8

            if ($modeIdx -eq 1) {
                # Mode 2: wrap each file
                $fname = $fi.Name
                [void]$sb.AppendLine("  <File name=`"$fname`">")
                [void]$sb.AppendLine('    ' + $doc.DocumentElement.OuterXml)
                [void]$sb.AppendLine('  </File>')
            } else {
                # Mode 1: append children
                [void]$sb.AppendLine($doc.DocumentElement.InnerXml)
            }
        }

        [void]$sb.AppendLine("</$rootTag>")
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($outPath, $sb.ToString(), $utf8NoBom)

        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 0)
        $lblStatus.Text = "完了！ $($selectedFiles.Count) ファイルをマージ → $outPath"
    } catch {
        $lblStatus.ForeColor = 'Red'
        $lblStatus.Text = "エラー: $($_.Exception.Message)"
    }
})

# ============================================================
# Show form
# ============================================================
[void]$f.ShowDialog()
