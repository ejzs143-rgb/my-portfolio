Attribute VB_Name = "Module4"
'==============================================================================
' Module4: RAWデータ生成 + PPT自動生成
'
' 【概要】
'   インボイスリスト（コアプロセスチェック対象インボイスリスト_YYYYMM.xlsx）から
'   RAWデータ形式のExcelを生成し、PPTの各スライドに自動入力する
'
' 【使い方】
'   ■ RAWデータ生成
'     Alt+F8 → Main_RAWデータ生成 → 実行
'     → 同フォルダ内の全インボイスリストを自動収集
'     → 保存先ダイアログでファイル名を指定して保存
'
'   ■ PPT自動生成
'     Alt+F8 → Main_PPT自動生成 → 実行
'     → ① PPTファイルを選択
'     → ② 法人を番号で選択
'     → ③ 店舗を番号で選択
'     → P4・P6・P8・P10が自動入力されたPPTが別名保存される
'
' 【前提条件】
'   ・このツール（.xlsm）と同じフォルダにインボイスリストを置く
'   ・PPT自動生成を実行する前に、RAWデータシートが必要
'     （Main_RAWデータ生成で生成したExcelを開き、
'       RAWデータシートをこのブックにコピーするか、
'       先にMain_RAWデータ生成を実行してRAWデータを用意する）
'   ・PPTのP4テーブル列1：法人名、P6テーブル列1：店舗名を事前に入力しておく
'   ・PPTのP8・P10の「ExampleBrand B ○○○○」が店舗名のプレースホルダー
'==============================================================================

Option Explicit

'==============================================================================
' ① RAWデータ生成（メイン）
'==============================================================================
Sub Main_RAWデータ生成()

    Dim folderPath  As String
    Dim fileName    As String
    Dim wbInv       As Workbook
    Dim wsTop3      As Worksheet
    Dim wsMaster    As Worksheet
    Dim critCodes   As Object
    Dim rawData()   As Variant
    Dim rawCount    As Long
    Dim lastRow     As Long
    Dim r           As Long
    Dim ci          As Integer
    Dim yyyymm      As String

    ' Top3(結果入り)シートの列定義
    ' 列8,12,16,20,24 = エラーコード①〜⑤
    ' 列10,14,18,22,26 = エラー内容①〜⑤
    ' 列11,15,19,23,27 = 追加コメント①〜⑤
    Dim codeCols(4)    As Integer
    Dim contentCols(4) As Integer
    Dim commentCols(4) As Integer
    codeCols(0) = 8:  contentCols(0) = 10: commentCols(0) = 11
    codeCols(1) = 12: contentCols(1) = 14: commentCols(1) = 15
    codeCols(2) = 16: contentCols(2) = 18: commentCols(2) = 19
    codeCols(3) = 20: contentCols(3) = 22: commentCols(3) = 23
    codeCols(4) = 24: contentCols(4) = 26: commentCols(4) = 27

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '----------------------------------------------------------------------
    ' STEP1: マスタ_重大エラーシートから重大コードを読み込む
    '----------------------------------------------------------------------
    Set critCodes = CreateObject("Scripting.Dictionary")
    critCodes.CompareMode = 1  ' 大文字小文字を区別しない

    On Error Resume Next
    Set wsMaster = ThisWorkbook.Sheets("マスタ_重大エラー")
    On Error GoTo ErrHandler

    If Not wsMaster Is Nothing Then
        Dim mi As Long
        For mi = 2 To 20
            Dim mc As String
            mc = Trim(CStr(wsMaster.Cells(mi, 1).Value))
            If mc = "" Then Exit For
            ' 数字のみ抽出（"E1223"や"エラー1223"にも対応）
            mc = ExtractNumbers_M4(mc)
            If mc <> "" And Not critCodes.Exists(mc) Then
                critCodes.Add mc, True
            End If
        Next mi
    Else
        MsgBox "「マスタ_重大エラー」シートが見つかりません。" & vbCrLf & _
               "重大フラグなしでRAWデータを生成します。", vbExclamation
    End If

    '----------------------------------------------------------------------
    ' STEP2: 同フォルダのインボイスリストを収集
    '----------------------------------------------------------------------
    folderPath = ThisWorkbook.Path & "\"
    Dim invFiles() As String
    Dim invCount   As Long
    invCount = 0
    ReDim invFiles(0)

    fileName = Dir(folderPath & "コアプロセスチェック対象インボイスリスト_*.xlsx")
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            ReDim Preserve invFiles(invCount)
            invFiles(invCount) = fileName
            invCount = invCount + 1
        End If
        fileName = Dir()
    Loop

    If invCount = 0 Then
        MsgBox "インボイスリストが見つかりませんでした。" & vbCrLf & vbCrLf & _
               "ファイル名の形式：コアプロセスチェック対象インボイスリスト_YYYYMM.xlsx" & vbCrLf & _
               "このツールと同じフォルダに置いてください。", vbExclamation
        GoTo CleanUp
    End If

    MsgBox invCount & "ヶ月分のインボイスリストを検出しました。処理を開始します。", vbInformation

    '----------------------------------------------------------------------
    ' STEP3: 各ファイルを読み込んでRAWデータ配列に蓄積
    '----------------------------------------------------------------------
    rawCount = 0
    ReDim rawData(0 To 60000, 0 To 9)

    Dim fi As Long
    For fi = 0 To invCount - 1
        Application.StatusBar = "読み込み中... (" & (fi + 1) & "/" & invCount & ") " & invFiles(fi)

        yyyymm = ExtractYYYYMM_M4(invFiles(fi))
        If yyyymm = "" Then
            MsgBox "ファイル名からYYYYMMを取得できませんでした：" & invFiles(fi) & vbCrLf & _
                   "スキップします。", vbExclamation
            GoTo NextFile_M4
        End If

        Set wbInv = Nothing
        On Error Resume Next
        Set wbInv = Workbooks.Open(folderPath & invFiles(fi), ReadOnly:=True, UpdateLinks:=False)
        On Error GoTo ErrHandler
        If wbInv Is Nothing Then GoTo NextFile_M4

        ' Top3(結果入り)シートを探す
        Set wsTop3 = Nothing
        Dim ws As Worksheet
        For Each ws In wbInv.Worksheets
            If ws.Name = "Top3(結果入り)" Then
                Set wsTop3 = ws
                Exit For
            End If
        Next ws

        If wsTop3 Is Nothing Then
            wbInv.Close False
            GoTo NextFile_M4
        End If

        lastRow = wsTop3.Cells(wsTop3.Rows.Count, 5).End(xlUp).Row

        For r = 4 To lastRow
            Dim roNum  As String
            Dim dNo    As String
            Dim dName  As String
            Dim vin    As String

            roNum = Trim(CStr(wsTop3.Cells(r, 5).Value))
            dNo   = Trim(CStr(wsTop3.Cells(r, 2).Value))
            dName = Trim(CStr(wsTop3.Cells(r, 3).Value))
            vin   = Trim(CStr(wsTop3.Cells(r, 6).Value))

            If roNum = "" Or dNo = "" Then GoTo NextRow_M4

            ' エラーコード①〜⑤を1行ずつ縦展開
            For ci = 0 To 4
                Dim eCode    As String
                Dim eContent As String
                Dim eComment As String

                eCode = Trim(CStr(wsTop3.Cells(r, codeCols(ci)).Value))
                eCode = ExtractNumbers_M4(eCode)

                ' 0・空・0000 はスキップ
                If eCode = "" Or eCode = "0" Or eCode = "0000" Then GoTo NextCode_M4

                eContent = Trim(CStr(wsTop3.Cells(r, contentCols(ci)).Value))
                eComment = Trim(CStr(wsTop3.Cells(r, commentCols(ci)).Value))

                ' 重大フラグ判定
                Dim flagStr As String
                If critCodes.Count > 0 Then
                    If critCodes.Exists(eCode) Then
                        flagStr = "重大"
                    Else
                        flagStr = "軽微"
                    End If
                Else
                    flagStr = ""
                End If

                ' 配列に1行追加
                ' 列順: 月, ディーラーNo, ディーラー名, RO番号, VIN,
                '        エラーコード, エラー内容, 追加コメント, サブカテゴリ, 重大フラグ
                rawData(rawCount, 0) = CLng(yyyymm)
                rawData(rawCount, 1) = CLng(dNo)
                rawData(rawCount, 2) = dName
                rawData(rawCount, 3) = roNum
                rawData(rawCount, 4) = vin
                rawData(rawCount, 5) = CLng(eCode)
                rawData(rawCount, 6) = eContent
                rawData(rawCount, 7) = eComment
                rawData(rawCount, 8) = ""        ' サブカテゴリ（将来用）
                rawData(rawCount, 9) = flagStr
                rawCount = rawCount + 1

                If rawCount >= 60000 Then
                    MsgBox "データが60,000件を超えました。処理を中断します。", vbExclamation
                    wbInv.Close False
                    GoTo WriteOut_M4
                End If

NextCode_M4:
            Next ci
NextRow_M4:
        Next r

        wbInv.Close False
NextFile_M4:
    Next fi

    If rawCount = 0 Then
        MsgBox "有効なエラーデータが見つかりませんでした。", vbExclamation
        GoTo CleanUp
    End If

WriteOut_M4:
    '----------------------------------------------------------------------
    ' STEP4: 保存先を選択して書き出し
    '----------------------------------------------------------------------
    Dim savePath As String
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=ThisWorkbook.Path & "\コアプロセスチェック_アクション用データ_" & _
                         Format(Now, "yyyymmdd") & ".xlsx", _
        FileFilter:="Excelファイル (*.xlsx),*.xlsx", _
        Title:="RAWデータの保存先を選択")

    If savePath = "False" Then GoTo CleanUp

    Application.ScreenUpdating = True
    Application.StatusBar = "ファイル書き出し中..."
    Application.ScreenUpdating = False

    ' 新規ブック作成
    Dim wbNew As Workbook
    Set wbNew = Workbooks.Add
    Dim wsRaw As Worksheet
    Set wsRaw = wbNew.Sheets(1)
    wsRaw.Name = "RAWデータ"

    ' ヘッダー書き込み
    Dim headers As Variant
    headers = Array("月", "ディーラーNo", "ディーラー名", "RO番号", "VIN", _
                    "エラーコード", "エラー内容", "追加コメント", "サブカテゴリ", "重大フラグ")
    Dim hi As Integer
    For hi = 0 To 9
        wsRaw.Cells(1, hi + 1).Value = headers(hi)
    Next hi

    ' ヘッダー装飾
    With wsRaw.Range("A1:J1")
        .Font.Bold = True
        .Font.Name = "Arial"
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' データ一括書き出し
    wsRaw.Range("A2").Resize(rawCount, 10).Value = _
        Application.Index(rawData, Evaluate("ROW(1:" & rawCount & ")"), _
                          Evaluate("COLUMN(A:J)"))

    ' 列幅調整
    wsRaw.Columns("A:A").ColumnWidth = 8
    wsRaw.Columns("B:B").ColumnWidth = 10
    wsRaw.Columns("C:C").ColumnWidth = 26
    wsRaw.Columns("D:D").ColumnWidth = 18
    wsRaw.Columns("E:E").ColumnWidth = 20
    wsRaw.Columns("F:F").ColumnWidth = 10
    wsRaw.Columns("G:G").ColumnWidth = 36
    wsRaw.Columns("H:H").ColumnWidth = 46
    wsRaw.Columns("I:I").ColumnWidth = 15
    wsRaw.Columns("J:J").ColumnWidth = 8

    ' ウィンドウ枠の固定・オートフィルタ
    wsRaw.Range("A2").Select
    ActiveWindow.FreezePanes = True
    wsRaw.Range("A1:J1").AutoFilter

    ' 保存
    Application.DisplayAlerts = False
    wbNew.SaveAs savePath, xlOpenXMLWorkbook
    wbNew.Close False
    Application.DisplayAlerts = True

    MsgBox "RAWデータ生成が完了しました！" & vbCrLf & vbCrLf & _
           "レコード数：" & rawCount & " 件" & vbCrLf & _
           "ファイル数：" & invCount & " ヶ月分" & vbCrLf & vbCrLf & _
           "保存先：" & savePath, vbInformation, "完了"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    MsgBox "エラーが発生しました：" & vbCrLf & Err.Description, vbCritical
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub


'==============================================================================
' ② PPT自動生成（メイン）
'==============================================================================
Sub Main_PPT自動生成()

    Dim pptPath  As String
    Dim corpCode As String
    Dim shopName As String
    Dim wsRaw    As Worksheet

    '----------------------------------------------------------------------
    ' STEP1: PPTファイルを選択
    '----------------------------------------------------------------------
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "対象のPPTファイルを選択してください"
    fd.Filters.Clear
    fd.Filters.Add "PowerPoint", "*.pptx;*.pptm"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        pptPath = fd.SelectedItems(1)
    Else
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If

    '----------------------------------------------------------------------
    ' STEP2: RAWデータシートの確認
    '----------------------------------------------------------------------
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets("RAWデータ")
    On Error GoTo ErrHandler2

    If wsRaw Is Nothing Then
        MsgBox "「RAWデータ」シートが見つかりません。" & vbCrLf & vbCrLf & _
               "【対処方法】" & vbCrLf & _
               "① 先に「Main_RAWデータ生成」を実行して出力されたExcelを開く" & vbCrLf & _
               "② そのExcelの「RAWデータ」シートをこのブックにコピーする" & vbCrLf & _
               "③ 再度「Main_PPT自動生成」を実行する", vbExclamation
        Exit Sub
    End If

    '----------------------------------------------------------------------
    ' STEP3: 法人選択
    '----------------------------------------------------------------------
    corpCode = SelectCorp_M4(wsRaw)
    If corpCode = "" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If

    '----------------------------------------------------------------------
    ' STEP4: 店舗選択
    '----------------------------------------------------------------------
    shopName = SelectShop_M4(wsRaw, corpCode)
    If shopName = "" Then
        MsgBox "キャンセルされました。", vbInformation
        Exit Sub
    End If

    '----------------------------------------------------------------------
    ' STEP5: データ集計
    '----------------------------------------------------------------------
    Dim scc As Object  ' 法人内 店舗×コード件数（P4・P6用）
    Dim sd  As Object  ' 選択店舗の詳細（P8・P10用）
    Call CollectData_M4(wsRaw, corpCode, shopName, scc, sd)

    '----------------------------------------------------------------------
    ' STEP6: PPTを開いて各スライドを更新
    '----------------------------------------------------------------------
    Dim pptApp  As Object
    Dim pptPres As Object

    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    On Error GoTo ErrHandler2

    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Open(pptPath)

    Application.StatusBar = "P4更新中（エリア別法人比較）..."
    Call UpdateP4_M4(pptPres, scc)

    Application.StatusBar = "P6更新中（店舗別マトリクス）..."
    Call UpdateP6_M4(pptPres, scc)

    Application.StatusBar = "P8更新中（店舗個別サマリー）..."
    Call UpdateP8_M4(pptPres, sd, shopName)

    Application.StatusBar = "P10更新中（発生コード確認）..."
    Call UpdateP10_M4(pptPres, sd, shopName)

    '----------------------------------------------------------------------
    ' STEP7: 別名保存（元ファイル名_店舗名_日付.pptx）
    '----------------------------------------------------------------------
    Dim savePath As String
    savePath = Left(pptPath, InStrRev(pptPath, ".") - 1) & _
               "_" & shopName & "_" & Format(Now, "yyyymmdd") & ".pptx"
    pptPres.SaveAs savePath

    Application.StatusBar = False
    MsgBox "PPT自動生成が完了しました！" & vbCrLf & vbCrLf & _
           "対象店舗：" & shopName & vbCrLf & _
           "保存先：" & savePath, vbInformation, "完了"

    Set pptPres = Nothing
    Set pptApp = Nothing
    Exit Sub

ErrHandler2:
    MsgBox "エラーが発生しました：" & vbCrLf & Err.Description, vbCritical
    Application.StatusBar = False
    If Not pptPres Is Nothing Then pptPres.Close False
End Sub


'==============================================================================
' ③ 内部処理（Module4専用プライベート関数）
'==============================================================================

' ---- 文字列から数字のみ抽出 ----
Private Function ExtractNumbers_M4(s As String) As String
    Dim i   As Integer
    Dim buf As String
    buf = ""
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then buf = buf & ch
    Next i
    ExtractNumbers_M4 = buf
End Function

' ---- ファイル名からYYYYMMを抽出 ----
Private Function ExtractYYYYMM_M4(fileName As String) As String
    Dim i   As Long
    Dim buf As String
    buf = ""
    For i = 1 To Len(fileName)
        Dim ch As String
        ch = Mid(fileName, i, 1)
        If ch >= "0" And ch <= "9" Then
            buf = buf & ch
            If Len(buf) = 6 Then
                If Left(buf, 2) = "20" Then
                    ExtractYYYYMM_M4 = buf
                    Exit Function
                Else
                    buf = Mid(buf, 2)
                End If
            End If
        Else
            buf = ""
        End If
    Next i
    ExtractYYYYMM_M4 = ""
End Function

' ---- 法人選択ダイアログ ----
Private Function SelectCorp_M4(wsRaw As Worksheet) As String
    Dim d       As Object
    Dim lastRow As Long
    Dim i       As Long

    Set d = CreateObject("Scripting.Dictionary")
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 2).End(xlUp).Row

    For i = 2 To lastRow
        If wsRaw.Cells(i, 2).Value <> "" Then
            Dim no As String
            no = Left(CStr(CLng(wsRaw.Cells(i, 2).Value)), 3)
            If Not d.Exists(no) Then d(no) = wsRaw.Cells(i, 3).Value
        End If
    Next i

    ' ソート
    Dim keys As Variant
    keys = d.Keys
    Dim j As Integer, k As Integer, tmp As String
    For j = 0 To UBound(keys) - 1
        For k = j + 1 To UBound(keys)
            If keys(j) > keys(k) Then
                tmp = keys(j): keys(j) = keys(k): keys(k) = tmp
            End If
        Next k
    Next j

    Dim lst As String
    For j = 0 To UBound(keys)
        lst = lst & (j + 1) & ". 法人コード:" & keys(j) & _
              "  (" & d(keys(j)) & " 他)" & vbCrLf
    Next j

    Dim sel As String
    sel = InputBox("法人を番号で選択してください：" & vbCrLf & vbCrLf & lst, "法人選択")
    If sel = "" Or Not IsNumeric(sel) Then Exit Function

    Dim idx As Integer
    idx = CInt(sel) - 1
    If idx < 0 Or idx > UBound(keys) Then
        MsgBox "無効な番号です。", vbExclamation
        Exit Function
    End If
    SelectCorp_M4 = keys(idx)
End Function

' ---- 店舗選択ダイアログ ----
Private Function SelectShop_M4(wsRaw As Worksheet, corpCode As String) As String
    Dim d       As Object
    Dim lastRow As Long
    Dim i       As Long

    Set d = CreateObject("Scripting.Dictionary")
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 2).End(xlUp).Row

    For i = 2 To lastRow
        If wsRaw.Cells(i, 2).Value <> "" Then
            If Left(CStr(CLng(wsRaw.Cells(i, 2).Value)), 3) = corpCode Then
                Dim nm As String
                nm = wsRaw.Cells(i, 3).Value
                If Not d.Exists(nm) Then d(nm) = 1
            End If
        End If
    Next i

    ' ソート
    Dim shops() As String
    ReDim shops(d.Count - 1)
    Dim j As Integer
    For j = 0 To d.Count - 1: shops(j) = d.Keys()(j): Next j
    Dim k As Integer, tmp As String
    For j = 0 To UBound(shops) - 1
        For k = j + 1 To UBound(shops)
            If shops(j) > shops(k) Then
                tmp = shops(j): shops(j) = shops(k): shops(k) = tmp
            End If
        Next k
    Next j

    Dim lst As String
    For j = 0 To UBound(shops)
        lst = lst & (j + 1) & ". " & shops(j) & vbCrLf
    Next j

    Dim sel As String
    sel = InputBox("店舗を番号で選択してください：" & vbCrLf & vbCrLf & lst, "店舗選択")
    If sel = "" Or Not IsNumeric(sel) Then Exit Function

    Dim idx As Integer
    idx = CInt(sel) - 1
    If idx < 0 Or idx > UBound(shops) Then
        MsgBox "無効な番号です。", vbExclamation
        Exit Function
    End If
    SelectShop_M4 = shops(idx)
End Function

' ---- データ集計（RAWデータシートから） ----
Private Sub CollectData_M4(wsRaw As Worksheet, corpCode As String, shopName As String, _
                            ByRef scc As Object, ByRef sd As Object)
    Set scc = CreateObject("Scripting.Dictionary")
    Set sd  = CreateObject("Scripting.Dictionary")

    Dim major(3) As Long
    major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003

    Dim lastRow As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    Dim i  As Long
    Dim ci As Integer
    For i = 2 To lastRow
        If wsRaw.Cells(i, 2).Value = "" Then GoTo Skip_M4
        If Not IsNumeric(wsRaw.Cells(i, 6).Value) Then GoTo Skip_M4

        Dim no As String: no = Left(CStr(CLng(wsRaw.Cells(i, 2).Value)), 3)
        Dim nm As String: nm = wsRaw.Cells(i, 3).Value
        Dim cd As Long:   cd = CLng(wsRaw.Cells(i, 6).Value)
        Dim fl As String: fl = wsRaw.Cells(i, 10).Value
        Dim mo As String: mo = CStr(wsRaw.Cells(i, 1).Value)

        ' 法人内の重大エラー集計（P4・P6用）
        If no = corpCode And fl = "重大" Then
            For ci = 0 To 3
                If cd = major(ci) Then
                    Dim k As String: k = nm & "_" & CStr(cd)
                    If scc.Exists(k) Then scc(k) = scc(k) + 1 Else scc(k) = 1
                End If
            Next ci
        End If

        ' 選択店舗の詳細（P8・P10用）
        If nm = shopName And fl = "重大" Then
            For ci = 0 To 3
                If cd = major(ci) Then
                    Dim ck As String: ck = "cnt_" & CStr(cd)
                    Dim mk As String: mk = "mon_" & CStr(cd)
                    If sd.Exists(ck) Then sd(ck) = sd(ck) + 1 Else sd(ck) = 1
                    If sd.Exists(mk) Then
                        If InStr(sd(mk), mo) = 0 Then sd(mk) = sd(mk) & "," & mo
                    Else
                        sd(mk) = mo
                    End If
                End If
            Next ci
        End If

Skip_M4:
    Next i
End Sub

' ---- P4更新：エリア別法人比較テーブル（スライド4） ----
Private Sub UpdateP4_M4(pptPres As Object, scc As Object)
    Dim sl As Object: Set sl = pptPres.Slides(4)
    ' 列順: 法人名(1), 1910(2), 1223(3), 1016(4), 1003(5), 合計(6), 優先度(7)
    Dim colCodes(3) As Long
    colCodes(0) = 1910: colCodes(1) = 1223: colCodes(2) = 1016: colCodes(3) = 1003

    Dim sh As Object
    For Each sh In sl.Shapes
        If sh.HasTable Then
            Dim tbl As Object: Set tbl = sh.Table
            Dim r As Integer
            For r = 2 To tbl.Rows.Count
                Dim nm As String
                nm = Trim(tbl.Cell(r, 1).Shape.TextFrame.TextRange.Text)
                If nm = "" Then GoTo NextR_P4

                Dim tot As Integer: tot = 0
                Dim ci  As Integer
                For ci = 0 To 3
                    Dim k   As String: k = nm & "_" & CStr(colCodes(ci))
                    Dim cnt As Integer: cnt = 0
                    If scc.Exists(k) Then cnt = scc(k)
                    Dim ct As String
                    If cnt = 0 Then ct = "－" ElseIf cnt >= 3 Then ct = "●多" Else ct = "●"
                    tbl.Cell(r, ci + 2).Shape.TextFrame.TextRange.Text = ct
                    tot = tot + cnt
                Next ci
                tbl.Cell(r, 6).Shape.TextFrame.TextRange.Text = CStr(tot) & "件"
                Dim pr As String
                If tot >= 3 Then pr = "★高" ElseIf tot >= 1 Then pr = "中" Else pr = "低"
                tbl.Cell(r, 7).Shape.TextFrame.TextRange.Text = pr
NextR_P4:
            Next r
        End If
    Next sh
End Sub

' ---- P6更新：店舗別マトリクス（スライド6） ----
Private Sub UpdateP6_M4(pptPres As Object, scc As Object)
    Dim sl As Object: Set sl = pptPres.Slides(6)
    ' 列順: 店舗名(1), 1910(2), 1223(3), 1016(4), 1003(5), 合計(6), 優先度(7)
    Dim colCodes(3) As Long
    colCodes(0) = 1910: colCodes(1) = 1223: colCodes(2) = 1016: colCodes(3) = 1003

    Dim sh As Object
    For Each sh In sl.Shapes
        If sh.HasTable Then
            Dim tbl As Object: Set tbl = sh.Table
            Dim r As Integer
            For r = 2 To tbl.Rows.Count
                Dim nm As String
                nm = Trim(tbl.Cell(r, 1).Shape.TextFrame.TextRange.Text)
                If nm = "" Then GoTo NextR_P6

                Dim tot As Integer: tot = 0
                Dim ci  As Integer
                For ci = 0 To 3
                    Dim k   As String: k = nm & "_" & CStr(colCodes(ci))
                    Dim cnt As Integer: cnt = 0
                    If scc.Exists(k) Then cnt = scc(k)
                    Dim ct As String: If cnt = 0 Then ct = "－" Else ct = "●"
                    tbl.Cell(r, ci + 2).Shape.TextFrame.TextRange.Text = ct
                    tot = tot + cnt
                Next ci
                tbl.Cell(r, 6).Shape.TextFrame.TextRange.Text = CStr(tot) & "件"
                Dim pr As String
                If tot >= 2 Then pr = "★高" ElseIf tot >= 1 Then pr = "中" Else pr = "低"
                tbl.Cell(r, 7).Shape.TextFrame.TextRange.Text = pr
NextR_P6:
            Next r
        End If
    Next sh
End Sub

' ---- P8更新：店舗個別サマリー（スライド8） ----
Private Sub UpdateP8_M4(pptPres As Object, sd As Object, shopName As String)
    Dim sl As Object: Set sl = pptPres.Slides(8)
    Dim sh As Object
    For Each sh In sl.Shapes
        ' 店舗名プレースホルダーを置換
        If sh.HasTextFrame Then
            Dim txt As String: txt = sh.TextFrame.TextRange.Text
            If InStr(txt, "ExampleBrand B ○○○○") > 0 Then
                sh.TextFrame.TextRange.Text = Replace(txt, "ExampleBrand B ○○○○", shopName)
            End If
        End If
        ' テーブル内のセルを更新
        If sh.HasTable Then
            Dim tbl As Object: Set tbl = sh.Table
            Dim ri  As Integer
            For ri = 1 To tbl.Rows.Count
                Dim lbl As String
                lbl = Trim(tbl.Cell(ri, 1).Shape.TextFrame.TextRange.Text)
                Select Case lbl
                    Case "重大エラーコード"
                        Dim codes As String: codes = ""
                        Dim major(3) As Long
                        major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003
                        Dim cj As Integer
                        For cj = 0 To 3
                            If sd.Exists("cnt_" & CStr(major(cj))) Then
                                If sd("cnt_" & CStr(major(cj))) > 0 Then
                                    If codes <> "" Then codes = codes & "　"
                                    codes = codes & CStr(major(cj))
                                End If
                            End If
                        Next cj
                        If codes = "" Then codes = "なし"
                        tbl.Cell(ri, 2).Shape.TextFrame.TextRange.Text = codes
                    Case "1910 発生件数"
                        Dim c1 As Integer: c1 = 0
                        If sd.Exists("cnt_1910") Then c1 = sd("cnt_1910")
                        tbl.Cell(ri, 2).Shape.TextFrame.TextRange.Text = CStr(c1) & "件"
                    Case "発生月"
                        tbl.Cell(ri, 2).Shape.TextFrame.TextRange.Text = GetMonths_M4(sd)
                    Case "主な担当区分"
                        ' 将来拡張用（現状は空欄のまま）
                End Select
            Next ri
        End If
    Next sh
End Sub

' ---- P10更新：発生コード確認（スライド10） ----
Private Sub UpdateP10_M4(pptPres As Object, sd As Object, shopName As String)
    Dim sl    As Object: Set sl = pptPres.Slides(10)
    Dim major(3) As Long
    major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003

    ' 全コードの合計件数
    Dim tot As Integer: tot = 0
    Dim cj  As Integer
    For cj = 0 To 3
        If sd.Exists("cnt_" & CStr(major(cj))) Then
            tot = tot + sd("cnt_" & CStr(major(cj)))
        End If
    Next cj

    Dim sh As Object
    For Each sh In sl.Shapes
        If sh.HasTextFrame Then
            Dim txt As String: txt = sh.TextFrame.TextRange.Text
            ' 店舗名置換
            If InStr(txt, "ExampleBrand B ○○○○") > 0 Then
                sh.TextFrame.TextRange.Text = Replace(txt, "ExampleBrand B ○○○○", shopName)
            End If
            ' 件数バー更新
            If InStr(txt, "重大エラー発生件数：　　　件") > 0 Then
                sh.TextFrame.TextRange.Text = Replace(txt, "重大エラー発生件数：　　　件", _
                                                      "重大エラー発生件数：" & CStr(tot) & "件")
            End If
        End If
    Next sh
End Sub

' ---- 発生月を「8月・9月・…」形式に変換 ----
Private Function GetMonths_M4(sd As Object) As String
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim major(3) As Long
    major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003
    Dim cj As Integer
    For cj = 0 To 3
        Dim mk As String: mk = "mon_" & CStr(major(cj))
        If sd.Exists(mk) Then
            Dim ms() As String: ms = Split(sd(mk), ",")
            Dim m As Variant
            For Each m In ms
                If Not d.Exists(m) Then d(m) = 1
            Next m
        End If
    Next cj

    ' ソート
    Dim keys As Variant: keys = d.Keys
    Dim i As Integer, j As Integer, tmp As String
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(i) > keys(j) Then tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
        Next j
    Next i

    Dim result As String
    For i = 0 To UBound(keys)
        Dim s As String: s = CStr(keys(i))
        If Len(s) = 6 Then
            If result <> "" Then result = result & "・"
            result = result & CStr(CInt(Right(s, 2))) & "月"
        End If
    Next i
    If result = "" Then result = "－"
    GetMonths_M4 = result
End Function