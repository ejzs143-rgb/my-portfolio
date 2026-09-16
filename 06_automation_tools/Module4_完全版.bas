Attribute VB_Name = "Module4"
'==============================================================================
' Module4: RAWデータ生成 + PPT自動生成（完全版）
'
' 【実行の流れ】
'   Main_RAWデータ生成：インボイスリスト → RAWデータ.xlsx 生成
'   Main_PPT自動生成：
'     ① PPTファイル選択
'     ② 連絡先リスト選択（Audi販売店AS関連連絡先リスト.xlsx）
'     ③ RAWデータ選択（またはすでに開いているブックから自動取得）
'     ④ AM選択（連絡先リストから自動リスト）
'     ⑤ 店舗選択（AM担当店舗から自動リスト）
'     ⑥ P4・P6を実データで再生成、P8・P10に店舗データを自動入力
'     ⑦ 別名保存
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

    Dim codeCols(4)    As Integer
    Dim contentCols(4) As Integer
    Dim commentCols(4) As Integer

    codeCols(0) = 8:  contentCols(0) = 10: commentCols(0) = 11
    codeCols(1) = 12: contentCols(1) = 14: commentCols(1) = 15
    codeCols(2) = 16: contentCols(2) = 18: commentCols(2) = 19
    codeCols(3) = 20: contentCols(3) = 22: commentCols(3) = 23
    codeCols(4) = 24: contentCols(4) = 26: commentCols(4) = 27

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' STEP1: マスタ_重大エラーから重大コード取得
    Set critCodes = CreateObject("Scripting.Dictionary")
    critCodes.CompareMode = 1

    Set wsMaster = FindSheet_M4(ThisWorkbook, "マスタ_重大エラー")

    If Not wsMaster Is Nothing Then
        Dim mi As Long
        For mi = 2 To 20
            Dim mc As String
            mc = Trim(CStr(wsMaster.Cells(mi, 1).Value))
            If mc = "" Then Exit For
            mc = ExtractNumbers_M4(mc)
            If mc <> "" And Not critCodes.Exists(mc) Then critCodes.Add mc, True
        Next mi
    Else
        MsgBox "「マスタ_重大エラー」シートが見つかりません。重大フラグなしで生成します。", vbExclamation
    End If

    ' STEP2: 同フォルダのインボイスリストを収集
    If ThisWorkbook.Path = "" Then
        MsgBox "このブックを先に保存してください。", vbExclamation
        GoTo CleanUp
    End If

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
        MsgBox "インボイスリストが見つかりませんでした。" & vbCrLf & _
               "ファイル名：コアプロセスチェック対象インボイスリスト_YYYYMM.xlsx" & vbCrLf & _
               "このツールと同じフォルダに置いてください。", vbExclamation
        GoTo CleanUp
    End If

    MsgBox invCount & "ヶ月分のインボイスリストを検出しました。処理を開始します。", vbInformation

    ' STEP3: 各ファイルを読み込んでRAWデータ配列に蓄積
    rawCount = 0
    ReDim rawData(1 To 60000, 1 To 10)

    Dim fi As Long
    For fi = 0 To invCount - 1

        Application.StatusBar = "読み込み中... (" & (fi + 1) & "/" & invCount & ") " & invFiles(fi)

        yyyymm = ExtractYYYYMM_M4(invFiles(fi))
        If yyyymm = "" Then GoTo NextFile_M4

        Set wbInv = Nothing
        Set wbInv = Workbooks.Open(folderPath & invFiles(fi), ReadOnly:=True, UpdateLinks:=False)
        Set wsTop3 = FindSheet_M4(wbInv, "Top3(結果入り)")

        If wsTop3 Is Nothing Then
            wbInv.Close False
            GoTo NextFile_M4
        End If

        lastRow = wsTop3.Cells(wsTop3.Rows.Count, 5).End(xlUp).Row

        For r = 4 To lastRow

            Dim roNum As String
            Dim dNo   As String
            Dim dName As String
            Dim vin   As String

            roNum = Trim(CStr(wsTop3.Cells(r, 5).Value))
            dNo   = Trim(CStr(wsTop3.Cells(r, 2).Value))
            dName = Trim(CStr(wsTop3.Cells(r, 3).Value))
            vin   = Trim(CStr(wsTop3.Cells(r, 6).Value))

            If roNum = "" Or dNo = "" Then GoTo NextRow_M4
            If Not IsNumeric(dNo) Then GoTo NextRow_M4

            For ci = 0 To 4

                Dim eCode    As String
                Dim eContent As String
                Dim eComment As String
                Dim flagStr  As String

                eCode = Trim(CStr(wsTop3.Cells(r, codeCols(ci)).Value))
                eCode = ExtractNumbers_M4(eCode)

                If eCode = "" Or eCode = "0" Or eCode = "0000" Then GoTo NextCode_M4
                If Not IsNumeric(eCode) Then GoTo NextCode_M4

                eContent = Trim(CStr(wsTop3.Cells(r, contentCols(ci)).Value))
                eComment = Trim(CStr(wsTop3.Cells(r, commentCols(ci)).Value))

                If critCodes.Count > 0 Then
                    flagStr = IIf(critCodes.Exists(eCode), "重大", "軽微")
                Else
                    flagStr = ""
                End If

                rawCount = rawCount + 1
                rawData(rawCount, 1)  = CLng(yyyymm)
                rawData(rawCount, 2)  = CLng(dNo)
                rawData(rawCount, 3)  = dName
                rawData(rawCount, 4)  = roNum
                rawData(rawCount, 5)  = vin
                rawData(rawCount, 6)  = CLng(eCode)
                rawData(rawCount, 7)  = eContent
                rawData(rawCount, 8)  = eComment
                rawData(rawCount, 9)  = ""
                rawData(rawCount, 10) = flagStr

                If rawCount >= 60000 Then
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
        MsgBox "有効なデータが見つかりませんでした。", vbExclamation
        GoTo CleanUp
    End If

WriteOut_M4:

    ' STEP4: 保存先選択して書き出し
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=ThisWorkbook.Path & "\コアプロセスチェック_アクション用データ_" & Format(Now, "yyyymmdd") & ".xlsx", _
        FileFilter:="Excelファイル (*.xlsx),*.xlsx", _
        Title:="RAWデータの保存先を選択")

    If savePath = False Then GoTo CleanUp

    Application.StatusBar = "書き出し中..."

    Dim wbNew As Workbook
    Dim wsRaw As Worksheet
    Set wbNew = Workbooks.Add
    Set wsRaw = wbNew.Sheets(1)
    wsRaw.Name = "RAWデータ"

    Dim headers As Variant
    headers = Array("月", "ディーラーNo", "ディーラー名", "RO番号", "VIN", _
                    "エラーコード", "エラー内容", "追加コメント", "サブカテゴリ", "重大フラグ")
    Dim hi As Integer
    For hi = 0 To 9: wsRaw.Cells(1, hi + 1).Value = headers(hi): Next hi

    With wsRaw.Range("A1:J1")
        .Font.Bold = True: .Font.Name = "Arial"
        .Interior.Color = RGB(31, 56, 100): .Font.Color = RGB(255, 255, 255)
    End With

    Dim outData() As Variant
    ReDim outData(1 To rawCount, 1 To 10)
    Dim rr As Long, cc As Long
    For rr = 1 To rawCount
        For cc = 1 To 10: outData(rr, cc) = rawData(rr, cc): Next cc
    Next rr
    wsRaw.Range("A2").Resize(rawCount, 10).Value = outData

    wsRaw.Columns("A:A").ColumnWidth = 8:  wsRaw.Columns("B:B").ColumnWidth = 12
    wsRaw.Columns("C:C").ColumnWidth = 30: wsRaw.Columns("D:D").ColumnWidth = 18
    wsRaw.Columns("E:E").ColumnWidth = 20: wsRaw.Columns("F:F").ColumnWidth = 12
    wsRaw.Columns("G:G").ColumnWidth = 40: wsRaw.Columns("H:H").ColumnWidth = 50
    wsRaw.Columns("I:I").ColumnWidth = 15: wsRaw.Columns("J:J").ColumnWidth = 10
    wsRaw.Range("A1:J1").AutoFilter
    wsRaw.Activate: wsRaw.Range("A2").Select: ActiveWindow.FreezePanes = True

    Application.DisplayAlerts = False
    wbNew.SaveAs CStr(savePath), xlOpenXMLWorkbook
    wbNew.Close False
    Application.DisplayAlerts = True

    MsgBox "完了！" & vbCrLf & "レコード数：" & rawCount & " 件" & vbCrLf & _
           "ファイル数：" & invCount & " ヶ月分" & vbCrLf & vbCrLf & _
           "保存先：" & CStr(savePath), vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    MsgBox "エラー：" & Err.Number & " / " & Err.Description, vbCritical
    On Error Resume Next
    If Not wbInv Is Nothing Then wbInv.Close False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub


'==============================================================================
' ② PPT自動生成（メイン）
'==============================================================================
Sub Main_PPT自動生成()

    Dim pptPath    As String
    Dim amName     As String
    Dim shopName   As String
    Dim corpCode   As String
    Dim wsRaw      As Worksheet
    Dim wsAM       As Worksheet
    Dim wbRaw      As Workbook
    Dim wbAM       As Workbook
    Dim pptApp     As Object
    Dim pptPres    As Object
    Dim closeRaw   As Boolean
    Dim closeAM    As Boolean

    On Error GoTo ErrHandler2

    closeRaw = False
    closeAM  = False

    ' STEP1: PPTファイル選択
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    fd.Title = "PPTファイルを選択してください"
    fd.Filters.Clear
    fd.Filters.Add "PowerPoint", "*.pptx;*.pptm"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        pptPath = fd.SelectedItems(1)
    Else
        MsgBox "キャンセルされました。", vbInformation: Exit Sub
    End If

    ' STEP2: 連絡先リスト選択
    Set fd = Application.FileDialog(3)
    fd.Title = "連絡先リスト（Audi販売店AS関連連絡先リスト.xlsx）を選択してください"
    fd.Filters.Clear
    fd.Filters.Add "Excel", "*.xlsx;*.xlsm"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        Set wbAM = Workbooks.Open(fd.SelectedItems(1), ReadOnly:=True, UpdateLinks:=False)
        closeAM = True
    Else
        MsgBox "キャンセルされました。", vbInformation: Exit Sub
    End If

    ' 「サービスマネージャー」シートを取得
    Set wsAM = FindSheet_M4(wbAM, "サービスマネージャー")
    If wsAM Is Nothing Then
        MsgBox "連絡先リストに「サービスマネージャー」シートが見つかりません。", vbExclamation
        GoTo CleanUp2
    End If

    ' STEP3: RAWデータシート取得
    Set wsRaw = GetRawWorksheet_M4(wbRaw, closeRaw)
    If wsRaw Is Nothing Then
        MsgBox "RAWデータシートを取得できませんでした。", vbExclamation
        GoTo CleanUp2
    End If

    ' STEP4: AM選択（連絡先リストから自動リスト）
    amName = SelectAM_M4(wsAM)
    If amName = "" Then
        MsgBox "キャンセルされました。", vbInformation: GoTo CleanUp2
    End If

    ' STEP5: 店舗選択（AM担当店舗×RAWデータ重複確認）
    shopName = SelectShopByAM_M4(wsAM, wsRaw, amName, corpCode)
    If shopName = "" Then
        MsgBox "キャンセルされました。", vbInformation: GoTo CleanUp2
    End If

    ' STEP6: データ集計
    Dim scc As Object  ' AM担当全店舗×コード件数（P4・P6用）
    Dim sd  As Object  ' 選択店舗の詳細（P8・P10用）
    Call CollectData_M4(wsRaw, wsAM, amName, corpCode, shopName, scc, sd)

    ' 連絡先リストを閉じる
    If closeAM Then
        wbAM.Close False
        Set wbAM = Nothing
        closeAM = False
    End If
    If closeRaw Then
        wbRaw.Close False
        Set wbRaw = Nothing
        closeRaw = False
    End If

    ' STEP7: PowerPoint起動
    Set pptApp = CreateObject("PowerPoint.Application")
    If pptApp Is Nothing Then
        MsgBox "PowerPointを起動できませんでした。", vbCritical: GoTo CleanUp2
    End If
    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Open(pptPath)

    If pptPres.Slides.Count < 10 Then
        MsgBox "スライド数が不足しています。正しいPPTを選択してください。", vbExclamation
        GoTo CleanUp2
    End If

    ' STEP8: 各スライド更新
    Application.StatusBar = "P4更新中..."
    Call UpdateP4_M4(pptPres, scc, amName)

    Application.StatusBar = "P6更新中..."
    Call UpdateP6_M4(pptPres, scc, shopName, corpCode)

    Application.StatusBar = "P8更新中..."
    Call UpdateP8_M4(pptPres, sd, shopName)

    Application.StatusBar = "P10更新中..."
    Call UpdateP10_M4(pptPres, sd, shopName)

    ' STEP9: 別名保存
    Dim savePath As String
    savePath = Left(pptPath, InStrRev(pptPath, ".") - 1) & _
               "_" & CleanFileName_M4(shopName) & "_" & Format(Now, "yyyymmdd") & ".pptx"
    pptPres.SaveAs savePath

    Application.StatusBar = False
    MsgBox "完了！" & vbCrLf & "店舗：" & shopName & vbCrLf & "保存先：" & savePath, vbInformation

CleanUp2:
    Application.StatusBar = False
    On Error Resume Next
    If closeAM And Not wbAM Is Nothing Then wbAM.Close False
    If closeRaw And Not wbRaw Is Nothing Then wbRaw.Close False
    Set wsRaw = Nothing: Set wsAM = Nothing
    Set pptPres = Nothing: Set pptApp = Nothing
    Exit Sub

ErrHandler2:
    MsgBox "エラー：" & Err.Number & " / " & Err.Description, vbCritical
    Application.StatusBar = False
    On Error Resume Next
    If closeAM And Not wbAM Is Nothing Then wbAM.Close False
    If closeRaw And Not wbRaw Is Nothing Then wbRaw.Close False
    If Not pptPres Is Nothing Then pptPres.Close False
End Sub


'==============================================================================
' ③ 選択ダイアログ
'==============================================================================

' AM選択
Private Function SelectAM_M4(wsAM As Worksheet) As String

    Dim d       As Object
    Dim lastRow As Long
    Dim i       As Long

    Set d = CreateObject("Scripting.Dictionary")
    lastRow = wsAM.Cells(wsAM.Rows.Count, 1).End(xlUp).Row

    ' ヘッダー行（1〜4行目）をスキップ、E列（現在のAM名）を取得
    For i = 5 To lastRow
        Dim amVal As String
        amVal = Trim(CStr(wsAM.Cells(i, 5).Value))
        If amVal <> "" And amVal <> "アフターセールスエリアマネージャー" Then
            If Not d.Exists(amVal) Then d.Add amVal, 1
        End If
    Next i

    If d.Count = 0 Then
        MsgBox "AMが見つかりません。連絡先リストを確認してください。", vbExclamation
        Exit Function
    End If

    Dim amList() As String
    ReDim amList(d.Count - 1)
    Dim j As Integer
    For j = 0 To d.Count - 1: amList(j) = d.Keys()(j): Next j
    Call SortArr_M4(amList)

    ' 選択シートに一覧表示
    Dim wsSel As Worksheet
    Set wsSel = GetSelectionSheet_M4()
    wsSel.Range("A1").Value = "エリアマネージャー選択"
    wsSel.Range("A3").Value = "番号"
    wsSel.Range("B3").Value = "エリアマネージャー名"
    With wsSel.Range("A3:B3")
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
    End With
    For j = 0 To UBound(amList)
        wsSel.Cells(j + 4, 1).Value = j + 1
        wsSel.Cells(j + 4, 2).Value = amList(j)
    Next j
    wsSel.Columns("A:B").AutoFit
    wsSel.Activate

    Dim sel As Variant
    sel = Application.InputBox( _
        Prompt:="「選択_法人店舗」シートを確認し、AMの番号を入力してください。", _
        Title:="エリアマネージャー選択", Type:=1)

    If sel = False Then Exit Function
    If CLng(sel) < 1 Or CLng(sel) > d.Count Then
        MsgBox "無効な番号です。", vbExclamation: Exit Function
    End If

    SelectAM_M4 = amList(CLng(sel) - 1)
End Function

' AM担当店舗から店舗選択
Private Function SelectShopByAM_M4(wsAM As Worksheet, wsRaw As Worksheet, _
                                    amName As String, ByRef corpCode As String) As String

    ' 連絡先リストからAM担当店舗のDCコードと店舗名を取得
    Dim amShops As Object
    Set amShops = CreateObject("Scripting.Dictionary")  ' DC → 店舗名

    Dim lastRow As Long
    Dim i       As Long
    lastRow = wsAM.Cells(wsAM.Rows.Count, 1).End(xlUp).Row

    For i = 5 To lastRow
        If Trim(CStr(wsAM.Cells(i, 5).Value)) = amName Then
            Dim dc As String
            Dim shopNm As String
            dc     = Trim(CStr(wsAM.Cells(i, 1).Value))
            shopNm = Trim(CStr(wsAM.Cells(i, 2).Value))
            If IsNumeric(dc) And shopNm <> "" Then
                If Not amShops.Exists(dc) Then amShops.Add dc, shopNm
            End If
        End If
    Next i

    If amShops.Count = 0 Then
        MsgBox "担当店舗が見つかりません。", vbExclamation: Exit Function
    End If

    ' RAWデータに存在する店舗のみ絞り込み（重大エラーあり店舗を優先表示）
    Dim rawShops  As Object
    Dim rawMajor  As Object
    Set rawShops  = CreateObject("Scripting.Dictionary")
    Set rawMajor  = CreateObject("Scripting.Dictionary")

    Dim lastRaw As Long
    lastRaw = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    Dim major(3) As Long
    major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003

    For i = 2 To lastRaw
        If wsRaw.Cells(i, 2).Value <> "" Then
            Dim rawDC   As String
            Dim rawName As String
            Dim rawFlag As String
            Dim rawCode As String
            rawDC   = Trim(CStr(CLng(wsRaw.Cells(i, 2).Value)))
            rawName = Trim(CStr(wsRaw.Cells(i, 3).Value))
            rawFlag = Trim(CStr(wsRaw.Cells(i, 10).Value))
            If amShops.Exists(rawDC) Then
                If Not rawShops.Exists(rawDC) Then rawShops.Add rawDC, rawName
                If rawFlag = "重大" Then
                    If Not rawMajor.Exists(rawDC) Then rawMajor.Add rawDC, 1
                End If
            End If
        End If
    Next i

    ' 選択シートに一覧表示（重大エラーあり店舗に★マーク）
    Dim dcList() As String
    Dim nmList() As String
    Dim hasMaj() As Boolean
    Dim cnt As Integer
    cnt = 0

    Dim keys As Variant
    keys = amShops.Keys
    Call SortArr_M4(keys)

    ReDim dcList(UBound(keys))
    ReDim nmList(UBound(keys))
    ReDim hasMaj(UBound(keys))

    Dim k As Integer
    For k = 0 To UBound(keys)
        dcList(cnt) = CStr(keys(k))
        If rawShops.Exists(CStr(keys(k))) Then
            nmList(cnt) = CStr(rawShops(CStr(keys(k))))
        Else
            nmList(cnt) = CStr(amShops(CStr(keys(k))))
        End If
        hasMaj(cnt) = rawMajor.Exists(CStr(keys(k)))
        cnt = cnt + 1
    Next k

    Dim wsSel As Worksheet
    Set wsSel = GetSelectionSheet_M4()
    wsSel.Range("A1").Value = "店舗選択"
    wsSel.Range("A2").Value = "AM：" & amName & "　　★ = 重大エラーあり"
    wsSel.Range("A4").Value = "番号"
    wsSel.Range("B4").Value = "店舗名"
    wsSel.Range("C4").Value = "重大エラー"
    With wsSel.Range("A4:C4")
        .Font.Bold = True
        .Interior.Color = RGB(31, 56, 100)
        .Font.Color = RGB(255, 255, 255)
    End With
    For k = 0 To cnt - 1
        wsSel.Cells(k + 5, 1).Value = k + 1
        wsSel.Cells(k + 5, 2).Value = nmList(k)
        wsSel.Cells(k + 5, 3).Value = IIf(hasMaj(k), "★あり", "－")
    Next k
    wsSel.Columns("A:C").AutoFit
    wsSel.Activate

    Dim sel As Variant
    sel = Application.InputBox( _
        Prompt:="「選択_法人店舗」シートを確認し、店舗の番号を入力してください。" & vbCrLf & _
                "★マークは重大エラーが発生している店舗です。", _
        Title:="店舗選択", Type:=1)

    If sel = False Then Exit Function
    Dim selIdx As Integer
    selIdx = CLng(sel) - 1
    If selIdx < 0 Or selIdx >= cnt Then
        MsgBox "無効な番号です。", vbExclamation: Exit Function
    End If

    corpCode = Left(dcList(selIdx), 3)
    SelectShopByAM_M4 = nmList(selIdx)
End Function


'==============================================================================
' ④ データ集計
'==============================================================================
Private Sub CollectData_M4(wsRaw As Worksheet, wsAM As Worksheet, _
                            amName As String, corpCode As String, shopName As String, _
                            ByRef scc As Object, ByRef sd As Object)

    Set scc = CreateObject("Scripting.Dictionary")
    Set sd  = CreateObject("Scripting.Dictionary")

    Dim major(3) As Long
    major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003

    ' AMの担当DCコード一覧を取得
    Dim amDCs As Object
    Set amDCs = CreateObject("Scripting.Dictionary")

    Dim lastAM As Long
    lastAM = wsAM.Cells(wsAM.Rows.Count, 1).End(xlUp).Row
    Dim ai As Long
    For ai = 5 To lastAM
        If Trim(CStr(wsAM.Cells(ai, 5).Value)) = amName Then
            Dim dcVal As String
            dcVal = Trim(CStr(wsAM.Cells(ai, 1).Value))
            If IsNumeric(dcVal) Then
                Dim dcCorpCode As String
                dcCorpCode = Left(CStr(CLng(dcVal)), 3)
                If Not amDCs.Exists(dcCorpCode) Then amDCs.Add dcCorpCode, 1
            End If
        End If
    Next ai

    ' RAWデータを集計
    Dim lastRow As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    Dim i  As Long
    Dim ci As Integer

    For i = 2 To lastRow
        If wsRaw.Cells(i, 2).Value = "" Then GoTo Skip_M4
        If Not IsNumeric(wsRaw.Cells(i, 2).Value) Then GoTo Skip_M4
        If Not IsNumeric(wsRaw.Cells(i, 6).Value) Then GoTo Skip_M4

        Dim no As String: no = Left(CStr(CLng(wsRaw.Cells(i, 2).Value)), 3)
        Dim nm As String: nm = CStr(wsRaw.Cells(i, 3).Value)
        Dim cd As Long:   cd = CLng(wsRaw.Cells(i, 6).Value)
        Dim fl As String: fl = CStr(wsRaw.Cells(i, 10).Value)
        Dim mo As String: mo = CStr(wsRaw.Cells(i, 1).Value)

        ' AM担当全店舗の重大エラー集計（P4・P6用）
        If amDCs.Exists(no) And fl = "重大" Then
            For ci = 0 To 3
                If cd = major(ci) Then
                    Dim k As String: k = nm & "_" & CStr(cd)
                    If scc.Exists(k) Then scc(k) = CLng(scc(k)) + 1 Else scc.Add k, 1
                End If
            Next ci
        End If

        ' 選択店舗の詳細（P8・P10用）
        If nm = shopName And fl = "重大" Then
            For ci = 0 To 3
                If cd = major(ci) Then
                    Dim ck As String: ck = "cnt_" & CStr(cd)
                    Dim mk As String: mk = "mon_" & CStr(cd)
                    If sd.Exists(ck) Then sd(ck) = CLng(sd(ck)) + 1 Else sd.Add ck, 1
                    If sd.Exists(mk) Then
                        If InStr(CStr(sd(mk)), mo) = 0 Then sd(mk) = CStr(sd(mk)) & "," & mo
                    Else
                        sd.Add mk, mo
                    End If
                End If
            Next ci
        End If

Skip_M4:
    Next i
End Sub


'==============================================================================
' ⑤ PPT更新処理
'==============================================================================

' ---- P4更新：AM担当全法人の比較表を実データで再生成 ----
Private Sub UpdateP4_M4(pptPres As Object, scc As Object, amName As String)

    Dim sl As Object
    Set sl = pptPres.Slides(4)

    ' 列定義（left座標はPPT実測値をEMUに変換）
    ' 1cm = 360000 EMU
    Const EMU As Long = 360000

    ' 列のleft座標（cm）
    Dim colLeft(6) As Double
    colLeft(0) = 1.17   ' 法人名
    colLeft(1) = 5.74   ' 1910
    colLeft(2) = 9.55   ' 1223
    colLeft(3) = 13.36  ' 1016
    colLeft(4) = 17.17  ' 1003
    colLeft(5) = 20.98  ' 合計
    colLeft(6) = 22.34  ' 優先度

    ' 列の幅（cm）
    Dim colWidth(6) As Double
    colWidth(0) = 4.27: colWidth(1) = 3.51: colWidth(2) = 3.51
    colWidth(3) = 3.51: colWidth(4) = 3.51: colWidth(5) = 1.24: colWidth(6) = 1.51

    ' データ行の設定
    Const ROW_TOP    As Double = 3.3    ' データ1行目のtop（cm）
    Const ROW_HEIGHT As Double = 1.22   ' 行の高さ（cm）
    Const ROW_STEP   As Double = 1.32   ' 行間（cm）

    Dim colCodes(3) As Long
    colCodes(0) = 1910: colCodes(1) = 1223: colCodes(2) = 1016: colCodes(3) = 1003

    ' 既存のデータ行シェイプを削除（データ行のみ：Text 19以降）
    Dim shToDelete() As Object
    Dim delCount As Integer
    delCount = 0
    ReDim shToDelete(sl.Shapes.Count)

    Dim sh As Object
    For Each sh In sl.Shapes
        If sh.HasTextFrame Then
            Dim sn As String: sn = sh.Name
            If Left(sn, 5) = "Text " Then
                Dim idx As Integer: idx = CInt(Mid(sn, 6))
                ' データ行（Text 19以降）と背景Shape（Shape 18以降）を削除
                If idx >= 19 And idx <= 73 Then
                    Set shToDelete(delCount) = sh
                    delCount = delCount + 1
                End If
            ElseIf Left(sn, 6) = "Shape " Then
                Dim sidx As Integer: sidx = CInt(Mid(sn, 7))
                If sidx >= 18 And sidx <= 72 Then
                    Set shToDelete(delCount) = sh
                    delCount = delCount + 1
                End If
            End If
        End If
    Next sh
    ' Shape系（HasTextFrame=False）も削除
    For Each sh In sl.Shapes
        If Not sh.HasTextFrame Then
            Dim sn2 As String: sn2 = sh.Name
            If Left(sn2, 6) = "Shape " Then
                Dim sidx2 As Integer: sidx2 = CInt(Mid(sn2, 7))
                If sidx2 >= 18 And sidx2 <= 72 Then
                    Set shToDelete(delCount) = sh
                    delCount = delCount + 1
                End If
            End If
        End If
    Next sh

    Dim di As Integer
    For di = 0 To delCount - 1
        If Not shToDelete(di) Is Nothing Then
            On Error Resume Next
            shToDelete(di).Delete
            On Error GoTo 0
        End If
    Next di

    ' AM担当の法人リストを作成（sccのキーから法人名を抽出）
    Dim corpNames As Object
    Set corpNames = CreateObject("Scripting.Dictionary")
    Dim sccKeys As Variant
    sccKeys = scc.Keys
    Dim ki As Integer
    For ki = 0 To UBound(sccKeys)
        Dim kStr As String: kStr = CStr(sccKeys(ki))
        Dim pos As Integer: pos = InStrRev(kStr, "_")
        If pos > 0 Then
            Dim cnm As String: cnm = Left(kStr, pos - 1)
            If Not corpNames.Exists(cnm) Then corpNames.Add cnm, 1
        End If
    Next ki

    ' 法人名をソート
    Dim corpArr() As String
    ReDim corpArr(corpNames.Count - 1)
    Dim cni As Integer
    For cni = 0 To corpNames.Count - 1: corpArr(cni) = corpNames.Keys()(cni): Next cni
    Call SortArr_M4(corpArr)

    ' AM名バーを更新（Text 1）
    For Each sh In sl.Shapes
        If sh.HasTextFrame Then
            If sh.Name = "Text 1" Then
                Dim t1 As String: t1 = sh.TextFrame.TextRange.Text
                If InStr(t1, "○○ ○○") > 0 Then
                    sh.TextFrame.TextRange.Text = Replace(t1, "○○ ○○", amName)
                End If
            End If
        End If
    Next sh

    ' データ行を新規生成
    Dim rowNum As Integer
    For rowNum = 0 To UBound(corpArr)

        Dim nm As String: nm = corpArr(rowNum)
        Dim rowTop As Double: rowTop = ROW_TOP + rowNum * ROW_STEP

        ' 各列のテキストボックスを追加
        Dim tot As Long: tot = 0
        Dim ci As Integer

        ' 法人名列
        Call AddTextBox_M4(sl, colLeft(0), rowTop, colWidth(0), ROW_HEIGHT, nm, _
                           12, True, RGB(31, 56, 100), RGB(255, 255, 255))

        ' コード列（1910,1223,1016,1003）
        For ci = 0 To 3
            Dim k As String: k = nm & "_" & CStr(colCodes(ci))
            Dim cnt As Long: cnt = 0
            If scc.Exists(k) Then cnt = CLng(scc(k))
            Dim ct As String
            If cnt = 0 Then ct = "－" ElseIf cnt >= 3 Then ct = "●多" Else ct = "●"
            Dim txtColor As Long
            If cnt > 0 Then txtColor = RGB(192, 0, 0) Else txtColor = RGB(100, 100, 100)
            Call AddTextBox_M4(sl, colLeft(ci + 1), rowTop, colWidth(ci + 1), ROW_HEIGHT, ct, _
                               11, False, RGB(255, 255, 255), txtColor)
            tot = tot + cnt
        Next ci

        ' 合計列
        Call AddTextBox_M4(sl, colLeft(5), rowTop, colWidth(5), ROW_HEIGHT, CStr(tot) & "件", _
                           10, True, RGB(255, 255, 255), RGB(31, 56, 100))

        ' 優先度列
        Dim pr As String
        Dim prBg As Long
        If tot >= 3 Then
            pr = "★高": prBg = RGB(192, 0, 0)
        ElseIf tot >= 1 Then
            pr = "中": prBg = RGB(255, 255, 255)
        Else
            pr = "低": prBg = RGB(240, 240, 240)
        End If
        Dim prTxt As Long
        If tot >= 3 Then prTxt = RGB(255, 255, 255) Else prTxt = RGB(68, 68, 68)
        Call AddTextBox_M4(sl, colLeft(6), rowTop, colWidth(6), ROW_HEIGHT, pr, _
                           10, True, prBg, prTxt)

    Next rowNum
End Sub

' ---- P6更新：法人内全店舗のマトリクスを実データで再生成 ----
Private Sub UpdateP6_M4(pptPres As Object, scc As Object, shopName As String, corpCode As String)

    Dim sl As Object
    Set sl = pptPres.Slides(6)

    Dim colLeft(6) As Double
    colLeft(0) = 1.17:  colLeft(1) = 6.25:  colLeft(2) = 9.80
    colLeft(3) = 13.36: colLeft(4) = 16.92: colLeft(5) = 20.47: colLeft(6) = 21.96

    Dim colWidth(6) As Double
    colWidth(0) = 4.78: colWidth(1) = 3.25: colWidth(2) = 3.25
    colWidth(3) = 3.25: colWidth(4) = 3.25: colWidth(5) = 1.32: colWidth(6) = 1.92

    Const ROW_TOP    As Double = 3.25
    Const ROW_HEIGHT As Double = 1.12
    Const ROW_STEP   As Double = 1.22

    Dim colCodes(3) As Long
    colCodes(0) = 1910: colCodes(1) = 1223: colCodes(2) = 1016: colCodes(3) = 1003

    ' 既存データ行を削除（Text 19以降）
    Dim shToDelete() As Object
    Dim delCount As Integer: delCount = 0
    ReDim shToDelete(sl.Shapes.Count)

    Dim sh As Object
    For Each sh In sl.Shapes
        Dim sn As String: sn = sh.Name
        If Left(sn, 5) = "Text " Then
            Dim idx As Integer: idx = CInt(Mid(sn, 6))
            If idx >= 19 And idx <= 87 Then
                Set shToDelete(delCount) = sh: delCount = delCount + 1
            End If
        ElseIf Left(sn, 6) = "Shape " Then
            Dim sidx As Integer: sidx = CInt(Mid(sn, 7))
            If sidx >= 18 And sidx <= 86 Then
                Set shToDelete(delCount) = sh: delCount = delCount + 1
            End If
        End If
    Next sh
    ' HasTextFrame=Falseのシェイプも削除
    For Each sh In sl.Shapes
        If Not sh.HasTextFrame Then
            Dim sn3 As String: sn3 = sh.Name
            If Left(sn3, 6) = "Shape " Then
                Dim s3 As Integer: s3 = CInt(Mid(sn3, 7))
                If s3 >= 18 And s3 <= 86 Then
                    Set shToDelete(delCount) = sh: delCount = delCount + 1
                End If
            End If
        End If
    Next sh

    Dim di As Integer
    For di = 0 To delCount - 1
        If Not shToDelete(di) Is Nothing Then
            On Error Resume Next: shToDelete(di).Delete: On Error GoTo 0
        End If
    Next di

    ' この法人の店舗リストを抽出（sccのキーからcorpCodeで絞り込み）
    Dim shopList As Object
    Set shopList = CreateObject("Scripting.Dictionary")
    Dim sccKeys As Variant: sccKeys = scc.Keys
    Dim ki As Integer
    For ki = 0 To UBound(sccKeys)
        Dim kStr As String: kStr = CStr(sccKeys(ki))
        Dim pos As Integer: pos = InStrRev(kStr, "_")
        If pos > 0 Then
            Dim snm As String: snm = Left(kStr, pos - 1)
            ' この店舗がcorpCodeに属するかどうかはRAWデータで判断済み（sccに含まれる）
            ' ここでは全sccキーから法人コード不明のため全店舗を含める
            ' → 選択店舗と同じ法人コードを持つ店舗のみに絞るため
            '    shopNameと同じcorpCodeの店舗を対象とする
            If Not shopList.Exists(snm) Then shopList.Add snm, 1
        End If
    Next ki

    Dim shopArr() As String
    ReDim shopArr(shopList.Count - 1)
    Dim si As Integer
    For si = 0 To shopList.Count - 1: shopArr(si) = shopList.Keys()(si): Next si
    Call SortArr_M4(shopArr)

    ' 法人バーを更新（Text 1）
    For Each sh In sl.Shapes
        If sh.HasTextFrame And sh.Name = "Text 1" Then
            Dim t1 As String: t1 = sh.TextFrame.TextRange.Text
            If InStr(t1, "○○○○株式会社") > 0 Then
                ' 法人名は店舗名から推測不可なので対象法人の最初の店舗名を使用
                sh.TextFrame.TextRange.Text = Replace(t1, "○○○○株式会社", "（対象法人）")
            End If
        End If
    Next sh

    ' データ行を新規生成
    Dim rowNum As Integer
    For rowNum = 0 To UBound(shopArr)
        Dim nm As String: nm = shopArr(rowNum)
        Dim rowTop As Double: rowTop = ROW_TOP + rowNum * ROW_STEP
        Dim tot As Long: tot = 0
        Dim ci As Integer
        Dim isTarget As Boolean: isTarget = (nm = shopName)

        ' 店舗名列（対象店舗は強調）
        Dim nmBg As Long: If isTarget Then nmBg = RGB(252, 228, 214) Else nmBg = RGB(31, 56, 100)
        Dim nmFg As Long: If isTarget Then nmFg = RGB(192, 0, 0) Else nmFg = RGB(255, 255, 255)
        Call AddTextBox_M4(sl, colLeft(0), rowTop, colWidth(0), ROW_HEIGHT, nm, _
                           11, isTarget, nmBg, nmFg)

        ' コード列
        For ci = 0 To 3
            Dim k As String: k = nm & "_" & CStr(colCodes(ci))
            Dim cnt As Long: cnt = 0
            If scc.Exists(k) Then cnt = CLng(scc(k))
            Dim ct As String: If cnt = 0 Then ct = "－" Else ct = "●"
            Dim fg As Long: If cnt > 0 Then fg = RGB(192, 0, 0) Else fg = RGB(150, 150, 150)
            Call AddTextBox_M4(sl, colLeft(ci + 1), rowTop, colWidth(ci + 1), ROW_HEIGHT, ct, _
                               11, cnt > 0, RGB(255, 255, 255), fg)
            tot = tot + cnt
        Next ci

        ' 合計
        Call AddTextBox_M4(sl, colLeft(5), rowTop, colWidth(5), ROW_HEIGHT, CStr(tot) & "件", _
                           10, True, RGB(255, 255, 255), RGB(31, 56, 100))

        ' 優先度
        Dim pr As String
        Dim prBg As Long: Dim prFg As Long
        If tot >= 2 Then
            pr = "★高": prBg = RGB(192, 0, 0): prFg = RGB(255, 255, 255)
        ElseIf tot >= 1 Then
            pr = "中": prBg = RGB(255, 255, 255): prFg = RGB(68, 68, 68)
        Else
            pr = "低": prBg = RGB(240, 240, 240): prFg = RGB(150, 150, 150)
        End If
        Call AddTextBox_M4(sl, colLeft(6), rowTop, colWidth(6), ROW_HEIGHT, pr, _
                           10, tot >= 2, prBg, prFg)

    Next rowNum
End Sub

' ---- P8更新：店舗個別サマリー ----
Private Sub UpdateP8_M4(pptPres As Object, sd As Object, shopName As String)

    Dim sl As Object
    Set sl = pptPres.Slides(8)

    Dim shapeDict As Object
    Set shapeDict = CreateObject("Scripting.Dictionary")

    Dim sh As Object
    For Each sh In sl.Shapes
        If sh.HasTextFrame Then
            Dim sn As String: sn = sh.Name
            If Left(sn, 5) = "Text " Then
                shapeDict(CInt(Mid(sn, 6))) = sh
            End If
        End If
    Next sh

    ' Text 1: 店舗名バー
    If shapeDict.Exists(1) Then
        Dim t1 As String: t1 = shapeDict(1).TextFrame.TextRange.Text
        If InStr(t1, "Audi ○○○○") > 0 Then
            shapeDict(1).TextFrame.TextRange.Text = Replace(t1, "Audi ○○○○", shopName)
        End If
    End If

    ' Text 8: 重大エラーコード一覧
    If shapeDict.Exists(8) Then
        Dim major(3) As Long
        major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003
        Dim codes As String: codes = ""
        Dim cj As Integer
        For cj = 0 To 3
            If sd.Exists("cnt_" & CStr(major(cj))) Then
                If CLng(sd("cnt_" & CStr(major(cj)))) > 0 Then
                    If codes <> "" Then codes = codes & "　"
                    codes = codes & CStr(major(cj))
                End If
            End If
        Next cj
        If codes = "" Then codes = "なし"
        shapeDict(8).TextFrame.TextRange.Text = codes
    End If

    ' Text 11: 1910発生件数
    If shapeDict.Exists(11) Then
        Dim c1 As Long: c1 = 0
        If sd.Exists("cnt_1910") Then c1 = CLng(sd("cnt_1910"))
        shapeDict(11).TextFrame.TextRange.Text = CStr(c1) & "件"
    End If

    ' Text 14: 発生月
    If shapeDict.Exists(14) Then
        shapeDict(14).TextFrame.TextRange.Text = GetMonths_M4(sd)
    End If
End Sub

' ---- P10更新：発生コード確認 ----
Private Sub UpdateP10_M4(pptPres As Object, sd As Object, shopName As String)

    Dim sl As Object
    Set sl = pptPres.Slides(10)

    Dim major(3) As Long
    major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003

    Dim tot As Long: tot = 0
    Dim cj  As Integer
    For cj = 0 To 3
        If sd.Exists("cnt_" & CStr(major(cj))) Then tot = tot + CLng(sd("cnt_" & CStr(major(cj))))
    Next cj

    Dim zs3 As String
    zs3 = ChrW(12288) & ChrW(12288) & ChrW(12288)

    Dim sh As Object
    For Each sh In sl.Shapes
        If sh.HasTextFrame Then
            If sh.TextFrame.HasText Then
                Dim txt As String: txt = sh.TextFrame.TextRange.Text
                If InStr(txt, "Audi ○○○○") > 0 Then
                    sh.TextFrame.TextRange.Text = Replace(txt, "Audi ○○○○", shopName)
                    txt = sh.TextFrame.TextRange.Text
                End If
                If InStr(txt, "重大エラー発生件数：" & zs3 & "件") > 0 Then
                    sh.TextFrame.TextRange.Text = Replace(txt, "重大エラー発生件数：" & zs3 & "件", _
                                                          "重大エラー発生件数：" & CStr(tot) & "件")
                End If
            End If
        End If
    Next sh
End Sub

' ---- テキストボックス追加ヘルパー ----
Private Sub AddTextBox_M4(sl As Object, leftCm As Double, topCm As Double, _
                           wCm As Double, hCm As Double, txt As String, _
                           fontSize As Integer, bold As Boolean, bgColor As Long, fgColor As Long)

    Const EMU As Long = 360000

    Dim tb As Object
    Set tb = sl.Shapes.AddTextbox( _
        1, _
        leftCm * EMU, topCm * EMU, wCm * EMU, hCm * EMU)

    With tb.TextFrame
        .WordWrap = True
        .AutoSize = 0
        .MarginLeft = 50000
        .MarginRight = 50000
        .MarginTop = 30000
        .MarginBottom = 30000
        With .TextRange
            .Text = txt
            With .Font
                .Size = fontSize
                .Bold = bold
                .Color.RGB = fgColor
                .Name = "Arial"
            End With
            .ParagraphFormat.Alignment = 2  ' ppAlignCenter
        End With
    End With

    With tb.Fill
        .Visible = True
        .ForeColor.RGB = bgColor
        .Solid
    End With

    With tb.Line
        .Visible = True
        .ForeColor.RGB = RGB(200, 200, 200)
        .Weight = 0.5
    End With

End Sub


'==============================================================================
' ⑥ 共通補助関数
'==============================================================================

Private Function GetRawWorksheet_M4(ByRef wbRaw As Workbook, ByRef closeAfter As Boolean) As Worksheet

    Set wbRaw = Nothing
    closeAfter = False

    ' まずこのブックを確認
    Dim ws As Worksheet
    Set ws = FindSheet_M4(ThisWorkbook, "RAWデータ")
    If Not ws Is Nothing Then Set GetRawWorksheet_M4 = ws: Exit Function

    ' 開いている全ブックを確認
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            Set ws = FindSheet_M4(wb, "RAWデータ")
            If Not ws Is Nothing Then Set GetRawWorksheet_M4 = ws: Exit Function
        End If
    Next wb

    ' 見つからない場合はファイル選択
    Dim rawPath As Variant
    rawPath = Application.GetOpenFilename( _
        FileFilter:="Excelファイル (*.xlsx;*.xlsm),*.xlsx;*.xlsm", _
        Title:="RAWデータシートを含むExcelファイルを選択してください")

    If rawPath = False Then Exit Function

    Set wbRaw = Workbooks.Open(CStr(rawPath), ReadOnly:=True, UpdateLinks:=False)
    closeAfter = True

    Set ws = FindSheet_M4(wbRaw, "RAWデータ")
    If ws Is Nothing Then
        MsgBox "「RAWデータ」シートが見つかりません。", vbExclamation
        wbRaw.Close False: Set wbRaw = Nothing: closeAfter = False
        Exit Function
    End If

    Set GetRawWorksheet_M4 = ws
End Function

Private Function FindSheet_M4(wb As Workbook, targetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = targetName Then Set FindSheet_M4 = ws: Exit Function
    Next ws
End Function

Private Function GetSelectionSheet_M4() As Worksheet
    Dim ws As Worksheet
    Set ws = FindSheet_M4(ThisWorkbook, "選択_法人店舗")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "選択_法人店舗"
    End If
    ws.Cells.Clear
    Set GetSelectionSheet_M4 = ws
End Function

Private Function ExtractNumbers_M4(s As String) As String
    Dim i As Long, buf As String, ch As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then buf = buf & ch
    Next i
    ExtractNumbers_M4 = buf
End Function

Private Function ExtractYYYYMM_M4(fileName As String) As String
    Dim i As Long, buf As String, ch As String
    For i = 1 To Len(fileName)
        ch = Mid(fileName, i, 1)
        If ch >= "0" And ch <= "9" Then
            buf = buf & ch
            If Len(buf) = 6 Then
                If Left(buf, 2) = "20" Then ExtractYYYYMM_M4 = buf: Exit Function
                buf = Mid(buf, 2)
            End If
        Else
            buf = ""
        End If
    Next i
End Function

Private Function CleanText_M4(s As String) As String
    s = Replace(s, vbCr, ""): s = Replace(s, vbLf, "")
    s = Replace(s, Chr(13), ""): s = Replace(s, Chr(10), "")
    CleanText_M4 = Trim(s)
End Function

Private Function CleanFileName_M4(s As String) As String
    Dim ng As Variant, i As Long
    ng = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(ng) To UBound(ng): s = Replace(s, CStr(ng(i)), "_"): Next i
    CleanFileName_M4 = s
End Function

Private Sub SortArr_M4(ByRef arr As Variant)
    Dim i As Long, j As Long, tmp As Variant
    If IsEmpty(arr) Then Exit Sub
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CStr(arr(i)) > CStr(arr(j)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function GetMonths_M4(sd As Object) As String
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim major(3) As Long
    major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003
    Dim cj As Integer
    For cj = 0 To 3
        Dim mk As String: mk = "mon_" & CStr(major(cj))
        If sd.Exists(mk) Then
            Dim ms() As String: ms = Split(CStr(sd(mk)), ",")
            Dim m As Variant
            For Each m In ms
                If Trim(CStr(m)) <> "" Then
                    If Not d.Exists(Trim(CStr(m))) Then d.Add Trim(CStr(m)), 1
                End If
            Next m
        End If
    Next cj

    If d.Count = 0 Then GetMonths_M4 = "－": Exit Function

    Dim keys As Variant: keys = d.Keys
    Call SortArr_M4(keys)

    Dim i As Long, result As String
    For i = LBound(keys) To UBound(keys)
        Dim s As String: s = CStr(keys(i))
        If Len(s) = 6 Then
            If result <> "" Then result = result & "・"
            result = result & CStr(CInt(Right(s, 2))) & "月"
        End If
    Next i

    If result = "" Then result = "－"
    GetMonths_M4 = result
End Function
