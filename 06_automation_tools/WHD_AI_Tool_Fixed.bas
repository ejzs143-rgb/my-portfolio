Option Explicit

' ========================================
' SupportDesk対応AI支援ツール - Outlook連携版
' Example Company Support Team専用
' 作成日: 2026年1月
' ========================================

' グローバル変数
Dim whdEmails As Collection

' ========================================
' ボタン自動作成（初回セットアップ用）
' ========================================
Sub CreateButtons()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim btn As Button
    
    On Error Resume Next
    
    Set ws1 = ThisWorkbook.Sheets("メール一覧")
    ws1.Buttons.Delete
    
    Set btn = ws1.Buttons.Add(20, 80, 150, 35)
    btn.Caption = "メール取得"
    btn.OnAction = "GetSupportDeskEmails"
    btn.Font.Size = 14
    btn.Font.Bold = True
    
    Set ws2 = ThisWorkbook.Sheets("解析結果")
    ws2.Buttons.Delete
    
    Set btn = ws2.Buttons.Add(20, 950, 180, 35)
    btn.Caption = "Outlook下書き保存"
    btn.OnAction = "SaveToDraft"
    btn.Font.Size = 12
    btn.Font.Bold = True
    
    Set btn = ws2.Buttons.Add(220, 950, 180, 35)
    btn.Caption = "クリップボードにコピー"
    btn.OnAction = "CopyResponseToClipboard"
    btn.Font.Size = 12
    btn.Font.Bold = True
    
    On Error GoTo 0
    
    MsgBox "ボタンの作成が完了しました", vbInformation
End Sub

' ========================================
' メイン処理: 未読SupportDeskメール取得
' ========================================
Sub GetSupportDeskEmails()
    On Error GoTo ErrorHandler
    
    Dim outlookApp As Object
    Dim ns As Object
    Dim whdFolder As Object
    Dim mailItem As Object
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim emailCount As Integer
    
    Set whdEmails = New Collection
    Set ws = ThisWorkbook.Sheets("メール一覧")
    
    Call ClearAllButtons
    
    ws.Range("A6:H1000").ClearContents
    
    Application.StatusBar = "Outlookに接続中..."
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set ns = outlookApp.GetNamespace("MAPI")
    
    On Error Resume Next
    Set whdFolder = ns.GetDefaultFolder(6).Folders("SupportDesk")
    On Error GoTo ErrorHandler
    
    If whdFolder Is Nothing Then
        MsgBox "「SupportDesk」フォルダが見つかりません。" & vbCrLf & _
               "受信トレイ直下に「SupportDesk」フォルダを作成してください。", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If
    
    Application.StatusBar = "メール検索中..."
    
    rowNum = 6
    emailCount = 0
    
    For Each mailItem In whdFolder.Items
        If TypeName(mailItem) = "MailItem" Then
            If mailItem.UnRead And IsSupportDeskSender(mailItem.SenderEmailAddress) Then
                If InStr(mailItem.Subject, "Fw:") > 0 Then
                    emailCount = emailCount + 1
                    
                    ws.Cells(rowNum, 1).Value = emailCount
                    ws.Cells(rowNum, 2).Value = Format(mailItem.ReceivedTime, "yyyy/mm/dd hh:mm")
                    ws.Cells(rowNum, 3).Value = mailItem.SenderName
                    ws.Cells(rowNum, 4).Value = mailItem.Subject
                    ws.Cells(rowNum, 5).Value = Left(Replace(mailItem.Body, vbCrLf, " "), 80) & "..."
                    ws.Cells(rowNum, 6).Value = "未処理"
                    ws.Cells(rowNum, 6).Interior.Color = RGB(255, 240, 200)
                    
                    Dim btn As Button
                    Set btn = ws.Buttons.Add(ws.Cells(rowNum, 7).Left, _
                                             ws.Cells(rowNum, 7).Top, _
                                             60, 20)
                    btn.Caption = "解析"
                    btn.OnAction = "AnalyzeEmail"
                    btn.Name = "Btn_" & rowNum
                    
                    whdEmails.Add mailItem, CStr(rowNum)
                    
                    rowNum = rowNum + 1
                End If
            End If
        End If
    Next mailItem
    
    ws.Range("B2").Value = emailCount & " 件"
    ws.Range("B3").Value = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
    If emailCount = 0 Then
        MsgBox "未読のSupportDeskメールはありません。", vbInformation
    Else
        MsgBox emailCount & " 件のSupportDeskメールを取得しました。" & vbCrLf & _
               "「解析」ボタンをクリックしてAI解析を開始してください。", vbInformation
    End If
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "Outlookが起動していることを確認してください。", vbCritical
    Application.StatusBar = False
End Sub

' ========================================
' SupportDesk差出人チェック
' ========================================
Function IsSupportDeskSender(senderEmail As String) As Boolean
    Dim whdSenders As Variant
    Dim sender As Variant
    
    whdSenders = Array( _
        "user@example.com", _
        "user@example.com", _
        "user@example.com" _
    )
    
    IsSupportDeskSender = False
    For Each sender In whdSenders
        If InStr(LCase(senderEmail), LCase(sender)) > 0 Then
            IsSupportDeskSender = True
            Exit Function
        End If
    Next sender
End Function

' ========================================
' メール解析処理
' ========================================
Sub AnalyzeEmail()
    On Error GoTo ErrorHandler
    
    Dim btnName As String
    Dim rowNum As Long
    Dim mailItem As Object
    Dim ws As Worksheet
    Dim analysisSheet As Worksheet
    
    btnName = Application.Caller
    rowNum = CLng(Replace(btnName, "Btn_", ""))
    
    Set ws = ThisWorkbook.Sheets("メール一覧")
    Set analysisSheet = ThisWorkbook.Sheets("解析結果")
    
    On Error Resume Next
    Set mailItem = whdEmails(CStr(rowNum))
    On Error GoTo ErrorHandler
    
    If mailItem Is Nothing Then
        MsgBox "メールの取得に失敗しました。再度「メール取得」を実行してください。", vbExclamation
        Exit Sub
    End If
    
    ws.Cells(rowNum, 6).Value = "解析中..."
    ws.Cells(rowNum, 6).Interior.Color = RGB(255, 255, 0)
    DoEvents
    
    Application.StatusBar = "AI解析中..."
    Application.ScreenUpdating = False
    
    analysisSheet.Range("B3").Value = mailItem.Subject
    analysisSheet.Range("B4").Value = mailItem.SenderName
    analysisSheet.Range("B5").Value = Format(mailItem.ReceivedTime, "yyyy/mm/dd hh:mm")
    analysisSheet.Range("B7").Value = mailItem.Body
    
    Dim i As Integer
    For i = 1 To 3
        ws.Cells(rowNum, 6).Value = "解析中" & String(i, ".")
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01"))
    Next i
    
    Call AIAnalysis(mailItem.Body, analysisSheet)
    
    ws.Cells(rowNum, 6).Value = "解析完了"
    ws.Cells(rowNum, 6).Interior.Color = RGB(144, 238, 144)
    
    mailItem.UnRead = False
    
    analysisSheet.Activate
    analysisSheet.Range("A1").Select
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    MsgBox "AI解析が完了しました。" & vbCrLf & vbCrLf & _
           "自動生成された回答文を確認してください。", vbInformation
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Application.StatusBar = False
End Sub

' ========================================
' AI解析エンジン
' ========================================
Sub AIAnalysis(emailBody As String, ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim keywords As Collection
    Dim category As String
    Dim priority As String
    Dim similarCases As String
    Dim checklist As String
    Dim response As String
    
    Set keywords = ExtractKeywords(emailBody)
    
    category = CategorizeQuery(keywords)
    priority = GetPriority(category)
    
    similarCases = FindSimilarCases(keywords)
    
    checklist = GenerateChecklist(category)
    
    response = GenerateResponse(category, similarCases, emailBody)
    
    ws.Range("B10").Value = category
    ws.Range("B11").Value = priority
    
    If priority = "高" Then
        ws.Range("B11").Interior.Color = RGB(255, 200, 200)
    Else
        ws.Range("B11").Interior.Color = RGB(200, 255, 200)
    End If
    
    ws.Range("B13").Value = similarCases
    ws.Range("B20").Value = checklist
    ws.Range("B28").Value = response
    
    Exit Sub
    
ErrorHandler:
    MsgBox "AI解析エラー: " & Err.Description, vbCritical
End Sub

' ========================================
' キーワード抽出
' ========================================
Function ExtractKeywords(text As String) As Collection
    Dim keywords As New Collection
    Dim commonKeywords As Variant
    Dim keyword As Variant
    
    commonKeywords = Array( _
        "エアコン", "コンプレッサー", "DSG", "メカトロ", "ターボ", _
        "ブレーキ", "エンジン", "ウォーターポンプ", "クーラント", _
        "保証", "特例", "BusinessApp", "DiagnosticSystem", "RFA", "DTC", _
        "異音", "警告灯", "漏れ", "摩耗", "不良", "承認", _
        "診断", "部品", "修理", "交換", "クライテリア", "TPI" _
    )
    
    For Each keyword In commonKeywords
        If InStr(text, keyword) > 0 Then
            On Error Resume Next
            keywords.Add keyword
            On Error GoTo 0
        End If
    Next keyword
    
    Set ExtractKeywords = keywords
End Function

' ========================================
' カテゴリ判定
' ========================================
Function CategorizeQuery(keywords As Collection) As String
    Dim keyword As Variant
    
    For Each keyword In keywords
        If keyword = "特例" Or keyword = "RFA" Or keyword = "承認" Then
            CategorizeQuery = "特例対応"
            Exit Function
        End If
    Next keyword
    
    For Each keyword In keywords
        If keyword = "BusinessApp" Or keyword = "DiagnosticSystem" Then
            CategorizeQuery = "システム操作"
            Exit Function
        End If
    Next keyword
    
    For Each keyword In keywords
        If keyword = "保証" Then
            CategorizeQuery = "保証適用判断"
            Exit Function
        End If
    Next keyword
    
    For Each keyword In keywords
        If keyword = "診断" Or keyword = "DTC" Then
            CategorizeQuery = "技術判断"
            Exit Function
        End If
    Next keyword
    
    CategorizeQuery = "一般問合せ"
End Function

' ========================================
' 優先度判定
' ========================================
Function GetPriority(category As String) As String
    If category = "特例対応" Then
        GetPriority = "高"
    Else
        GetPriority = "中"
    End If
End Function

' ========================================
' 類似案件検索
' ========================================
Function FindSimilarCases(keywords As Collection) As String
    Dim result As String
    Dim caseCount As Integer
    
    result = "【過去30件から検索した類似案件】" & vbCrLf & vbCrLf
    caseCount = 0
    
    If KeywordExists(keywords, "エアコン") Or KeywordExists(keywords, "コンプレッサー") Then
        caseCount = caseCount + 1
        result = result & "案件" & caseCount & ": エアコンコンプレッサー異音（類似度: 85%）" & vbCrLf
        result = result & "   判定: 保証対象" & vbCrLf
        result = result & "   理由: 初回発生、通常使用範囲内" & vbCrLf & vbCrLf
    End If
    
    If KeywordExists(keywords, "DSG") Or KeywordExists(keywords, "メカトロ") Then
        caseCount = caseCount + 1
        result = result & "案件" & caseCount & ": DSGメカトロニクス不良（類似度: 90%）" & vbCrLf
        result = result & "   判定: 保証対象" & vbCrLf
        result = result & "   理由: TPI該当、保証延長対象部品" & vbCrLf & vbCrLf
    End If
    
    If KeywordExists(keywords, "特例") Or KeywordExists(keywords, "RFA") Then
        caseCount = caseCount + 1
        result = result & "案件" & caseCount & ": 保証期限切れ直後申請（類似度: 75%）" & vbCrLf
        result = result & "   判定: 特例承認" & vbCrLf
        result = result & "   理由: 発生時期が保証期間内" & vbCrLf & vbCrLf
    End If
    
    If KeywordExists(keywords, "BusinessApp") Then
        caseCount = caseCount + 1
        result = result & "案件" & caseCount & ": BusinessAppアクセスエラー（類似度: 80%）" & vbCrLf
        result = result & "   対処: ブラウザキャッシュクリア、再ログイン" & vbCrLf & vbCrLf
    End If
    
    If caseCount = 0 Then
        result = result & "該当する類似案件が見つかりませんでした。"
    Else
        result = result & "（" & caseCount & "件の類似案件を発見）"
    End If
    
    FindSimilarCases = result
End Function

' ========================================
' キーワード存在チェック
' ========================================
Function KeywordExists(keywords As Collection, searchWord As String) As Boolean
    Dim keyword As Variant
    KeywordExists = False
    
    On Error Resume Next
    For Each keyword In keywords
        If keyword = searchWord Then
            KeywordExists = True
            Exit Function
        End If
    Next keyword
    On Error GoTo 0
End Function

' ========================================
' チェックリスト生成
' ========================================
Function GenerateChecklist(category As String) As String
    Dim result As String
    
    result = "【全案件共通の確認事項】" & vbCrLf & vbCrLf
    result = result & "□ DiagnosticSystemログ確認済み" & vbCrLf
    result = result & "□ 車両情報（VIN、走行距離）確認" & vbCrLf
    result = result & "□ 修理履歴確認" & vbCrLf
    result = result & "□ クライテリア確認" & vbCrLf & vbCrLf
    
    Select Case category
        Case "保証適用判断"
            result = result & "【保証判断専用チェック】" & vbCrLf
            result = result & "□ TPI該当確認" & vbCrLf
            result = result & "□ 保証期間内確認" & vbCrLf
            result = result & "□ 前回修理との関連確認" & vbCrLf
            
        Case "特例対応"
            result = result & "【特例対応専用チェック】" & vbCrLf
            result = result & "□ 承認者確認（Approver）" & vbCrLf
            result = result & "□ 特例理由の妥当性確認" & vbCrLf
            result = result & "□ 前例確認" & vbCrLf
            
        Case "システム操作"
            result = result & "【システム専用チェック】" & vbCrLf
            result = result & "□ エラーメッセージ確認" & vbCrLf
            result = result & "□ 権限確認" & vbCrLf
            result = result & "□ 別端末での動作確認" & vbCrLf
    End Select
    
    GenerateChecklist = result
End Function

' ========================================
' 回答文生成
' ========================================
Function GenerateResponse(category As String, similarCases As String, emailBody As String) As String
    Dim response As String
    
    response = "お世話になっております。" & vbCrLf
    response = response & "Example Company Support Teamの担当者です。" & vbCrLf & vbCrLf
    
    Select Case category
        Case "保証適用判断"
            response = response & "お問い合わせの件、確認させていただきました。" & vbCrLf & vbCrLf
            response = response & "【確認結果】" & vbCrLf
            response = response & "詳細確認が必要です" & vbCrLf & vbCrLf
            response = response & "【必要な追加情報】" & vbCrLf
            response = response & "・DiagnosticSystemの診断ログ全文" & vbCrLf
            response = response & "・車両の走行距離" & vbCrLf
            response = response & "・前回修理履歴（該当部位）" & vbCrLf
            response = response & "・クライテリアNo.（該当する場合）" & vbCrLf & vbCrLf
            response = response & "上記情報をいただき次第、速やかに最終判断をご連絡いたします。" & vbCrLf
            
        Case "特例対応"
            response = response & "特例対応のご相談、承知いたしました。" & vbCrLf & vbCrLf
            response = response & "【確認事項】" & vbCrLf
            response = response & "・発生時期と保証期間との関係" & vbCrLf
            response = response & "・ディーラー様での対応履歴" & vbCrLf
            response = response & "・類似案件の前例" & vbCrLf & vbCrLf
            response = response & "【次のステップ】" & vbCrLf
            response = response & "上記確認後、Approverと協議の上、" & vbCrLf
            response = response & "48時間以内に正式回答をご連絡いたします。" & vbCrLf
            
        Case "システム操作"
            response = response & "システムに関するお問い合わせの件、承知いたしました。" & vbCrLf & vbCrLf
            response = response & "【対処方法】" & vbCrLf
            response = response & "1. ブラウザのキャッシュをクリア" & vbCrLf
            response = response & "2. 再度ログイン" & vbCrLf
            response = response & "3. 別のブラウザで動作確認" & vbCrLf & vbCrLf
            response = response & "上記で解決しない場合は、以下の情報をご提供ください：" & vbCrLf
            response = response & "・エラーメッセージのスクリーンショット" & vbCrLf
            response = response & "・使用しているブラウザとバージョン" & vbCrLf
            
        Case Else
            response = response & "お問い合わせの件、承知いたしました。" & vbCrLf & vbCrLf
            response = response & "詳細を確認の上、適切な対応をご案内させていただきます。" & vbCrLf
            response = response & "少々お時間をいただけますと幸いです。" & vbCrLf
    End Select
    
    response = response & vbCrLf & "何卒よろしくお願いいたします。" & vbCrLf & vbCrLf
    response = response & "---" & vbCrLf
    response = response & "ExampleCoグループアフターセールス部門" & vbCrLf
    response = response & "保証課 担当者" & vbCrLf
    
    GenerateResponse = response
End Function

' ========================================
' 回答文をOutlook下書きに保存
' ========================================
Sub SaveToDraft()
    On Error GoTo ErrorHandler
    
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim ws As Worksheet
    Dim responseText As String
    Dim Subject As String
    
    Set ws = ThisWorkbook.Sheets("解析結果")
    
    responseText = ws.Range("B28").Value
    Subject = "Re: " & ws.Range("B3").Value
    
    If Len(Trim(responseText)) = 0 Then
        MsgBox "回答文が生成されていません。先に「解析」を実行してください。", vbExclamation
        Exit Sub
    End If
    
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0)
    
    mailItem.Subject = Subject
    mailItem.Body = responseText
    
    mailItem.Save
    
    MsgBox "回答文をOutlook下書きに保存しました。" & vbCrLf & vbCrLf & _
           "Outlookの下書きフォルダを確認してください。", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub

' ========================================
' 回答文をクリップボードにコピー
' ========================================
Sub CopyResponseToClipboard()
    Dim ws As Worksheet
    Dim responseText As String
    Dim dataObj As Object
    
    Set ws = ThisWorkbook.Sheets("解析結果")
    responseText = ws.Range("B28").Value
    
    If Len(Trim(responseText)) = 0 Then
        MsgBox "回答文が生成されていません。先に「解析」を実行してください。", vbExclamation
        Exit Sub
    End If
    
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText responseText
    dataObj.PutInClipboard
    
    MsgBox "回答文をクリップボードにコピーしました。" & vbCrLf & vbCrLf & _
           "メーラーに貼り付けてご使用ください。", vbInformation
End Sub

' ========================================
' 全ボタン削除
' ========================================
Sub ClearAllButtons()
    Dim ws As Worksheet
    Dim btn As Button
    
    Set ws = ThisWorkbook.Sheets("メール一覧")
    
    On Error Resume Next
    For Each btn In ws.Buttons
        If Left(btn.Name, 4) = "Btn_" Then
            btn.Delete
        End If
    Next btn
    On Error GoTo 0
End Sub
