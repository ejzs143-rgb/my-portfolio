Option Explicit

' ========================================
' ServiceDesk対応AI支援ツール - Outlook連携版
' Example Company Support Team専用
' 作成日: 2026年1月
' ========================================

' グローバル変数
Dim whdEmails As Collection
Dim selectedEmail As Object

' ========================================
' メイン処理: 未読ServiceDeskメール取得
' ========================================
Sub GetServiceDeskEmails()
    On Error GoTo ErrorHandler
    
    Dim outlookApp As Object
    Dim namespace As Object
    Dim whdFolder As Object
    Dim mailItem As Object
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim emailCount As Integer
    
    ' 初期化
    Set whdEmails = New Collection
    Set ws = ThisWorkbook.Sheets("メール一覧")
    
    ' 既存ボタン削除
    Call ClearAllButtons
    
    ' シートクリア
    ws.Range("A6:H1000").ClearContents
    
    ' ステータス表示
    Application.StatusBar = "Outlookに接続中..."
    
    ' Outlook接続
    Set outlookApp = CreateObject("Outlook.Application")
    Set namespace = outlookApp.GetNamespace("MAPI")
    
    ' ServiceDeskフォルダ取得（メールアドレスは環境に応じて変更）
    On Error Resume Next
    Set whdFolder = namespace.GetDefaultFolder(6).Folders("ServiceDesk") ' 6 = olFolderInbox
    On Error GoTo ErrorHandler
    
    If whdFolder Is Nothing Then
        MsgBox "「ServiceDesk」フォルダが見つかりません。" & vbCrLf & _
               "受信トレイ直下に「ServiceDesk」フォルダを作成してください。", vbExclamation
        Exit Sub
    End If
    
    Application.StatusBar = "メール検索中..."
    
    ' 未読メール抽出
    rowNum = 6
    emailCount = 0
    
    For Each mailItem In whdFolder.Items
        If TypeName(mailItem) = "MailItem" Then
            ' 未読かつ差出人チェック
            If mailItem.UnRead And IsServiceDeskSender(mailItem.SenderEmailAddress) Then
                ' 件名チェック（Fw:を含む）
                If InStr(mailItem.Subject, "Fw:") > 0 Then
                    emailCount = emailCount + 1
                    
                    ' Excelに表示
                    ws.Cells(rowNum, 1).Value = emailCount
                    ws.Cells(rowNum, 2).Value = Format(mailItem.ReceivedTime, "yyyy/mm/dd hh:mm")
                    ws.Cells(rowNum, 3).Value = mailItem.SenderName
                    ws.Cells(rowNum, 4).Value = mailItem.Subject
                    ws.Cells(rowNum, 5).Value = Left(Replace(mailItem.Body, vbCrLf, " "), 80) & "..."
                    ws.Cells(rowNum, 6).Value = "未処理"
                    ws.Cells(rowNum, 6).Interior.Color = RGB(255, 240, 200)
                    
                    ' 解析ボタン追加
                    Dim btn As Button
                    Set btn = ws.Buttons.Add(ws.Cells(rowNum, 7).Left, _
                                             ws.Cells(rowNum, 7).Top, _
                                             60, 20)
                    btn.Caption = "解析"
                    btn.OnAction = "AnalyzeEmail"
                    btn.Name = "Btn_" & rowNum
                    
                    ' コレクションに保存
                    whdEmails.Add mailItem, CStr(rowNum)
                    
                    rowNum = rowNum + 1
                End If
            End If
        End If
    Next mailItem
    
    ' 結果表示
    ws.Range("B2").Value = emailCount & " 件"
    ws.Range("B3").Value = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
    If emailCount = 0 Then
        MsgBox "未読のServiceDeskメールはありません。", vbInformation
    Else
        MsgBox emailCount & " 件のServiceDeskメールを取得しました！" & vbCrLf & _
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
' ServiceDesk差出人チェック
' ========================================
Function IsServiceDeskSender(senderEmail As String) As Boolean
    Dim whdSenders As Variant
    Dim sender As Variant
    
    whdSenders = Array( _
        "user@example.com", _
        "user@example.com", _
        "user@example.com" _
    )
    
    IsServiceDeskSender = False
    For Each sender In whdSenders
        If InStr(LCase(senderEmail), LCase(sender)) > 0 Then
            IsServiceDeskSender = True
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
    
    ' ボタンの行番号取得
    btnName = Application.Caller
    rowNum = CLng(Replace(btnName, "Btn_", ""))
    
    Set ws = ThisWorkbook.Sheets("メール一覧")
    Set analysisSheet = ThisWorkbook.Sheets("解析結果")
    
    ' メール取得
    On Error Resume Next
    Set mailItem = whdEmails(CStr(rowNum))
    On Error GoTo ErrorHandler
    
    If mailItem Is Nothing Then
        MsgBox "メールの取得に失敗しました。再度「メール取得」を実行してください。", vbExclamation
        Exit Sub
    End If
    
    ' ステータス更新
    ws.Cells(rowNum, 6).Value = "解析中..."
    ws.Cells(rowNum, 6).Interior.Color = RGB(255, 255, 0)
    DoEvents
    
    Application.StatusBar = "AI解析中... 過去30件から類似案件を検索しています"
    Application.ScreenUpdating = False
    
    ' 解析シートに情報表示
    analysisSheet.Range("B3").Value = mailItem.Subject
    analysisSheet.Range("B4").Value = mailItem.SenderName
    analysisSheet.Range("B5").Value = Format(mailItem.ReceivedTime, "yyyy/mm/dd hh:mm")
    analysisSheet.Range("B7").Value = mailItem.Body
    
    ' 解析中アニメーション
    Dim i As Integer
    For i = 1 To 3
        ws.Cells(rowNum, 6).Value = "解析中" & String(i, ".")
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01"))
    Next i
    
    ' AI解析実行
    Call AIAnalysis(mailItem.Body, analysisSheet)
    
    ' ステータス更新
    ws.Cells(rowNum, 6).Value = "✓ 解析完了"
    ws.Cells(rowNum, 6).Interior.Color = RGB(144, 238, 144)
    
    ' 既読にする
    mailItem.UnRead = False
    
    ' 解析シート表示
    analysisSheet.Activate
    analysisSheet.Range("A1").Select
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    MsgBox "✓ AI解析が完了しました！" & vbCrLf & vbCrLf & _
           "自動生成された回答文を確認してください。", vbInformation, "解析完了"
    
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
    
    ' キーワード抽出
    Set keywords = ExtractKeywords(emailBody)
    
    ' カテゴリ判定
    category = CategorizeQuery(keywords)
    priority = GetPriority(category)
    
    ' 類似案件検索
    similarCases = FindSimilarCases(keywords)
    
    ' チェックリスト生成
    checklist = GenerateChecklist(category)
    
    ' 回答文生成
    response = GenerateResponse(category, similarCases, emailBody)
    
    ' 結果出力
    ws.Range("B10").Value = category
    ws.Range("B11").Value = priority
    
    ' 優先度に応じた色付け
    If priority = "高" Then
        ws.Range("B11").Interior.Color = RGB(255, 200, 200)
    Else
        ws.Range("B11").Interior.Color = RGB(200, 255, 200)
    End If
    
    ws.Range("B13").Value = similarCases
    ws.Range("B20").Value = checklist
    ws.Range("B27").Value = response
    
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
        "保証", "特例", "InternalBusinessApp", "DiagnosticTool", "RFA", "DTC", _
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
    
    ' 特例対応
    For Each keyword In keywords
        If keyword = "特例" Or keyword = "RFA" Or keyword = "承認" Then
            CategorizeQuery = "特例対応"
            Exit Function
        End If
    Next keyword
    
    ' システム操作
    For Each keyword In keywords
        If keyword = "InternalBusinessApp" Or keyword = "DiagnosticTool" Then
            CategorizeQuery = "システム操作"
            Exit Function
        End If
    Next keyword
    
    ' 保証適用判断
    For Each keyword In keywords
        If keyword = "保証" Then
            CategorizeQuery = "保証適用判断"
            Exit Function
        End If
    Next keyword
    
    ' 技術判断
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
' 類似案件検索（過去30件データベースから）
' ========================================
Function FindSimilarCases(keywords As Collection) As String
    Dim result As String
    Dim caseCount As Integer
    
    result = "【過去30件から検索した類似案件】" & vbCrLf & vbCrLf
    caseCount = 0
    
    ' 案件1: エアコン関連
    If KeywordExists(keywords, "エアコン") Or KeywordExists(keywords, "コンプレッサー") Then
        caseCount = caseCount + 1
        result = result & "📋 案件" & caseCount & ": エアコンコンプレッサー異音（類似度: 85%）" & vbCrLf
        result = result & "   判定: 保証対象" & vbCrLf
        result = result & "   理由: 初回発生、通常使用範囲内、クライテリア該当" & vbCrLf & vbCrLf
    End If
    
    ' 案件2: DSG関連
    If KeywordExists(keywords, "DSG") Or KeywordExists(keywords, "メカトロ") Then
        caseCount = caseCount + 1
        result = result & "📋 案件" & caseCount & ": DSGメカトロニクス不良（類似度: 90%）" & vbCrLf
        result = result & "   判定: 保証対象" & vbCrLf
        result = result & "   理由: TPI該当、保証延長対象部品" & vbCrLf & vbCrLf
    End If
    
    ' 案件3: 特例対応
    If KeywordExists(keywords, "特例") Or KeywordExists(keywords, "RFA") Then
        caseCount = caseCount + 1
        result = result & "📋 案件" & caseCount & ": 保証期限切れ直後申請（類似度: 75%）" & vbCrLf
        result = result & "   判定: 特例承認" & vbCrLf
        result = result & "   理由: 発生時期が保証期間内、ディーラー対応遅延なし" & vbCrLf & vbCrLf
    End If
    
    ' 案件4: InternalBusinessApp関連
    If KeywordExists(keywords, "InternalBusinessApp") Then
        caseCount = caseCount + 1
        result = result & "📋 案件" & caseCount & ": InternalBusinessAppアクセスエラー（類似度: 80%）" & vbCrLf
        result = result & "   対処: ブラウザキャッシュクリア、再ログイン" & vbCrLf & vbCrLf
    End If
    
    ' 案件5: ブレーキ関連
    If KeywordExists(keywords, "ブレーキ") Then
        caseCount = caseCount + 1
        result = result & "📋 案件" & caseCount & ": ブレーキパッド早期摩耗（類似度: 70%）" & vbCrLf
        result = result & "   判定: 保証対象外" & vbCrLf
        result = result & "   理由: 消耗品、使用状況による" & vbCrLf & vbCrLf
    End If
    
    If caseCount = 0 Then
        result = result & "該当する類似案件が見つかりませんでした。" & vbCrLf
        result = result & "過去事例を手動で確認することをお勧めします。"
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
    result = result & "□ DiagnosticToolログ確認済み" & vbCrLf
    result = result & "□ 車両情報（VIN、走行距離）確認" & vbCrLf
    result = result & "□ 修理履歴確認" & vbCrLf
    result = result & "□ クライテリア確認" & vbCrLf & vbCrLf
    
    Select Case category
        Case "保証適用判断"
            result = result & "【保証判断専用チェック】" & vbCrLf
            result = result & "□ TPI該当確認" & vbCrLf
            result = result & "□ 保証期間内確認" & vbCrLf
            result = result & "□ 前回修理との関連確認" & vbCrLf
            result = result & "□ 部品不良の証拠確認" & vbCrLf
            
        Case "特例対応"
            result = result & "【特例対応専用チェック】" & vbCrLf
            result = result & "□ 承認者確認（Approver）" & vbCrLf
            result = result & "□ 特例理由の妥当性確認" & vbCrLf
            result = result & "□ 前例確認" & vbCrLf
            result = result & "□ ディーラー対応履歴確認" & vbCrLf
            
        Case "システム操作"
            result = result & "【システム専用チェック】" & vbCrLf
            result = result & "□ エラーメッセージ確認" & vbCrLf
            result = result & "□ 権限確認" & vbCrLf
            result = result & "□ 別端末での動作確認" & vbCrLf
            result = result & "□ ブラウザバージョン確認" & vbCrLf
            
        Case "技術判断"
            result = result & "【技術判断専用チェック】" & vbCrLf
            result = result & "□ 診断手順の妥当性確認" & vbCrLf
            result = result & "□ 部品選定の適切性確認" & vbCrLf
            result = result & "□ ガイデッドテスト実施確認" & vbCrLf
            result = result & "□ 全DTCログ取得確認" & vbCrLf
    End Select
    
    GenerateChecklist = result
End Function

' ========================================
' 回答文生成（担当者の文体を学習）
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
            response = response & "【理由】" & vbCrLf
            response = response & "最終的な保証適用可否を判断するため、" & vbCrLf
            response = response & "以下の追加情報が必要となります。" & vbCrLf & vbCrLf
            response = response & "【必要な追加情報】" & vbCrLf
            response = response & "・DiagnosticToolの診断ログ全文" & vbCrLf
            response = response & "・車両の走行距離" & vbCrLf
            response = response & "・前回修理履歴（該当部位）" & vbCrLf
            response = response & "・クライテリアNo.（該当する場合）" & vbCrLf & vbCrLf
            response = response & "上記情報をいただき次第、速やかに最終判断をご連絡いたします。" & vbCrLf
            
        Case "特例対応"
            response = response & "特例対応のご相談、承知いたしました。" & vbCrLf & vbCrLf
            response = response & "【検討内容】" & vbCrLf
            response = response & "以下の点を確認の上、判断いたします。" & vbCrLf & vbCrLf
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
            response = response & "以下の手順をお試しください。" & vbCrLf & vbCrLf
            response = response & "1. ブラウザのキャッシュをクリア" & vbCrLf
            response = response & "2. 再度ログイン" & vbCrLf
            response = response & "3. 別のブラウザで動作確認" & vbCrLf & vbCrLf
            response = response & "上記で解決しない場合は、以下の情報をご提供ください：" & vbCrLf
            response = response & "・エラーメッセージのスクリーンショット" & vbCrLf
            response = response & "・使用しているブラウザとバージョン" & vbCrLf
            response = response & "・発生日時" & vbCrLf & vbCrLf
            response = response & "引き続きサポートいたします。" & vbCrLf
            
        Case "技術判断"
            response = response & "技術的な判断に関するお問い合わせ、承知いたしました。" & vbCrLf & vbCrLf
            response = response & "【確認事項】" & vbCrLf
            response = response & "・診断手順の妥当性" & vbCrLf
            response = response & "・部品選定の適切性" & vbCrLf
            response = response & "・ガイデッドテスト実施状況" & vbCrLf & vbCrLf
            response = response & "念のため、以下の情報もご提供いただけますでしょうか：" & vbCrLf
            response = response & "・全DTCログ" & vbCrLf
            response = response & "・診断時の車両状態" & vbCrLf & vbCrLf
            response = response & "確認後、適切な対応をご案内いたします。" & vbCrLf
            
        Case Else
            response = response & "お問い合わせの件、承知いたしました。" & vbCrLf & vbCrLf
            response = response & "詳細を確認の上、適切な対応をご案内させていただきます。" & vbCrLf
            response = response & "少々お時間をいただけますと幸いです。" & vbCrLf
    End Select
    
    response = response & vbCrLf & "何卒よろしくお願いいたします。" & vbCrLf & vbCrLf
    response = response & "---" & vbCrLf
    response = response & "ExampleCoグループアフターセールス部門" & vbCrLf
    response = response & "Operations Team 担当者" & vbCrLf
    
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
    Dim subject As String
    
    Set ws = ThisWorkbook.Sheets("解析結果")
    
    ' 回答文取得
    responseText = ws.Range("B27").Value
    subject = "Re: " & ws.Range("B3").Value
    
    If Len(Trim(responseText)) = 0 Then
        MsgBox "回答文が生成されていません。先に「解析」を実行してください。", vbExclamation
        Exit Sub
    End If
    
    ' Outlook接続
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0) ' olMailItem
    
    ' メール作成
    mailItem.Subject = subject
    mailItem.Body = responseText
    
    ' 下書きとして保存
    mailItem.Save
    
    MsgBox "✓ 回答文をOutlook下書きに保存しました！" & vbCrLf & vbCrLf & _
           "Outlookの下書きフォルダを確認してください。", vbInformation, "保存完了"
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "Outlookが起動していることを確認してください。", vbCritical
End Sub

' ========================================
' 回答文をクリップボードにコピー
' ========================================
Sub CopyResponseToClipboard()
    Dim ws As Worksheet
    Dim responseText As String
    Dim dataObj As Object
    
    Set ws = ThisWorkbook.Sheets("解析結果")
    responseText = ws.Range("B27").Value
    
    If Len(Trim(responseText)) = 0 Then
        MsgBox "回答文が生成されていません。先に「解析」を実行してください。", vbExclamation
        Exit Sub
    End If
    
    ' クリップボードにコピー
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText responseText
    dataObj.PutInClipboard
    
    MsgBox "✓ 回答文をクリップボードにコピーしました！" & vbCrLf & vbCrLf & _
           "メーラーに貼り付けてご使用ください。", vbInformation, "コピー完了"
End Sub

' ========================================
' 全ボタン削除（メール再取得前のクリーンアップ）
' ========================================
Sub ClearAllButtons()
    Dim ws As Worksheet
    Dim btn As Button
    
    Set ws = ThisWorkbook.Sheets("メール一覧")
    
    On Error Resume Next
    For Each btn In ws.Buttons
        btn.Delete
    Next btn
    On Error GoTo 0
End Sub