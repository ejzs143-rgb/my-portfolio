'==============================================================================
' PPT更新処理 修正版
' P4・P6がテーブルではなくテキストボックスシェイプで構成されているため
' シェイプ名（Text N）ベースで更新する
'
' 【PPTのシェイプ構造】
' P4（Slide4）/ P6（Slide6）共通：
'   ヘッダー行（法人名/店舗名,1910,1223,1016,1003,合計,優先度）はText 5〜17（偶数インデックス）
'   データ行1行目：Text 19（法人名/店舗名）, Text 21（1910）, Text 23（1223）,
'                  Text 25（1016）, Text 27（1003）, Text 29（合計）, Text 31（優先度）
'   データ行2行目：Text 33, Text 35, Text 37, Text 39, Text 41, Text 43, Text 45
'   データ行3行目：Text 47, Text 49, Text 51, Text 53, Text 55, Text 57, Text 59
'   ...以降14おきに次の行
'
' P8（Slide8）：
'   「Audi ○○○○」を含むテキストボックスを店舗名に置換
'   テーブルオブジェクトは存在しないため、「1910 発生件数」「発生月」を
'   含むテキストボックスを直接書き換える
'
' P10（Slide10）：
'   「Audi ○○○○」と「重大エラー発生件数：　　　件」（全角スペース3つ）を置換
'==============================================================================

'------------------------------------------------------------------------------
' P4更新：エリア別法人比較（Slide 4）
' シェイプ構造：
'   データ1行目のText番号 = 19（法人名）
'   以降の行は14つずつ増加（Text 33, 47, 61...）
'   各行内の列オフセット: 0=法人名, 2=1910, 4=1223, 6=1016, 8=1003, 10=合計, 12=優先度
'------------------------------------------------------------------------------
Private Sub UpdateP4_M4(pptPres As Object, scc As Object)

    Dim sl As Object
    Set sl = pptPres.Slides(4)

    ' 列順（1910, 1223, 1016, 1003）
    Dim colCodes(3) As Long
    colCodes(0) = 1910
    colCodes(1) = 1223
    colCodes(2) = 1016
    colCodes(3) = 1003

    ' データ行1行目の法人名シェイプのText番号
    Const FIRST_NAME_IDX As Integer = 19
    ' 行間のインデックス差
    Const ROW_STEP As Integer = 14
    ' 最大行数（A〜D法人 = 4行）
    Const MAX_ROWS As Integer = 10

    ' シェイプをインデックス（Text N の N）でアクセスできるDictionaryを作成
    Dim shapeDict As Object
    Set shapeDict = CreateObject("Scripting.Dictionary")

    Dim sh As Object
    For Each sh In sl.Shapes
        If sh.HasTextFrame Then
            ' "Text 19" のような名前からインデックスを取得
            Dim sName As String
            sName = sh.Name
            If Left(sName, 5) = "Text " Then
                Dim idxNum As Integer
                idxNum = CInt(Mid(sName, 6))
                shapeDict(idxNum) = sh
            End If
        End If
    Next sh

    ' 各データ行を処理
    Dim rowNum As Integer
    For rowNum = 0 To MAX_ROWS - 1

        Dim nameIdx As Integer
        nameIdx = FIRST_NAME_IDX + rowNum * ROW_STEP

        ' 法人名シェイプが存在しない場合は終了
        If Not shapeDict.Exists(nameIdx) Then Exit For

        Dim nmShape As Object
        Set nmShape = shapeDict(nameIdx)

        Dim nm As String
        nm = CleanText_M4(nmShape.TextFrame.TextRange.Text)
        If nm = "" Then GoTo NextRowP4

        Dim tot As Long
        tot = 0
        Dim ci As Integer

        For ci = 0 To 3
            Dim colIdx As Integer
            colIdx = nameIdx + 2 + ci * 2  ' 1910=+2, 1223=+4, 1016=+6, 1003=+8

            Dim k As String
            k = nm & "_" & CStr(colCodes(ci))

            Dim cnt As Long
            cnt = 0
            If scc.Exists(k) Then cnt = CLng(scc(k))

            Dim ct As String
            If cnt = 0 Then
                ct = "－"
            ElseIf cnt >= 3 Then
                ct = "●多"
            Else
                ct = "●"
            End If

            If shapeDict.Exists(colIdx) Then
                shapeDict(colIdx).TextFrame.TextRange.Text = ct
            End If

            tot = tot + cnt
        Next ci

        ' 合計（+10）
        Dim totIdx As Integer
        totIdx = nameIdx + 10
        If shapeDict.Exists(totIdx) Then
            shapeDict(totIdx).TextFrame.TextRange.Text = CStr(tot) & "件"
        End If

        ' 優先度（+12）
        Dim prIdx As Integer
        prIdx = nameIdx + 12
        Dim pr As String
        If tot >= 3 Then
            pr = "★高"
        ElseIf tot >= 1 Then
            pr = "中"
        Else
            pr = "低"
        End If
        If shapeDict.Exists(prIdx) Then
            shapeDict(prIdx).TextFrame.TextRange.Text = pr
        End If

NextRowP4:
    Next rowNum

End Sub


'------------------------------------------------------------------------------
' P6更新：店舗別マトリクス（Slide 6）
' シェイプ構造：
'   データ1行目のText番号 = 19（店舗名）
'   以降の行は14つずつ増加
'   各行内の列オフセット: 0=店舗名, 2=1910, 4=1223, 6=1016, 8=1003, 10=合計, 12=優先度
'------------------------------------------------------------------------------
Private Sub UpdateP6_M4(pptPres As Object, scc As Object)

    Dim sl As Object
    Set sl = pptPres.Slides(6)

    Dim colCodes(3) As Long
    colCodes(0) = 1910
    colCodes(1) = 1223
    colCodes(2) = 1016
    colCodes(3) = 1003

    Const FIRST_NAME_IDX As Integer = 19
    Const ROW_STEP As Integer = 14
    Const MAX_ROWS As Integer = 15  ' 最大15店舗まで対応

    ' シェイプDictionaryを構築
    Dim shapeDict As Object
    Set shapeDict = CreateObject("Scripting.Dictionary")

    Dim sh As Object
    For Each sh In sl.Shapes
        If sh.HasTextFrame Then
            Dim sName As String
            sName = sh.Name
            If Left(sName, 5) = "Text " Then
                Dim idxNum As Integer
                idxNum = CInt(Mid(sName, 6))
                shapeDict(idxNum) = sh
            End If
        End If
    Next sh

    ' 各データ行を処理
    Dim rowNum As Integer
    For rowNum = 0 To MAX_ROWS - 1

        Dim nameIdx As Integer
        nameIdx = FIRST_NAME_IDX + rowNum * ROW_STEP

        If Not shapeDict.Exists(nameIdx) Then Exit For

        Dim nmShape As Object
        Set nmShape = shapeDict(nameIdx)

        Dim nm As String
        nm = CleanText_M4(nmShape.TextFrame.TextRange.Text)
        If nm = "" Then GoTo NextRowP6

        Dim tot As Long
        tot = 0
        Dim ci As Integer

        For ci = 0 To 3
            Dim colIdx As Integer
            colIdx = nameIdx + 2 + ci * 2

            Dim k As String
            k = nm & "_" & CStr(colCodes(ci))

            Dim cnt As Long
            cnt = 0
            If scc.Exists(k) Then cnt = CLng(scc(k))

            Dim ct As String
            If cnt = 0 Then ct = "－" Else ct = "●"

            If shapeDict.Exists(colIdx) Then
                shapeDict(colIdx).TextFrame.TextRange.Text = ct
            End If

            tot = tot + cnt
        Next ci

        ' 合計（+10）
        Dim totIdx As Integer
        totIdx = nameIdx + 10
        If shapeDict.Exists(totIdx) Then
            shapeDict(totIdx).TextFrame.TextRange.Text = CStr(tot) & "件"
        End If

        ' 優先度（+12）
        Dim prIdx As Integer
        prIdx = nameIdx + 12
        Dim pr As String
        If tot >= 2 Then
            pr = "★高"
        ElseIf tot >= 1 Then
            pr = "中"
        Else
            pr = "低"
        End If
        If shapeDict.Exists(prIdx) Then
            shapeDict(prIdx).TextFrame.TextRange.Text = pr
        End If

NextRowP6:
    Next rowNum

End Sub


'------------------------------------------------------------------------------
' P8更新：店舗個別サマリー（Slide 8）
' テーブルなし。テキストボックスを直接書き換える。
' 「Audi ○○○○」→ 店舗名
' 「1910 発生件数」の次の行のテキストボックスに件数を入れる
' ※P8はテーブルではなくテキストボックスで構成されているため
'   シェイプのテキストを直接書き換える
'------------------------------------------------------------------------------
Private Sub UpdateP8_M4(pptPres As Object, sd As Object, shopName As String)

    Dim sl As Object
    Set sl = pptPres.Slides(8)

    Dim sh As Object
    For Each sh In sl.Shapes

        If sh.HasTextFrame Then
            If sh.TextFrame.HasText Then

                Dim txt As String
                txt = sh.TextFrame.TextRange.Text

                ' 店舗名置換
                If InStr(txt, "Audi ○○○○") > 0 Then
                    sh.TextFrame.TextRange.Text = Replace(txt, "Audi ○○○○", shopName)
                    txt = sh.TextFrame.TextRange.Text
                End If

                ' 1910件数の置換（「　　件」や「　件」パターン）
                ' ラベルシェイプではなく値シェイプを特定するため
                ' 「1910 発生件数」テキストのすぐ後のシェイプを更新するのではなく
                ' テキストに数字プレースホルダーが含まれている場合に書き換える
                Dim cleaned As String
                cleaned = CleanText_M4(txt)

                If cleaned = "1910 発生件数" Then
                    ' このシェイプはラベル。次のシェイプ（値）を書き換えるため
                    ' インデックスを記録しておく
                    ' → 下記の値シェイプ特定方式で処理
                End If

            End If
        End If

    Next sh

    ' P8の値テキストボックスを特定して更新
    ' シェイプ名のText番号順に並べて処理
    Dim shapeArr() As Object
    Dim shapeIdxArr() As Integer
    Dim shCount As Integer
    shCount = 0

    ReDim shapeArr(sl.Shapes.Count)
    ReDim shapeIdxArr(sl.Shapes.Count)

    For Each sh In sl.Shapes
        If sh.HasTextFrame Then
            Dim sn As String
            sn = sh.Name
            If Left(sn, 5) = "Text " Then
                shapeArr(shCount) = sh
                shapeIdxArr(shCount) = CInt(Mid(sn, 6))
                shCount = shCount + 1
            End If
        End If
    Next sh

    ' インデックス順にソート（バブルソート）
    Dim ii As Integer, jj As Integer
    Dim tmpObj As Object
    Dim tmpIdx As Integer
    For ii = 0 To shCount - 2
        For jj = ii + 1 To shCount - 1
            If shapeIdxArr(ii) > shapeIdxArr(jj) Then
                tmpIdx = shapeIdxArr(ii): shapeIdxArr(ii) = shapeIdxArr(jj): shapeIdxArr(jj) = tmpIdx
                Set tmpObj = shapeArr(ii): Set shapeArr(ii) = shapeArr(jj): Set shapeArr(jj) = tmpObj
            End If
        Next jj
    Next ii

    ' ラベルの次のシェイプを値として更新
    Dim major(3) As Long
    major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003

    For ii = 0 To shCount - 1
        Dim lblTxt As String
        lblTxt = CleanText_M4(shapeArr(ii).TextFrame.TextRange.Text)

        Select Case lblTxt
            Case "重大エラーコード"
                ' 次のシェイプに発生コードを記入
                If ii + 1 < shCount Then
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
                    shapeArr(ii + 1).TextFrame.TextRange.Text = codes
                End If

            Case "1910 発生件数"
                If ii + 1 < shCount Then
                    Dim c1 As Long: c1 = 0
                    If sd.Exists("cnt_1910") Then c1 = CLng(sd("cnt_1910"))
                    shapeArr(ii + 1).TextFrame.TextRange.Text = CStr(c1) & "件"
                End If

            Case "発生月"
                If ii + 1 < shCount Then
                    shapeArr(ii + 1).TextFrame.TextRange.Text = GetMonths_M4(sd)
                End If

        End Select
    Next ii

End Sub


'------------------------------------------------------------------------------
' P10更新：発生コード確認（Slide 10）
' 「Audi ○○○○」と「重大エラー発生件数：　　　件」を置換
' ※全角スペース3つ（Chr(12288) x 3）に注意
'------------------------------------------------------------------------------
Private Sub UpdateP10_M4(pptPres As Object, sd As Object, shopName As String)

    Dim sl As Object
    Set sl = pptPres.Slides(10)

    Dim major(3) As Long
    major(0) = 1223: major(1) = 1910: major(2) = 1016: major(3) = 1003

    ' 合計件数
    Dim tot As Long: tot = 0
    Dim cj As Integer
    For cj = 0 To 3
        If sd.Exists("cnt_" & CStr(major(cj))) Then
            tot = tot + CLng(sd("cnt_" & CStr(major(cj))))
        End If
    Next cj

    ' 全角スペース3つのプレースホルダー
    Dim zenkakuSp3 As String
    zenkakuSp3 = ChrW(12288) & ChrW(12288) & ChrW(12288)

    Dim sh As Object
    For Each sh In sl.Shapes
        If sh.HasTextFrame Then
            If sh.TextFrame.HasText Then

                Dim txt As String
                txt = sh.TextFrame.TextRange.Text

                ' 店舗名置換
                If InStr(txt, "Audi " & ChrW(12288) & ChrW(12288) & ChrW(12288) & ChrW(12288)) > 0 Or _
                   InStr(txt, "Audi ○○○○") > 0 Then
                    txt = Replace(txt, "Audi ○○○○", shopName)
                    sh.TextFrame.TextRange.Text = txt
                End If

                txt = sh.TextFrame.TextRange.Text

                ' 件数バー置換（全角スペース3つパターン）
                If InStr(txt, "重大エラー発生件数：" & zenkakuSp3 & "件") > 0 Then
                    sh.TextFrame.TextRange.Text = Replace(txt, _
                        "重大エラー発生件数：" & zenkakuSp3 & "件", _
                        "重大エラー発生件数：" & CStr(tot) & "件")
                End If

            End If
        End If
    Next sh

End Sub
