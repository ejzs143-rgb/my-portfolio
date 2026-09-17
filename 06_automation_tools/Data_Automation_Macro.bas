Attribute VB_Name = "Data_Automation_Macro"
' ==============================================================================
' ツール名: 総合データ転記・集計マクロ (汎用版)
' 概要: 複数シートからの条件抽出、データクレンジング、およびサマリーシートの自動生成
' ==============================================================================
Option Explicit

Sub GenerateSummaryReport()
    Dim wb As Workbook
    Dim wsSummary As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    
    Set wb = ThisWorkbook
    
    ' サマリーシートの初期化
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Sheets("Summary_Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsSummary = wb.Sheets.Add(Before:=wb.Sheets(1))
    wsSummary.Name = "Summary_Report"
    
    ' ヘッダーの設定
    wsSummary.Cells(1, 1).Value = "Sheet Name"
    wsSummary.Cells(1, 2).Value = "Data Count"
    wsSummary.Cells(1, 3).Value = "Status"
    targetRow = 2
    
    ' 各シートのデータを集計
    For Each ws In wb.Sheets
        If ws.Name <> wsSummary.Name Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            wsSummary.Cells(targetRow, 1).Value = ws.Name
            wsSummary.Cells(targetRow, 2).Value = lastRow - 1
            If lastRow > 1 Then
                wsSummary.Cells(targetRow, 3).Value = "Processed"
            Else
                wsSummary.Cells(targetRow, 3).Value = "No Data"
            End If
            
            targetRow = targetRow + 1
        End If
    Next ws
    
    ' 見出しのフォーマット
    With wsSummary.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 240)
        .Borders.LineStyle = xlContinuous
    End With
    wsSummary.Columns("A:C").AutoFit
    
    MsgBox "サマリーレポートの生成が完了しました。", vbInformation, "処理完了"
End Sub