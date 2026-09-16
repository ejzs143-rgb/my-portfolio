Attribute VB_Name = "Module_mail"
Public Sub Send_Mail()
'ϐ錾
Dim fln As String
Dim i As Integer
Dim objBook As Excel.Workbook
Dim objSheet As Excel.Worksheet
Dim objMail As MailItem
Dim objRecipient As Recipient

'mF
i = MsgBox("[𑗐M܂B" & Chr(13) & "OK{^NbNA𒆎~邱Ƃ͂ł܂B", 289, "mF")
If i <> 1 Then Exit Sub

'ݒ
fln = "C:\Users\Public\Personal_Folder\Personal\Desktop\v[V\[M\"
Set objBook = GetObject(fln & "Xg.xlsm")
Set objSheet = objBook.Sheets(1)
i = 2

'[쐬AM
With objSheet
Do While .Cells(i, 1) <> ""
Set objMail = Application.CreateItemFromTemplate(fln & "[ev[g.oft")

'ccǉ
Set objRecipient = objMail.Recipients.Add(.Cells(i, 6))
objRecipient.Type = olCC

'ǉ
objMail.Recipients.Add .Cells(i, 4)

'{ύX
objMail.Body = Replace(objMail.Body, "%@lR[h%", .Cells(i, 1))
objMail.Body = Replace(objMail.Body, "%@l%", .Cells(i, 2))
objMail.Body = Replace(objMail.Body, "%T[rXӔCҖ%", .Cells(i, 3))

'Ytt@Cǉ
objMail.Attachments.Add fln & "Ytt@C\" & .Cells(i, 5).Value

'[M
objMail.Send

i = i + 1
Loop
End With

'I
objBook.Close
i = MsgBox("[̑M܂B", 64, "")
End Sub
