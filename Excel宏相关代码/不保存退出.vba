Sub 不保存退出()
'
' 不保存退出 宏

If Workbooks.Count > 1 Then
     ActiveWorkbook.Close SaveChanges:=False
End If
If Workbooks.Count = 1 Then
Application.Quit
ActiveWorkbook.Close SaveChanges:=False
End If

'
End Sub