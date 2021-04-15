Sub 修改团员名单()
'
' 修改团员名单 宏
'
If Sheets("团员名单").Visible = False Then
    Sheets("团员名单").Visible = True
    Sheets("团员名单").Select
Else
    Sheets("团员名单").Visible = Fales
        Sheets("函数调用").Select
    End If
'
End Sub