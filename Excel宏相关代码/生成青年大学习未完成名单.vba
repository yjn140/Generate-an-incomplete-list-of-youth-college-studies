Sub 生成青年大学习未完成名单()
'
' 生成青年大学习未完成名单 宏
'

'
    Dim time As String
    time = CStr(Now())
    Dim filePath As String
    filePath = [函数调用!B5] & "导出文件.csv"
  
    If IsFileExists(filePath) = True Then
     Sheets("函数调用").Select
    Else
     MsgBox "未找到导出文件.csv，请对csv文件进行更名！"
     If Workbooks.Count > 1 Then
      ActiveWorkbook.Close SaveChanges:=False
     End If
     If Workbooks.Count = 1 Then
      Application.Quit
      ActiveWorkbook.Close SaveChanges:=False
     End If
    End If

    If Sheets("团员名单").Visible = Ture Then
    Sheets("团员名单").Visible = False
    End If
    If Sheets("导出文件").Visible = Ture Then
    Sheets("导出文件").Visible = False
    End If
    Sheets("团员名单").Visible = True
    Sheets("导出文件").Visible = True
    Sheets("导出文件").Select
    ActiveWorkbook.RefreshAll
    If Application.Wait(Now + TimeValue("0:00:10")) Then
    Sheets("函数调用").Select
    End If
    Application.CommandBars("Queries and Connections").Visible = False
    Sheets("团员名单").Select
    Columns("B:B").Select
    Selection.Replace What:="1", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="2", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="3", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="4", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="5", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="6", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="7", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="8", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="9", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="0", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("A:B").Select
    Selection.Copy
    Sheets("未完成名单").Select
    Columns("A:A").Select
    ActiveSheet.Paste
    Sheets("导出文件").Select
    Columns("F:F").Select
    Selection.Replace What:="1", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="2", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="3", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="4", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="5", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="6", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="7", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="8", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="9", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="0", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="-", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="/", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="\", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="(", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=")", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="（", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="）", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("未完成名单").Select
    Columns("D:D").Select
    ActiveSheet.Paste
    Range("D1:D1048574,B:B").Select
    Range("B1").Activate
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlUnique
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Application.CutCopyMode = False
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "班级"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "姓名"
    Range("A1").Select
    ActiveWindow.SmallScroll Down:=-69
    Rows("1:1").Select
    Selection.FormatConditions.Delete
    Range("B1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("未完成名单").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("未完成名单").AutoFilter.Sort.SortFields.Add(Range( _
        "B1:B1610"), xlSortOnCellColor, xlAscending, , xlSortNormal).SortOnValue.Color _
        = RGB(255, 199, 206)
    With ActiveWorkbook.Worksheets("未完成名单").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$A$1:$B$1610").AutoFilter Field:=2, Criteria1:=RGB(255, _
        199, 206), Operator:=xlFilterCellColor
    Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-45
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "临时生成页面"
    Sheets("临时生成页面").Select
    Columns("A:A").Select
    ActiveSheet.Paste
    Cells.Select
    Selection.FormatConditions.Delete
    Selection.ColumnWidth = 15
    Selection.RowHeight = 20
    Rows("1:1").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Sheets("未完成名单").Select
    Cells.Select
    Selection.Delete Shift:=xlUp
    Cells.FormatConditions.Delete
    Cells.Select
    Selection.ClearContents
    Sheets("临时生成页面").Select
    Columns("A:B").Select
    Range("A2").Activate
    Selection.Copy
    Sheets("未完成名单").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("临时生成页面").Select
    Application.CutCopyMode = False
    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
    Sheets("未完成名单").Select
    Cells.Select
    Range("A2").Activate
    Selection.ColumnWidth = 15
    Selection.RowHeight = 20
    Sheets("团员名单").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("导出文件").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("未完成名单").Select
    Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    With Selection.Font
        .Name = "等线"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    ActiveCell.FormulaR1C1 = "=导出文件!R[3]C&""未完成名单"""
End Sub

Private Function IsFileExists(ByVal strFileName As String) As Boolean
  If Len(Dir(strFileName)) <> 0 Then
    IsFileExists = True
  Else
    IsFileExists = False
  End If
End Function




