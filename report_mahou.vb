Function find_in_cells(search_range As Range, pattern As String) As Boolean
    For Each cel In search_range
        If InStr(1, cel.Value, pattern) > 0 Then
            find_in_cells = True
            Exit Function
        End If
    Next cel
    find_in_cells = False
End Function

Function has_cleanup_run() As Boolean
    Dim check_string As String
    check_string = "Bid Due Date Report"
    has_cleanup_run = Not find_in_cells(Range("A1"), check_string)
End Function

Sub first_clean()
    'Remove the first 14 rows
    Rows("1:14").EntireRow.Delete
End Sub

Sub second_clean()
    'Remove everything past the last opportunity row
    Dim bottomRow As Long
    bottomRow = Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.UsedRange.Rows(bottomRow).Select
    Selection.Offset(-4, 0).Select
    Selection.Resize(5, 1).Select
    Selection.EntireRow.Delete
End Sub

Sub big_sort()
    Dim BD, ST, CD, workspace As Range
    Set workspace = ActiveSheet.UsedRange
    Set BD = Intersect(workspace, Range("B2", Range("B2").End(xlDown)))
    Set ST = Intersect(workspace, Range("K2", Range("K2").End(xlDown)))
    Set CD = Intersect(workspace, Range("A2", Range("A2").End(xlDown)))
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add2 _
        Key:=BD, _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveSheet.Sort.SortFields.Add2 _
        Key:=ST, _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        CustomOrder:="Active Prospect,Qualified,Identified,Quoted", _
        DataOption:=xlSortNormal
    ActiveSheet.Sort.SortFields.Add2 _
        Key:=CD, _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange workspace
        .Header = xlYes
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Sub ending()
    bottomRow = Cells(Rows.Count, "A").End(xlUp).Row
    ActiveSheet.UsedRange.Rows(bottomRow).Offset(1, 0).Select
    Selection.Merge
    Selection.HorizontalAlignment = xlCenter
    Selection.Value = "END OF REPORT"
    Selection.Interior.Color = RGB(170, 170, 204)
End Sub

Sub x_completed()
    Dim prev As String
    prev = Range("B1").End(xlDown)
    'Delete rows where EC ID <> "-"
    Range("D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Find(what:="-").Activate
    Selection.ColumnDifferences(ActiveCell).Select
    Selection.EntireRow.Delete
End Sub

Sub sheet_edits()
    Dim lastCol As Long
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    'Remove Quote Date and Amount columns, unwrap sheet, add Notes, resize
    Cells.WrapText = False
    Columns("F:F").ColumnWidth = 70.71
    Columns(lastCol).Select
    ActiveCell.Offset(0, 1).Value = "Notes"
    With Range("L1")
        .Font.Bold = True
        .Interior.Color = RGB(170, 170, 255)
    End With
    Columns("L:L").ColumnWidth = 63.57
    ActiveCell.Offset(0, 2).Value = "Count"
    With Range("M1")
        .Font.Bold = True
        .Interior.Color = RGB(170, 170, 255)
    End With
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        ActiveWindow.FreezePanes = True
    End With
    Columns("I:J").EntireColumn.Delete
End Sub

Sub hl_dupes(col As Range)
    'Highlight Duplicate Values
    With col.FormatConditions.AddUniqueValues
        .DupeUnique = xlDuplicate
        With .Font
            .Bold = True
            .Italic = True
        End With
    End With
End Sub

Sub hl_oppo_dupes()
    'Run hl_dupes on Opportunity Name
    hl_dupes Range("F:F")
End Sub

Sub hl_yday(col As Range)
    'Highlight Yesterdays
    With col.FormatConditions.Add(xlTimePeriod, DateOperator:=xlYesterday)
        .Font.Color = -16383844
        .Interior.Color = 13551615
    End With
End Sub

Sub hl_created_yday()
    'Run hl_yday on Created Date
    hl_yday Range("A:A")
End Sub

Sub gray_out(col As Range)
    'Gray out when LS ID...
    Dim team As String
    team = "=OR(E1=""CJ"",E1=""AT"",E1=""EC"")"
    With Range("D:D").FormatConditions.Add(xlExpression, Formula1:=team)
        With .Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.499984740745262
        End With
    End With
End Sub

Sub gray_out_claimed()
    'Gray out EC ID when...
    gray_out Range("D:D")
End Sub

Sub find_splits(dateCol As String, colTop As Long)
    Dim cur As Range, last As Range, splitrange As Range
    Dim lastsplit As Integer
    lastsplit = 2
    ActiveSheet.UsedRange 'Refresh the used range
    For i = (colTop + 2) To Cells(Rows.Count, dateCol).End(xlUp).Row
        Set cur = Cells(i, dateCol)
        Set last = Cells(i - 1, dateCol)
        If cur.Value <> last.Value Then
            thicken_split_border i
            Set splitrange = Range(Cells(lastsplit, "E").Address, _
                                   Cells(i - 1, "E").Address)
            job_counter splitrange, lastsplit
            lastsplit = i
        End If
    Next
    thicken_split_border i
    Set splitrange = Range(Cells(lastsplit, "E").Address, _
                           Cells(i, "E").Address)
    job_counter splitrange, lastsplit
End Sub

Sub thicken_split_border(ByVal i As Long)
    With ActiveSheet.UsedRange.Rows(i - 1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
End Sub

Sub job_counter(splitrange As Range, lastsplit As Integer)
    Cells(lastsplit, "K").Value = "=COUNTIF(" & splitrange.Address & "," & " ""-"")"
    With Cells(lastsplit, "K")
        .Style = "Calculation"
    End With
End Sub

Sub thicctim()
    find_splits "B", 1
End Sub

Sub main()
    'Sanity check
    If has_cleanup_run() Then
        MsgBox "Setup already run. Cannot run again.", vbCritical
        Exit Sub
    End If
    Dim answer As Integer
    answer = MsgBox("Start Bid Due Date Report Setup?", vbOKCancel)
    If answer = vbOK Then
        'Begin running daily report setup
        first_clean
        second_clean
        big_sort
        sheet_edits
        hl_oppo_dupes
        hl_created_yday
        gray_out_claimed
        x_completed
        thicctim
        ending
        'Confirm completion
        MsgBox "Setup complete."
        'Return to top
        ActiveWindow.ScrollRow = 1
    End If
End Sub
