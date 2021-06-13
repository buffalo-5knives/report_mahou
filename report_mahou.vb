Dim dict As Object

Function collect_column_coords() As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim col_names_arr As Variant
    col_names_arr = Array( _
                    "Opportunity Name", _
                    "Bid Date", _
                    "EC ID", _
                    "Stage", _
                    "Created Date", _
                    "LS ID")
    Dim tmp_cel As Range
    For Each col_name In col_names_arr
        Set tmp_cel = find_cell_in_cells(get_row(1), CStr(col_name))
        If Not tmp_cel Is Nothing Then
            dict.Add CStr(col_name), tmp_cel
        End If
    Next
    Set collect_column_coords = dict
End Function

Function find_cell_in_cells(search_range As Range, pattern As String) As Range
    For Each cel In search_range
        If InStr(1, CStr(cel.Value), pattern) > 0 Then
            Set find_cell_in_cells = cel
            Exit Function
        End If
    Next cel
    Set find_cell_in_cells = Nothing
End Function

Function find_in_cells(search_range As Range, pattern As String) As Boolean
    If find_cell_in_cells(search_range, pattern) Is Nothing Then
        find_in_cells = False
    Else
        find_in_cells = True
    End If
End Function

Function has_cleanup_run() As Boolean
    Dim check_string As String
    check_string = "Bid Due Date Report"
    has_cleanup_run = Not find_in_cells(Range("A1"), check_string)
End Function

Sub first_clean()
    'Remove the first 14 rows
    If Not has_cleanup_run() Then Rows("1:14").EntireRow.Delete
End Sub

Sub second_clean()
    'Remove everything past the last opportunity row
    Dim bottomCell As Range
    Set bottomCell = Cells(Rows.Count, "A").End(xlUp)
    If find_in_cells(bottomCell, "Copyright") Then
        ActiveSheet.UsedRange.Rows(bottomCell.Row).Select
        Selection.Offset(-4, 0).Select
        Selection.Resize(5, 1).Select
        Selection.EntireRow.Delete
    End If
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
        .header = xlYes
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Sub x_completed()
    'Delete rows where EC ID <> "-"
    If MsgBox("Have Opportunites completed before today already been removed?", vbYesNo) = vbNo Then
        get_col("EC ID").Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Find(what:="-").Activate
        Selection.ColumnDifferences(ActiveCell).Select
        Selection.EntireRow.Delete
    End If
End Sub

Sub sheet_edits()
    Dim lastCol As Long
    lastCol = dict.Item("Stage").Column
    'Remove Quote Date and Amount columns, unwrap sheet, add Notes, resize
    Cells.WrapText = False
    Dim col As Long
    col = dict.Item("Opportunity Name").Column
    With Columns(col)
        .ColumnWidth = 52
    End With
    Columns(lastCol).Select
    With ActiveCell.Offset(0, 1)
        .Value = "Notes"
        .Font.Bold = True
        .Interior.Color = RGB(170, 170, 255)
        With Columns(.Column)
            .ColumnWidth = 63.57
        End With
    End With
    With ActiveCell.Offset(0, 2)
        .Value = "Count"
        .Font.Bold = True
        .Interior.Color = RGB(170, 170, 255)
    End With
    With ActiveWindow
        If Not .FreezePanes Then
            .SplitColumn = 0
            .SplitRow = 1
            .FreezePanes = True
        End If
    End With
End Sub

Sub hl_dupes(col As Range)
    'Highlight Duplicate Values
    With col.FormatConditions
        If .Count < 1 Then
            With .AddUniqueValues
                .DupeUnique = xlDuplicate
                With .Font
                    .Bold = True
                    .Italic = True
                End With
            End With
        End If
    End With
End Sub

Sub hl_oppo_dupes()
    'Run hl_dupes on Opportunity Name
    If dict.Exists("Opportunity Name") Then hl_dupes get_col("Opportunity Name")
End Sub

Sub hl_yday(col As Range)
    'Highlight Yesterdays
    With col.FormatConditions
        If .Count < 1 Then
            With .Add(xlTimePeriod, DateOperator:=xlYesterday)
                .Font.Color = -16383844
                .Interior.Color = 13551615
            End With
        End If
    End With
End Sub

Sub hl_created_yday()
    'Run hl_yday on Created Date
    If dict.Exists("Created Date") Then hl_yday get_col("Created Date")
End Sub

Sub gray_out(col As Range)
    'Gray out when LS ID...
    Dim team As String
    team = "=OR(E1=""CJ"",E1=""AT"",E1=""EC"")"
    With col.FormatConditions
        If .Count < 1 Then
            With .Add(xlExpression, Formula1:=team)
                With .Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.499984740745262
                End With
            End With
        End If
    End With
End Sub

Sub gray_out_claimed()
    'Gray out EC ID when...
    If dict.Exists("EC ID") Then gray_out get_col("EC ID")
End Sub

Sub find_splits(dateCol As Long, colTop As Long)
    Dim cur As Range, last As Range, splitrange As Range
    Dim lastsplit As Integer, lsID As Long
    lsID = get_col("LS ID").Column
    lastsplit = 2
    ActiveSheet.UsedRange 'Refresh the used range
    For i = (colTop + 2) To Cells(Rows.Count, dateCol).End(xlUp).Row
        Set cur = Cells(i, dateCol)
        Set last = Cells(i - 1, dateCol)
        If cur.Value <> last.Value Then
            thicken_split_border i
            Set splitrange = Range(Cells(lastsplit, lsID).Address, _
                                   Cells(i - 1, lsID).Address)
            job_counter splitrange, lastsplit
            lastsplit = i
        End If
    Next
    thicken_split_border i
    Set splitrange = Range(Cells(lastsplit, lsID).Address, _
                           Cells(i, lsID).Address)
    job_counter splitrange, lastsplit
End Sub

Sub thicken_split_border(ByVal i As Long)
    With ActiveSheet.UsedRange.Rows(i - 1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
End Sub

Sub job_counter(splitrange As Range, lastsplit As Integer)
    Dim Cnt As Long
    Cnt = get_col("Stage").Offset(0, 2).Column
    Cells(lastsplit, Cnt).Value = "=COUNTIF(" & splitrange.Address & "," & " ""-"")"
    With Cells(lastsplit, Cnt)
        .Style = "Calculation"
    End With
End Sub

Sub thicctim()
    find_splits get_col("Bid Date").Column, 1
End Sub

Sub main()
    Dim answer As Integer
    answer = MsgBox("Start Bid Due Date Report Setup?", vbOKCancel)
    If answer = vbOK Then
        'Begin running daily report setup
        first_clean
        collect_column_coords
        second_clean
        big_sort
        sheet_edits
        hl_oppo_dupes
        hl_created_yday
        gray_out_claimed
        x_completed
        thicctim
        'Confirm completion
        MsgBox "Setup complete."
        'Return to top
        ActiveWindow.ScrollRow = 1
    End If
End Sub

Function get_row(row_num As Long) As Range
    Set get_row = ActiveSheet.Range( _
        Cells(row_num, 1), _
        Cells(row_num, Columns.Count).End(xlToLeft) _
    )
End Function

Function get_col(header As String) As Range
    With dict.Item(header)
        Set get_col = ActiveSheet.Range( _
            Cells(2, .Column), _
            Cells(Rows.Count, .Column).End(xlUp) _
        )
    End With
End Function