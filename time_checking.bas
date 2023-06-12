Attribute VB_Name = "time_checking"

Private Sub csv_import(path As String)
    On Error Resume Next
    Dim ws As Worksheet, strFile As String
    Set ws = ActiveWorkbook.Sheets("temp_sheet") 'set to current worksheet name
    
    With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A1"))
         .TextFileParseType = xlDelimited
         .TextFileCommaDelimiter = True
         .Refresh
    End With
End Sub

Private Sub remove_rows()
    On Error Resume Next
    Sheets("temp_sheet").Select
    With Cells(1, 1).CurrentRegion
        .AutoFilter 1, "<>*question*"   'Filter for any instance of ""<>*question*" in column A (1)
        .Offset(1).EntireRow.Delete
        .AutoFilter
    End With
End Sub

Private Function add_calculation()
    On Error Resume Next
    Sheets("temp_sheet").Select
    If WorksheetFunction.CountA(ActiveSheet.UsedRange) = 0 And ActiveSheet.Shapes.Count = 0 Then
        add_calculation = -1
        Exit Function
    End If

    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]-RC[-2])/1000"
    
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & CStr(lRow))
    
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-1])/60"
    Range("F2").Select
    add_calculation = Range("F2")
    
End Function

Private Sub clear_sheet()
    On Error Resume Next
    Sheets("temp_sheet").Select
    Cells.Select
    Selection.QueryTable.Delete
    Selection.ClearContents
End Sub

Sub time_check()
    Call check_uuid
    progress_form.LabelTitle.Caption = "Time Checking"
    progress_form.Show
    Application.ScreenUpdating = False
    main_sheet = ActiveSheet.Name
    
    uuid_col_number = column_number("_uuid")
    start_col_number = column_number("start")
    end_col_number = column_number("end")
    
    new_col = Cells(1, Columns.Count).End(xlToLeft).Column + 1
    new_col_letter = Split(Cells(1, new_col).Address, "$")(1)
    
    Sheets.Add.Name = "temp_sheet"
    Dim iCell As Range
    
    base_path = ThisWorkbook.path & "\audit\"
    Sheets(main_sheet).Select
    
    uuid_col = column_letter("_uuid")
    record_count = Cells(Rows.Count, uuid_col_number).End(xlUp).Row
    
    percentage_value = Round(record_count / 100, 0)
    progress_value = record_count / 270
    
    For Each iCell In Range(uuid_col & "2:" & uuid_col & CStr(record_count)).Cells
        progress_form.percentage.Caption = CStr(Round(iCell.Row / percentage_value, 0)) & " %"
        progress_form.bar.Width = CDec(iCell.Row / progress_value)
        DoEvents
        Call csv_import(base_path & iCell & "\audit.csv")
        Call remove_rows
        
        Duration = add_calculation()
        Sheets(main_sheet).Select
        If Duration = -1 Then
            Duration = DateDiff("s", Cells(iCell.Row, start_col_number), Cells(iCell.Row, end_col_number)) / 60
            Range(new_col_letter & CStr(iCell.Row)).Offset(, 1).Value = "no audit file"
        End If
        
        Range(new_col_letter & CStr(iCell.Row)).Value = Round(Duration, 1)
        Call clear_sheet
        Sheets(main_sheet).Select
    Next iCell
    
    Range(new_col_letter & 1).Offset(, 1).Value = "duration_remark"
    Range(new_col_letter & 1).Offset(, 1).Select
    
    Range(new_col_letter & 1).Value = "duration"
    Range(new_col_letter & 1).Select
    
    If ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    If Not ActiveSheet.AutoFilterMode Then Selection.AutoFilter
    
    Unload progress_form
    Application.DisplayAlerts = False
    Sheets("temp_sheet").Delete
    Application.DisplayAlerts = True
     
    Application.ScreenUpdating = True
End Sub

Sub check_uuid()
    On Error GoTo errHandler:
    col = WorksheetFunction.Match("_uuid", Sheets(ActiveSheet.Name).Rows(1), 0)
Exit Sub

errHandler:
    MsgBox "_uuid column dose not exist.     ", vbInformation
    End
End Sub

Public Sub last_col_number()
    last_col = Cells(1, Columns.Count).End(xlToLeft).Column
    MsgBox last_col
End Sub

