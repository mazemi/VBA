Attribute VB_Name = "Time_Checking_Module"

Private Sub csv_import(path As String)
    On Error Resume Next
    
    Dim ws As Worksheet, strFile As String
    Set ws = CURRENT_WORK_BOOK.sheets("temp_sheet")

    With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A1"))
        .Name = "qt"
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh
    End With

    For Each cn In CURRENT_WORK_BOOK.Connections
        Set cn = Nothing
        cn.Delete
    Next cn
    
'    For Each qt In ActiveSheet.QueryTables
    For Each qt In ws.QueryTables
        Set qt = Nothing
        qt.Delete
    Next qt
    
End Sub

Private Sub remove_rows()
    On Error Resume Next

    With CURRENT_WORK_BOOK.sheets("temp_sheet").Cells(1, 1).CurrentRegion
        .AutoFilter 1, "<>*question*"            'Filter for any instance of ""<>*question*" in column A (1)
        .Offset(1).EntireRow.Delete
        .AutoFilter
    End With
    
End Sub

Private Function add_calculation()
    On Error Resume Next
   
    If WorksheetFunction.CountA(CURRENT_WORK_BOOK.sheets("temp_sheet").UsedRange) = 0 And CURRENT_WORK_BOOK.sheets("temp_sheet").Shapes.count = 0 Then
        add_calculation = -1
        Exit Function
    End If

    CURRENT_WORK_BOOK.sheets("temp_sheet").Range("E2").FormulaR1C1 = "=(RC[-1]-RC[-2])/1000"
    
    lRow = CURRENT_WORK_BOOK.sheets("temp_sheet").Cells(rows.count, 1).End(xlUp).row

    CURRENT_WORK_BOOK.sheets("temp_sheet").Range("E2").AutoFill Destination:=CURRENT_WORK_BOOK.sheets("temp_sheet").Range("E2:E" & CStr(lRow))
    
    add_calculation = Application.Sum(CURRENT_WORK_BOOK.sheets("temp_sheet").columns("E:E")) / 60
    
End Function

Private Sub clear_sheet()
    On Error Resume Next

    CURRENT_WORK_BOOK.sheets("temp_sheet").Cells.ClearContents
    
End Sub

Sub check_uuid()
    On Error GoTo errHandler:
    col = WorksheetFunction.Match("_uuid", sheets(ActiveSheet.Name).rows(1), 0)
    Exit Sub

errHandler:
    MsgBox "_uuid column dose not exist.     ", vbInformation
    End
End Sub

Function is_divisible(x As Long, d As Long) As Boolean
    On Error Resume Next
    If (x Mod d) = 0 Then
        is_divisible = True
    Else
        is_divisible = False
    End If
End Function

Sub partial_time_check(start_point As Long, end_point As Long)
    On Error Resume Next
    Dim ws As Worksheet
    counter = 0

    On Error GoTo errHandler:
    Call check_uuid
    progress_form.LabelTitle.Caption = "Time Checking till: " & end_point
    progress_form.Show
    Application.ScreenUpdating = False
    main_sheet = find_main_data
    
    Set ws = CURRENT_WORK_BOOK.sheets(main_sheet)
    
    uuid_col_number = column_number("_uuid")
    start_col_number = column_number("start")
    end_col_number = column_number("end")
    
    If start_col_number = 0 Or end_col_number = 0 Then
        Unload progress_form
        MsgBox "The start or end columns dose not exist.", vbInformation
        End
    End If

    duration_col_number = column_number("duration")
    
    If column_number("duration") = 0 Then
        new_col = ws.Cells(1, columns.count).End(xlToLeft).column + 1
    Else
        new_col = column_number("duration")
    End If
    
    new_col_letter = Split(ws.Cells(1, new_col).Address, "$")(1)
        
    If Not worksheet_exists("temp_sheet") Then
        Call create_sheet(ws.Name, "temp_sheet")
        sheets("temp_sheet").visible = False
    End If
    
    CURRENT_WORK_BOOK.sheets("temp_sheet").Cells.ClearContents
    
    Dim iCell As Range
    
    base_path = CURRENT_WORK_BOOK.path & "\audit\"
    CURRENT_WORK_BOOK.sheets(main_sheet).Select
    
    uuid_col = column_letter("_uuid")
    record_count = end_point - start_point
    
    ' need check the decimal number
    percentage_value = Round(record_count / 100, 1)
    progress_value = record_count / 270
    
    ws.Range(new_col_letter & 1).value = "duration"
    ws.Range(new_col_letter & 1).Offset(, 1).value = "duration_remark"
    
    ws.columns(new_col).ColumnWidth = 10
    ws.columns(new_col + 1).ColumnWidth = 18
    
    For Each iCell In ws.Range(uuid_col & start_point & ":" & uuid_col & end_point).Cells
        If Round((iCell.row - start_point) / percentage_value, 0) < 100 Then
            progress_form.percentage.Caption = CStr(Round((iCell.row - start_point) / percentage_value, 0)) & " %"
        Else
            progress_form.percentage.Caption = "100 %"
        End If
        progress_form.bar.Width = CDec((iCell.row - start_point) / progress_value)
        DoEvents
        Call csv_import(base_path & iCell & "\audit.csv")
        Call remove_rows
        
        Duration = add_calculation()
        
        If Duration = -1 Then
            Duration = DateDiff("s", ws.Cells(iCell.row, start_col_number), ws.Cells(iCell.row, end_col_number)) / 60
            ws.Range(new_col_letter & CStr(iCell.row)).Offset(, 1).value = "no audit file"
        End If
        
        ws.Range(new_col_letter & CStr(iCell.row)).value = Round(Duration, 1)
        Call clear_sheet

    Next iCell
      
    If ws.AutoFilterMode Then Selection.AutoFilter
    If Not ws.AutoFilterMode Then Selection.AutoFilter
    
'    Unload progress_form
    
    Application.DisplayAlerts = False
    sheets("temp_sheet").Delete
    Application.DisplayAlerts = True
     
    Application.ScreenUpdating = True
    
    Exit Sub
    
errHandler:

    If worksheet_exists("temp_sheet") Then
        sheets("temp_sheet").Delete
    End If
    
End Sub

Sub time_check()
    On Error Resume Next
    Set CURRENT_WORK_BOOK = ActiveWorkbook
    Dim s_collection As New Collection
    Dim e_collection As New Collection
    Dim record_count As Long
    
    sheets(find_main_data).Select
    uuid_col = column_letter("_uuid")
    uuid_col_number = column_number("_uuid")
    record_count = sheets(find_main_data).Cells(rows.count, uuid_col_number).End(xlUp).row
    
    If record_count < 1010 Then
        Call partial_time_check(2, record_count)
    Else
        parts = Application.WorksheetFunction.RoundDown(record_count / 1000, 0)
        
        For i = 1 To parts
            If i = 1 Then
                s_collection.Add 2
                e_collection.Add 1000
            Else
                s_collection.Add (i - 1) * 1000 + 1
                e_collection.Add i * 1000
            End If
        Next

        s_collection.Add parts * 1000 + 1
        e_collection.Add record_count

    End If
               
    For i = 1 To e_collection.count
        Call partial_time_check(s_collection(i), e_collection(i))
    Next
    
    Unload progress_form

End Sub


