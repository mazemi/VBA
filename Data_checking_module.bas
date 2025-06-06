Attribute VB_Name = "Data_checking_module"

Sub pattern_check(auto_checking As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim uuid_col_number As Long
    Dim t As Single
    Dim ws As Worksheet
    Dim main_ws, log_ws As Worksheet
    Dim cel As Range
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    Dim current_sheet_name As String

    current_sheet_name = ActiveSheet.Name
    
    If current_sheet_name <> find_main_data Then
         MsgBox "Please select the main sheet at first.", vbInformation
         Exit Sub
    End If
    
    Call check_uuid
    
    Set ws = sheets(find_main_data)
    ' check if the selected range is in one column
    If selectedRange.Columns.count > 1 Then
        MsgBox "Please select form one column.", vbInformation
        Exit Sub
    End If
    
    last_col = ws.Cells(1, Columns.count).End(xlToLeft).Column
    
    If selectedRange.Column > last_col Then
        Exit Sub
    End If

    uuid_col_number = column_number("_uuid")
    data_col_number = selectedRange.Column
    question_value = ws.Cells(1, data_col_number).value

    ' check the seleted range is not in the first row
    If ActiveCell.Row = 1 Then
        MsgBox "Please do not select header row.", vbInformation
        Exit Sub
    End If
    
    If Not auto_checking Then
        ' open issue choices form
        data_checking_form.Show
    End If
    
    If Not PATTERN_CHECK_ACTION Then
        Exit Sub
    End If

    Set main_ws = sheets(find_main_data)
        
    'check if log_book sheet exist
    If worksheet_exists("log_book") <> True Then
        Call create_log_sheet(main_ws.Name)
    End If

    Set log_ws = Worksheets("log_book")
    Call clear_filter(log_ws)
    
    main_ws.Activate

    If selectedRange.count > 1 Then
        ' if the selected range have more than one row, we need a loop
         For Each row_item In get_selected_rows(uuid_col_number)
            If row_item > 0 Then
                ' getting new row number
                newRow = log_ws.Cells(Rows.count, 1).End(xlUp).Row + 1
    
                log_ws.Cells(newRow, "A").value = main_ws.Cells(row_item, uuid_col_number)
                log_ws.Cells(newRow, "B").value = question_value
                log_ws.Cells(newRow, "C").value = ISSUE_TEXT
                log_ws.Cells(newRow, "D").value = main_ws.Cells(row_item, data_col_number)
    
                ' add new columns from setting:
                On Error GoTo ErrorHandlerArray:

            End If
        Next row_item
    Else
        ' if the selected range has one row, we do not need loop
        ' getting new row number
        If main_ws.Cells(selectedRange.Row, uuid_col_number) <> "" Then
            newRow = log_ws.Cells(Rows.count, 1).End(xlUp).Row + 1
    
            log_ws.Cells(newRow, "A").value = main_ws.Cells(selectedRange.Row, uuid_col_number)
            log_ws.Cells(newRow, "B").value = question_value
            log_ws.Cells(newRow, "C").value = ISSUE_TEXT
            log_ws.Cells(newRow, "D").value = main_ws.Cells(selectedRange.Row, data_col_number)
    
        End If
    End If

    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandlerArray:
    MsgBox "There is an issue.                       ", vbCritical
           
    Application.ScreenUpdating = True
    
End Sub

Sub create_log_sheet(SHEET_NAME)
    Dim validationList As String
    sheets.Add(after:=sheets(SHEET_NAME)).Name = "log_book"
    
    With sheets("log_book")
        ' new columns
        .Range("A1").value = "uuid"
        .Range("B1").value = "question.name"
        .Range("C1").value = "issue"
        .Range("D1").value = "old.value"
        .Range("E1").value = "new.value"
        .Range("F1").value = "changed"
        .Columns("A:A").ColumnWidth = 40
        .Columns("B:B").ColumnWidth = 30
        .Columns("C:L").ColumnWidth = 15
    End With
    
    validationList = "yes,no"

    sheets("log_book").Range("F2:F" & Rows.count).Validation.Delete
    
    With sheets("log_book").Range("F2:F" & Rows.count).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=validationList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    
    ActiveWindow.FreezePanes = True
End Sub

Function get_selected_rows(key_col As Long) As Variant
    On Error Resume Next
    Dim visibleCells As Range
    Dim cell As Range
    Dim cellCount As Long
    Dim rowIndex As Long
    Dim ws As Worksheet
    Dim row_arr() As Long
    
    Set ws = sheets(find_main_data)
    
    Set visibleCells = Selection.SpecialCells(xlCellTypeVisible)
    
    If Not visibleCells Is Nothing Then
        cellCount = visibleCells.Cells.count
        
        ReDim row_arr(cellCount - 1)
        
        rowIndex = 0
        For Each cell In visibleCells.Cells
            If ws.Cells(cell.Row, key_col).value <> "" Then
                row_arr(rowIndex) = cell.Row
                rowIndex = rowIndex + 1
            End If
        Next cell
    End If
    
    get_selected_rows = row_arr
End Function

Sub find_duplicate()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim lastRow As Long
   
    Set ws = sheets(find_main_data)
    
    Call check_uuid
    Call clear_filter(ws)
     
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.count).Row

    uuid_col_letter = column_letter("_uuid")

    new_col = ws.Cells(1, Columns.count).End(xlToLeft).Column + 1
    new_col_letter = Split(Cells(1, new_col).Address, "$")(1)
    
    If ws.Cells(1, new_col) <> vbNullString Then
        Columns(new_col).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
    
    ws.Range(new_col_letter & 1).value = "check_duplicate"
    For m = 2 To lastRow
        If Application.WorksheetFunction.CountIf(ws.Range(uuid_col_letter & "2:" & uuid_col_letter & lastRow), _
                                                 ws.Range(uuid_col_letter & m)) > 1 Then
            ws.Range(new_col_letter & m).value = "duplicated"
        Else
            ws.Range(new_col_letter & m).value = "ok"
        End If
    Next m
    
    Application.ScreenUpdating = True
    
    If new_col > 10 Then
        ActiveWindow.ScrollColumn = new_col - 7
    ElseIf new_col > 3 Then
        ActiveWindow.ScrollColumn = new_col - 3
    Else
        ActiveWindow.ScrollColumn = new_col
    End If
    
    ActiveWindow.ScrollRow = 1
    
    ws.Range(new_col_letter & 1).ColumnWidth = 18
    ws.Range(new_col_letter & 1).Activate

    Call remove_auto_filter
    Call add_auto_filter
    
End Sub

Sub calulate_quartiles()

    On Error GoTo Handle_Error
    Dim is_date As Boolean
    Dim ws As Worksheet
    Dim last_col As Long
    Dim selectedRange As Range
    
    Set ws = ActiveWorkbook.ActiveSheet
    
    Set selectedRange = Application.Selection
    
    ' check if the selected range is in one column
    If selectedRange.Columns.count > 1 Then
        MsgBox "Please select one column.      ", vbInformation
        Exit Sub
    End If
    
    data_col_number = selectedRange.Column
    
    Set selectedRange = ws.Columns(data_col_number)

    ' Q1 and Q3 calculation:
    Dim q1 As Variant
    q1 = Application.WorksheetFunction.Quartile(selectedRange, 1)
    q3 = Application.WorksheetFunction.Quartile(selectedRange, 3)
    
    l_value = q1 - 1.5 * (q3 - q1)
    h_value = q3 + 1.5 * (q3 - q1)
        
    MsgBox "Minimum IQR value: " & l_value & "   Maximum IQR value: " & h_value
    
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilter.Range.AutoFilter
        ActiveSheet.AutoFilterMode = False
    End If
    
    ActiveSheet.AutoFilterMode = False
    
    ' Selection.AutoFilter
    ws.Cells(1, data_col_number).AutoFilter
    ws.Columns(data_col_number).AutoFilter Field:=data_col_number, _
                                           Criteria1:="<" & CStr(l_value), Operator:=xlOr, Criteria2:=">" & CStr(h_value)
                                                       
    Exit Sub

Handle_Error:

    Select Case Err.Number
    Case 1004
        MsgBox "Quartile can not be calculated!    ", vbExclamation
        Err.Clear
    Case Else
        MsgBox "Quartile can not be calculated!    ", vbExclamation
    End Select

End Sub

Sub no_value_col()
    On Error Resume Next
    Dim dt_ws As Worksheet
    Dim i As Long
    Dim last_col As Long
    Dim colle As New Collection
    Dim str As String
    Dim rng As Range
    
    Set dt_ws = sheets(find_main_data)
    
    last_col = dt_ws.Cells(1, Columns.count).End(xlToLeft).Column
    
    For i = 1 To last_col
        If WorksheetFunction.CountA(dt_ws.Columns(i)) = 1 Or WorksheetFunction.CountA(dt_ws.Columns(i)) = 0 Then
            colle.Add i
        End If
    Next

    If colle.count > 0 Then
    
        If Not worksheet_exists("temp_sheet") Then
            Call create_sheet(find_main_data, "temp_sheet")
            sheets("temp_sheet").Visible = False
        End If
        
        With sheets("temp_sheet")
            .Cells.Clear
            .Range("A1") = "Column"
            .Range("B1") = "Value"
            For j = 1 To colle.count
                .Range("A" & j + 1) = number_to_letter(colle.item(j), dt_ws)
                .Range("B" & j + 1) = dt_ws.Cells(1, colle.item(j))
            Next j
        End With
        
        With empty_col_form.ListBoxEmptyCols
            .ColumnHeads = True
            .columnCount = 2
            .columnWidths = "60;140"
        End With
        
        Set rng = sheets("temp_sheet").Range("A1").CurrentRegion
        empty_col_form.ListBoxEmptyCols.RowSource = _
            rng.Parent.Name & "!" & rng.Resize(rng.Rows.count - 1).Offset(1).Address
        empty_col_form.Show
    
    Else
        MsgBox "No empty column.   ", vbInformation
    End If
    Debug.Print empty_col_form.ListBoxEmptyCols.Height
End Sub

