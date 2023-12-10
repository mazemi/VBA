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

    current_sheet_name = ActiveSheet.name
    
    If current_sheet_name <> find_main_data Then
         MsgBox "Please select the main sheet at first.", vbInformation
         Exit Sub
    End If
    
    Call check_uuid
    
    Set ws = sheets(find_main_data)
    ' check if the selected range is in one column
    If selectedRange.columns.count > 1 Then
        MsgBox "Please select form one column.", vbInformation
        Exit Sub
    End If
    
    last_col = ws.Cells(1, columns.count).End(xlToLeft).Column
    
    If selectedRange.Column > last_col Then
        Exit Sub
    End If

    uuid_col_number = column_number("_uuid")
    data_col_number = selectedRange.Column
    question_value = ws.Cells(1, data_col_number).Value

    ' check the seleted range is not in the first row
    If ActiveCell.Row = 1 Then
        MsgBox "Please do not select header row.", vbInformation
        Exit Sub
    End If
    
    If Not auto_checking Then
        ' open issue choices form
        data_checking_form.Show
    End If
    
    If PATTERN_CHECK_ACTION = False Then
        Exit Sub
    End If

    Set main_ws = sheets(find_main_data)
        
    'check if log_book sheet exist
    If worksheet_exists("log_book") <> True Then
        Call create_log_sheet(main_ws.name)
    End If

    Set log_ws = Worksheets("log_book")
    Call clear_filter(log_ws)
    
    main_ws.Activate

    If selectedRange.count > 1 Then
        ' if the selected range have more than one row, we need a loop
        Call selected_rows(uuid_col_number)
        For Each row_item In ROW_ARRAY()
            If row_item > 0 Then
                ' getting new row number
                newRow = log_ws.Cells(rows.count, 1).End(xlUp).Row + 1
    
                log_ws.Cells(newRow, "A").Value = main_ws.Cells(row_item, uuid_col_number)
                log_ws.Cells(newRow, "B").Value = question_value
                log_ws.Cells(newRow, "C").Value = ISSUE_TEXT
                log_ws.Cells(newRow, "D").Value = main_ws.Cells(row_item, data_col_number)
    
                ' add new columns from setting:
                On Error GoTo errHandlerArray:

            End If
        Next row_item
    Else
        ' if the selected range has one row, we do not need loop
        ' getting new row number
        If main_ws.Cells(selectedRange.Row, uuid_col_number) <> "" Then
            newRow = log_ws.Cells(rows.count, 1).End(xlUp).Row + 1
    
            log_ws.Cells(newRow, "A").Value = main_ws.Cells(selectedRange.Row, uuid_col_number)
            log_ws.Cells(newRow, "B").Value = question_value
            log_ws.Cells(newRow, "C").Value = ISSUE_TEXT
            log_ws.Cells(newRow, "D").Value = main_ws.Cells(selectedRange.Row, data_col_number)
    
        End If
    End If

    Application.ScreenUpdating = True
    Exit Sub
    
errHandlerArray:
    MsgBox "There is an issue.                       ", vbCritical
           
    Application.ScreenUpdating = True
    
End Sub

Sub create_log_sheet(sheet_name)

    sheets.Add(after:=sheets(sheet_name)).name = "log_book"
    
    With sheets("log_book")
        ' new columns
        .Range("A1").Value = "uuid"
        .Range("B1").Value = "question.name"
        .Range("C1").Value = "issue"
        .Range("D1").Value = "old.value"
        .Range("E1").Value = "new.value"
        .Range("F1").Value = "changed"
        ' set columns widths:
        .columns("A:A").ColumnWidth = 40
        .columns("B:B").ColumnWidth = 30
        .columns("C:L").ColumnWidth = 15
    End With
    
    'freeze top row:
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    
    ActiveWindow.FreezePanes = True
End Sub

Sub selected_rows(key_col As Long)
    On Error Resume Next
    Dim visibleCells As Range
    Dim cell As Range
    Dim cellCount As Long
    Dim rowIndex As Long
    Dim ws As Worksheet
    Set ws = sheets(find_main_data)
    ' Set the range of visible cells in the selection
    Set visibleCells = Selection.SpecialCells(xlCellTypeVisible)
    
    ' Get the number of visible cells
    cellCount = visibleCells.Cells.count
    
    ReDim ROW_ARRAY(cellCount - 1)

    rowIndex = 0
    For Each cell In visibleCells.Cells
        If ws.Cells(cell.Row, key_col).Value <> "" Then
            ROW_ARRAY(rowIndex) = cell.Row
            rowIndex = rowIndex + 1
        End If
    Next cell
End Sub

Sub find_duplicate()
    On Error Resume Next
    Application.ScreenUpdating = False
    Call check_uuid
    Dim ws As Worksheet
    Set ws = sheets(find_main_data)
    Call clear_filter(ws)
    Dim LastRow As Long

    LastRow = ws.UsedRange.rows(ws.UsedRange.rows.count).Row

    uuid_col_letter = column_letter("_uuid")

    new_col = ws.Cells(1, columns.count).End(xlToLeft).Column + 1
    new_col_letter = Split(Cells(1, new_col).Address, "$")(1)

    ws.Range(new_col_letter & 1).Value = "check_duplicate"
    For m = 2 To LastRow
        If Application.WorksheetFunction.CountIf(ws.Range(uuid_col_letter & "2:" & uuid_col_letter & LastRow), _
                                                 ws.Range(uuid_col_letter & m)) > 1 Then
            ws.Range(new_col_letter & m).Value = "duplicated"
        Else
            ws.Range(new_col_letter & m).Value = "ok"
        End If
    Next m

    Application.ScreenUpdating = True
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
    If selectedRange.columns.count > 1 Then
        MsgBox "Please select one column.      ", vbInformation
        Exit Sub
    End If
    
    data_col_number = selectedRange.Column
    
    Set selectedRange = ws.columns(data_col_number)

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
    ws.columns(data_col_number).AutoFilter Field:=data_col_number, _
                                           Criteria1:="<" & CStr(l_value), Operator:=xlOr, Criteria2:=">" & CStr(h_value)
                                                       
    Exit Sub

Handle_Error:

    Select Case err.Number
    Case 1004
        MsgBox "Quartile can not be calculated!    ", vbExclamation
        err.Clear
    Case Else
        MsgBox "Quartile can not be calculated!    ", vbExclamation
    End Select

End Sub


