Attribute VB_Name = "Data_checking_module"


Sub find_duplicate_log()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    'check if log_book sheet exist
    If Not worksheet_exists("log_book") Then
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim last_col As Long
    Dim last_row As Long
    Dim r_col As Long
    Dim k_col As Long
    Dim has_duplicate As Boolean
    has_duplicate = False
    
    Set ws = sheets("log_book")
    
    ws.Activate
    
    Call clear_active_filter
    
    r_col = gen_column_number("row", ws.Name)
    If r_col > 0 Then
        ws.columns(r_col).Delete
    End If
     
    k_col = gen_column_number("key", ws.Name)
    If k_col > 0 Then
        ws.columns(k_col).Delete
    End If
    
    last_col = ws.Cells(1, columns.count).End(xlToLeft).column
    last_row = ws.Cells(rows.count, 1).End(xlUp).row
     
    If last_row < 3 Then
        extra_logs_form.LabelMessage.Caption = "No Duplicate :)"
        Exit Sub
    End If
     
'    Debug.Print last_row
    ws.Cells(1, last_col + 1) = "row"
    ws.Cells(1, last_col + 2) = "key"
    
    ws.Cells(2, last_col + 1).Formula = "1"
    ws.Cells(2, last_col + 1).AutoFill Destination:=ws.Range(ws.Cells(2, last_col + 1), _
                        ws.Cells(last_row, last_col + 1)), Type:=xlFillSeries
    
    ws.Range(ws.Cells(2, last_col + 2), ws.Cells(last_row, last_col + 2)).Formula = "=A2 & B2"
    
    key_col_letter = gen_column_letter("key", ws.Name)
    
    For M = 2 To last_row
        If Application.WorksheetFunction.CountIf(ws.Range(key_col_letter & "2:" & key_col_letter & last_row), _
                                                 ws.Range(key_col_letter & M)) > 1 Then
            has_duplicate = True
            Exit For
        End If
    Next M
    
    If has_duplicate Then
        ' find duplicated
        ws.columns(last_col + 2).FormatConditions.AddUniqueValues
        ws.columns(last_col + 2).FormatConditions(ws.columns(last_col + 2).FormatConditions.count).SetFirstPriority
        ws.columns(last_col + 2).FormatConditions(1).DupeUnique = xlDuplicate
        With ws.columns(last_col + 2).FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 13551615
            .TintAndShade = 0
        End With
        ws.columns(last_col + 2).FormatConditions(1).StopIfTrue = False
        
        ws.columns(last_col + 2).ColumnWidth = 50
    Else
        ws.columns(last_col + 2).Delete
        ws.columns(last_col + 1).Delete
        extra_logs_form.LabelMessage.Caption = "No Duplicate :)"
    End If
       
    Application.ScreenUpdating = True
End Sub

Sub show_issue()

    If ActiveSheet.Name = "log_book" Then
    
        Dim main_ws As Worksheet
        Dim log_ws As Worksheet
        Dim found_uuid As Range
        Dim question As String
        Dim question_col As Long
        Dim uuid_col As String
        
        Set main_ws = sheets(find_main_data)
        
        Call clear_filter(main_ws)
        
        Set log_ws = sheets("log_book")
        
        question = log_ws.Cells(ActiveCell.row, 2)
        
        uuid_col = data_column_letter("_uuid")
        
        question_col = gen_column_number(question, main_ws.Name)
        
        If log_ws.Cells(ActiveCell.row, 1) <> vbNullString And uuid_col <> "" And question_col <> 0 Then
        
            Set found_uuid = main_ws.Range(uuid_col & ":" & uuid_col).Find(What:=log_ws.Cells(ActiveCell.row, 1))
            
            If Not found_uuid Is Nothing Then
                main_ws.Activate
                main_ws.Cells(found_uuid.row, question_col).Select
            End If
            
        End If
        
    End If

End Sub

Sub pattern_check(auto_check As Boolean)
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
    If selectedRange.columns.count > 1 Then
        MsgBox "Please select form one column.", vbInformation
        Exit Sub
    End If
    
    last_col = ws.Cells(1, columns.count).End(xlToLeft).column
    
    If selectedRange.column > last_col Then
        Exit Sub
    End If

    uuid_col_number = column_number("_uuid")
    data_col_number = selectedRange.column
    question_value = ws.Cells(1, data_col_number).value

    ' check the seleted range is not in the first row
    If ActiveCell.row = 1 Then
        MsgBox "Please do not select header row.", vbInformation
        Exit Sub
    End If
    
    If Not auto_check Then
        ' open issue choices form
        data_checking_form.Show
    End If
    
    If PATTERN_CHECK_ACTION = False Then
        Exit Sub
    End If

    Set main_ws = ActiveWorkbook.ActiveSheet
        
    'check if log_book sheet exist
    If worksheet_exists("log_book") <> True Then
        Call create_log_sheet(main_ws.Name)
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
                newRow = log_ws.Cells(rows.count, 1).End(xlUp).row + 1
    
                log_ws.Cells(newRow, "A").value = main_ws.Cells(row_item, uuid_col_number)
                log_ws.Cells(newRow, "E").value = main_ws.Cells(row_item, data_col_number)
                log_ws.Cells(newRow, "B").value = question_value
                log_ws.Cells(newRow, "C").value = ISSUE_TEXT
    
                ' add new columns from setting:
                On Error GoTo errHandlerArray:

            End If
        Next row_item
    Else
        ' if the selected range has one row, we do not need loop
        ' getting new row number
        If main_ws.Cells(selectedRange.row, uuid_col_number) <> "" Then
            newRow = log_ws.Cells(rows.count, 1).End(xlUp).row + 1
    
            log_ws.Cells(newRow, "A").value = main_ws.Cells(selectedRange.row, uuid_col_number)
            log_ws.Cells(newRow, "E").value = main_ws.Cells(selectedRange.row, data_col_number)
            log_ws.Cells(newRow, "B").value = question_value
            log_ws.Cells(newRow, "C").value = ISSUE_TEXT
    
        End If
    End If

    Application.ScreenUpdating = True
    Exit Sub
    
errHandlerArray:
    MsgBox "There is an issue.                       ", vbCritical
           
    Application.ScreenUpdating = True
    
End Sub

Sub create_log_sheet(sheet_name)
    On Error Resume Next
    sheets.Add(After:=sheets(sheet_name)).Name = "log_book"
    
    With sheets("log_book")
        ' new columns
        .Range("A1").value = "uuid"
        .Range("B1").value = "question.name"
        .Range("C1").value = "issue"
        .Range("D1").value = "feedback"
        .Range("E1").value = "old.value"
        .Range("F1").value = "new.value"
        .Range("G1").value = "changed"
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
        If ws.Cells(cell.row, key_col).value <> "" Then
            ROW_ARRAY(rowIndex) = cell.row
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

    LastRow = ws.UsedRange.rows(ws.UsedRange.rows.count).row

    uuid_col_letter = column_letter("_uuid")

    new_col = ws.Cells(1, columns.count).End(xlToLeft).column + 1
    new_col_letter = Split(Cells(1, new_col).Address, "$")(1)

    ws.Range(new_col_letter & 1).value = "check_duplicate"
    For M = 2 To LastRow
        If Application.WorksheetFunction.CountIf(ws.Range(uuid_col_letter & "2:" & uuid_col_letter & LastRow), _
                                                 ws.Range(uuid_col_letter & M)) > 1 Then
            ws.Range(new_col_letter & M).value = "duplicated"
        Else
            ws.Range(new_col_letter & M).value = "ok"
        End If
    Next M

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
    
    data_col_number = selectedRange.column
    
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


