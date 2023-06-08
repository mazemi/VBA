Attribute VB_Name = "data_checking"
'*****************************************************************************
' faster version of pattern check by avoiding .select .selection .activate
' Date: 9, June 2023
'*****************************************************************************
Global issueText As String
Global patternCheckAction As Boolean
Global rowArray() As Integer

Sub pattern_check()
    Dim t As Single
    t = Timer
    Application.ScreenUpdating = False
    Call check_uuid
    Dim main_ws, log_ws As Worksheet
    Dim cel As Range
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    ' check if the selected range is in one column
    If selectedRange.Columns.Count > 1 Then
        MsgBox "Please select form one column.", vbInformation
    Exit Sub
    End If
    
    uuid_col_number = column_number("_uuid")
    data_col_number = selectedRange.Column
    question_value = Cells(1, data_col_number).Value
    
    ' open issue choices form
    data_checking_form.Show
    
    If patternCheckAction = False Then
        Exit Sub
    End If
    ' check the seleted range is not in the first row
    If ActiveCell.Row = 1 Then
        Exit Sub
    End If
    
    Set main_ws = ThisWorkbook.ActiveSheet
    'check if log_book sheet exist
    If WorksheetExists("log_book") <> True Then
        Call create_sheet(main_ws.Name)
    End If
    
    Set log_ws = Sheets("log_book")

    main_ws.Activate
    
    If selectedRange.Count > 1 Then
        ' if the selected range have more than one row, we need a loop
        Call selected_rows
        For Each row_item In rowArray()
            ' getting new row number
            newRow = log_ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
            
            log_ws.Cells(newRow, "A").Value = main_ws.Cells(row_item, uuid_col_number).Value
            log_ws.Cells(newRow, "E").Value = main_ws.Cells(row_item, data_col_number)
            log_ws.Cells(newRow, "B").Value = question_value
            log_ws.Cells(newRow, "C").Value = issueText
        Next row_item
    Else
            ' if the selected range has one row, we do not need loop
            ' getting new row number
            newRow = log_ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
            
            log_ws.Cells(newRow, "A").Value = main_ws.Cells(selectedRange.Row, uuid_col_number)
            log_ws.Cells(newRow, "E").Value = main_ws.Cells(selectedRange.Row, data_col_number)
            log_ws.Cells(newRow, "B").Value = question_value
            log_ws.Cells(newRow, "C").Value = issueText
      
    End If
    ' format the selected cells with issue
    selectedRange.Interior.Color = RGB(255, 254, 240)

    Application.ScreenUpdating = True
    MsgBox Timer - t
End Sub

Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function

Sub create_sheet(sheet_name)
    Sheets.Add(After:=Sheets(sheet_name)).Name = "log_book"
    ' new columns
    Range("A1").Value = "uuid"
    Range("B1").Value = "question.name"
    Range("C1").Value = "issue"
    Range("D1").Value = "feedback"
    Range("E1").Value = "old.value"
    Range("F1").Value = "new.value"
    Range("G1").Value = "changed"
    ' set columns widths
    Columns("A:A").ColumnWidth = 40
    Columns("B:B").ColumnWidth = 30
    Columns("C:G").ColumnWidth = 15
    'freeze top row
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub

Sub selected_rows()
    ' Declare variables with proper data types and names
    Dim visibleCells As Range
    Dim cell As Range
    Dim cellCount As Long
    Dim rowIndex As Long
         
    ' Set the range of visible cells in the selection
    Set visibleCells = Selection.SpecialCells(xlCellTypeVisible)
    ' Get the number of visible cells
    cellCount = visibleCells.Cells.Count
    ' Resize the array to hold the row numbers
    ReDim rowArray(cellCount - 1)
    
    ' Loop through the visible cells and store the row numbers in the array
    rowIndex = 0
    For Each cell In visibleCells.Cells
        rowArray(rowIndex) = cell.Row
        rowIndex = rowIndex + 1
    Next cell
End Sub





