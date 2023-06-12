Attribute VB_Name = "data_checking"
'*****************************************************************************
' faster version of pattern check by avoiding .select .selection .activate
' Date: 9, June 2023
'*****************************************************************************
Global issueText As String
Global patternCheckAction As Boolean
Global rowArray() As Integer
Global globalLogColsArr() As String

Sub pattern_check()
    Dim uuid_col_number As Long
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
    
    last_col = Cells(1, Columns.Count).End(xlToLeft).Column
    
    If selectedRange.Column > last_col Then
        Exit Sub
    End If

    uuid_col_number = column_number("_uuid")
    data_col_number = selectedRange.Column
    question_value = Cells(1, data_col_number).Value

    ' check the seleted range is not in the first row
    If ActiveCell.Row = 1 Then
        MsgBox "Please do not select header row.", vbInformation
        Exit Sub
    End If

    ' open issue choices form
    data_checking_form.Show

    If patternCheckAction = False Then
        Exit Sub
    End If

    Set main_ws = ThisWorkbook.ActiveSheet

    'check if log_book sheet exist
    If WorksheetExists("log_book") <> True Then
        Call create_log_sheet(main_ws.Name)
    End If

    Set log_ws = Sheets("log_book")

    main_ws.Activate

    If selectedRange.Count > 1 Then
        ' if the selected range have more than one row, we need a loop
        Call selected_rows(uuid_col_number)
        For Each row_item In rowArray()
            If row_item > 0 Then
                ' getting new row number
                newRow = log_ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
                log_ws.Cells(newRow, "A").Value = main_ws.Cells(row_item, uuid_col_number)
                log_ws.Cells(newRow, "E").Value = main_ws.Cells(row_item, data_col_number)
                log_ws.Cells(newRow, "B").Value = question_value
                log_ws.Cells(newRow, "C").Value = issueText
    
                ' add new columns from setting:
                If globalLogColsArr(0) <> "0" Then
                    arrLen = UBound(globalLogColsArr) - LBound(globalLogColsArr) + 1
                    For i = 0 To arrLen - 1
                        ' 72 is the character code of "H" and so on:
                        log_ws.Cells(newRow, Chr(i + 72)).Value = main_ws.Cells(row_item, column_number(globalLogColsArr(i)))
                    Next i
                End If
            End If
        Next row_item
    Else
        ' if the selected range has one row, we do not need loop
        ' getting new row number
          If Cells(selectedRange.Row, uuid_col_number) <> "" Then
                newRow = log_ws.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
                log_ws.Cells(newRow, "A").Value = main_ws.Cells(selectedRange.Row, uuid_col_number)
                log_ws.Cells(newRow, "E").Value = main_ws.Cells(selectedRange.Row, data_col_number)
                log_ws.Cells(newRow, "B").Value = question_value
                log_ws.Cells(newRow, "C").Value = issueText
    
                ' add new columns from setting:
                If globalLogColsArr(0) <> "0" Then
                    arrLen = UBound(globalLogColsArr) - LBound(globalLogColsArr) + 1
                    For i = 0 To arrLen - 1
                        ' 72 is the character code of "H" and so on:
                        log_ws.Cells(newRow, Chr(i + 72)).Value = main_ws.Cells(selectedRange.Row, column_number(globalLogColsArr(i)))
                    Next i
                End If
            End If
    End If

    ' format the selected cells with issue
    selectedRange.Interior.Color = RGB(255, 254, 240)

    Application.ScreenUpdating = True
    ' MsgBox Timer - t
End Sub

Sub create_log_sheet(sheet_name)

    Sheets.Add(After:=Sheets(sheet_name)).Name = "log_book"
    ' new columns
    Range("A1").Value = "uuid"
    Range("B1").Value = "question.name"
    Range("C1").Value = "issue"
    Range("D1").Value = "feedback"
    Range("E1").Value = "old.value"
    Range("F1").Value = "new.value"
    Range("G1").Value = "changed"
    ' set columns widths:
    Columns("A:A").ColumnWidth = 40
    Columns("B:B").ColumnWidth = 30
    Columns("C:L").ColumnWidth = 15

    ' add new columns from setting:
    Dim SubStringArr() As String
    Dim SrcString As String

    logColsString = GetRegistrySetting("ramSetting", "koboLogReg")

    ' extra logs column is starting by 8:
    new_log_column = 8

    If logColsString <> "" Then
        globalLogColsArr = Split(logColsString, ",")
        For Each i In globalLogColsArr
            Cells(1, new_log_column) = i
            new_log_column = new_log_column + 1
        Next
    Else
        ReDim globalLogColsArr(1 To 2)
        logColsString = "0,1"
        globalLogColsArr = Split(logColsString, ",")
    End If

    'freeze top row:
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True
End Sub

Sub selected_rows(key_col As Long)
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
        If Cells(cell.Row, key_col).Value <> "" Then
            rowArray(rowIndex) = cell.Row
            rowIndex = rowIndex + 1
        End If
    Next cell
End Sub

Sub find_duplicate()
    Application.ScreenUpdating = False
    Call check_uuid
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Dim lastRow As Long

    lastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row

    uuid_col_letter = column_letter("_uuid")

    new_col = Cells(1, Columns.Count).End(xlToLeft).Column + 1
    new_col_letter = Split(Cells(1, new_col).Address, "$")(1)

    ws.Range(new_col_letter & 1).Value = "check_duplicate"
    For m = 2 To lastRow
        If Application.WorksheetFunction.CountIf(ws.Range(uuid_col_letter & "2:" & uuid_col_letter & lastRow), _
        ws.Range(uuid_col_letter & m)) > 1 Then
            ws.Range(new_col_letter & m).Value = "duplicated"
        Else
            ws.Range(new_col_letter & m).Value = "ok"
        End If
    Next m

    Application.ScreenUpdating = True
End Sub

Sub bb()
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    MsgBox selectedRange.Column
End Sub


