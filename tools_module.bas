Attribute VB_Name = "tools_module"
Sub import_survey()
'    On Error Resume Next
    Application.DisplayAlerts = False
    tools_path = GetRegistrySetting("ramSetting", "koboToolsReg")

    Application.ScreenUpdating = False
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim main_ws As Worksheet
    Set main_ws = ActiveWorkbook.ActiveSheet

    'check if log_book sheet exist
    If WorksheetExists("survey") <> True Then
        Call create_sheet(main_ws.Name, "survey")
    End If

    Set ImportWorkbook = Workbooks.Open(Filename:=tools_path)
    ImportWorkbook.Worksheets("survey").UsedRange.Copy
    ThisWorkbook.Worksheets("survey").Range("A1").PasteSpecial _
    Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ImportWorkbook.Close

    ' trime all three columns:
    Dim Rng As Range
        For i = 1 To 3
        Set Rng = Columns(i)
        Rng.Value = Application.Trim(Rng)
    Next i

    Call deleteIrrelevantColumns("survey")

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub import_choices()
'    On Error Resume Next
    Application.DisplayAlerts = False
    tools_path = GetRegistrySetting("ramSetting", "koboToolsReg")

    Application.ScreenUpdating = False
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim main_ws As Worksheet
    Set main_ws = ActiveWorkbook.ActiveSheet

    'check if temp_choices sheet exist
    If WorksheetExists("choices") <> True Then
        Call create_sheet(main_ws.Name, "choices")
    End If

    Set ImportWorkbook = Workbooks.Open(Filename:=tools_path)

    ImportWorkbook.Worksheets("choices").UsedRange.Copy 'ThisWorkbook.Worksheets("choices").Range("A1")
    ThisWorkbook.Worksheets("choices").Range("A1").PasteSpecial _
    Paste:=xlPasteValues, SkipBlanks:=False

    ImportWorkbook.Close

    Call deleteIrrelevantColumns("choices")

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub deleteIrrelevantColumns(sheet_name As String)
    Dim keepColumn As Boolean
    Dim currentColumn As Integer
    Dim columnHeading As String
    Dim temp_ws As Worksheet
    Set temp_ws = Worksheets(sheet_name)

    currentColumn = 1

    While currentColumn <= temp_ws.UsedRange.Columns.Count
        columnHeading = temp_ws.UsedRange.Cells(1, currentColumn).Value

        'CHECK WHETHER TO KEEP THE COLUMN
        keepColumn = False
        If columnHeading = "list_name" Then keepColumn = True
        If columnHeading = "type" Then keepColumn = True
        If columnHeading = "name" Then keepColumn = True
        If columnHeading = "label::English" Then keepColumn = True

        If keepColumn Then
            'IF YES THEN SKIP TO THE NEXT COLUMN,
            currentColumn = currentColumn + 1
        Else
            'IF NO DELETE THE COLUMN
            temp_ws.Columns(currentColumn).Delete
        End If

        'LASTLY AN ESCAPE IN CASE THE SHEET HAS NO COLUMNS LEFT
        If (temp_ws.UsedRange.Address = "$A$1") And (temp_ws.Range("$A$1").Text = "") Then Exit Sub
    Wend

End Sub

Function match_type(col_name As String)
    'Declare the variables
    Dim temp_ws As Worksheet
    Dim TableRange As Range
    Dim matchRow As Variant
    
    'Set the variables
    Set temp_ws = Worksheets("survey")
    Set TableRange = temp_ws.Range("A:B")
    
    'Use Application.Match instead of WorksheetFunction.Match to avoid errors
    matchRow = Application.Match(col_name, TableRange.Columns(2), 0)
    
    'Check if the match was successful
    If Not IsError(matchRow) Then
        type_value = TableRange.Cells(matchRow, 1)
        match_type = type_value
    Else
        match_type = ""
    End If
End Function

Sub add_label(question_name As String)
    Dim main_ws As Worksheet
    Set main_ws = ActiveWorkbook.ActiveSheet

    Columns(column_number(question_name) + 1).Insert
    Cells(1, column_number(question_name) + 1).Value = question_name & "_name"

    qeustion_col = column_letter(question_name)
    qeustion_label_col = column_letter(question_name & "_name")

    ' Sheets("survey").Activate
    q_type = match_type(question_name)
    If Left(q_type, 10) = "select_one" Then
        vaArray = Split(q_type, " ")
        key_name = vaArray(UBound(vaArray))

        Dim SourceRange As Range
        Dim DestSheet As Worksheet
        Dim Criteria As String

        last_row = Worksheets("choices").Cells(Rows.Count, 1).End(xlUp).Row

        Set SourceRange = Worksheets("choices").Range("A1:C" & last_row)
        Set DestSheet = Worksheets("temp")

        'Apply the filter on the source range
        SourceRange.AutoFilter Field:=1, Criteria1:=key_name

        'Copy only the visible cells to the destination sheet
        SourceRange.SpecialCells(xlCellTypeVisible).Copy DestSheet.Range("A1")

        'Remove the filter
        SourceRange.AutoFilter

        'Clear the clipboard
        Application.CutCopyMode = False

        ' apply lookup formula:
        main_ws.Select


        last_row = main_ws.Cells(Rows.Count, 1).End(xlUp).Row
        last_row_choices = Sheets("temp").Cells(Rows.Count, 1).End(xlUp).Row

        Range(qeustion_label_col & "2:" & qeustion_label_col & CStr(last_row)).Formula = "=VLOOKUP(" & qeustion_col & "2,'temp'!B$2:C$" & last_row_choices & ",2,False)"
'            Call lookup_label(question_name)
        End If
        
        ' look up formula to values:
        Columns(qeustion_label_col & ":" & qeustion_label_col).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range(qeustion_label_col & "1").Select
        
        Sheets("temp").Cells.Clear

    Worksheets("temp").Visible = False

End Sub

Sub lookup_label(question_name2 As String)
    Dim main_ws As Worksheet
    Set main_ws = ActiveWorkbook.ActiveSheet

    qeustion_col = column_letter(question_name2)
    qeustion_label_col = column_letter(question_name & "_name")

    last_row = main_ws.Cells(Rows.Count, 1).End(xlUp).Row
    last_row_choices = Sheets("temp").Cells(Rows.Count, 1).End(xlUp).Row

    Range(qeustion_label_col & "2:" & qeustion_label_col & CStr(last_row)).Formula = "=VLOOKUP(" & qeustion_col & "2,'temp'!B$2:C$" & last_row_choices & ",2,False)"
End Sub

Sub test_label()
    Dim last_col As Long
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    
    ' check if the selected range is in one column
    If selectedRange.Columns.Count > 1 Or selectedRange.Rows.Count > 1 Or selectedRange.Row <> 1 Then
        MsgBox "Please select form one cell in the first row.", vbInformation
    Exit Sub
    End If

    Call add_label(selectedRange.Value)
End Sub


