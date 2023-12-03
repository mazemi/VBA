Attribute VB_Name = "tools_module"
Option Explicit

Sub import_survey(tools_path As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim ImportWorkbook As Workbook
    Dim i As Long
    Dim label_col As Long
    
    Set wb = ThisWorkbook
    
    ' clear survey sheet
    wb.sheets("xsurvey").Cells.Clear

    Set ImportWorkbook = Workbooks.Open(Filename:=tools_path)
    
    If (ImportWorkbook.Worksheets("survey").AutoFilterMode And ImportWorkbook.Worksheets("survey").FilterMode) Or _
       ImportWorkbook.Worksheets("survey").FilterMode Then
        ImportWorkbook.Worksheets("survey").ShowAllData
    End If
    
    ImportWorkbook.Worksheets("survey").Cells.NumberFormat = "@"
    
    ImportWorkbook.Worksheets("survey").UsedRange.Copy
    wb.Worksheets("xsurvey").Range("A1").PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ImportWorkbook.Close
    
    Call convert_to_lower(wb.sheets("xsurvey"))
    Call delete_irrelevant_columns("xsurvey")
    
    ' trime all three columns:
    Dim rng As Range
    
    For i = 1 To 3
        Set rng = columns(i)
        rng.Value = Application.Trim(rng)
    Next i

    label_col = this_gen_column_number("label::english", "xsurvey")
     
    If label_col > 0 Then
        wb.sheets("xsurvey").Cells(1, label_col).Value = "label"
    End If
    
   
    If wb.sheets("xsurvey").Range("A1") <> "type" Or wb.sheets("xsurvey").Range("B1") <> "name" Or _
        wb.sheets("xsurvey").Range("C1") <> "label" Then
        Call xsurvey_column_order
        
    End If
    
'    wb.sheets("xsurvey").columns("B").NumberFormat = "@"
    
    If wb.sheets("xsurvey").Range("A1") <> "type" Or wb.sheets("xsurvey").Range("B1") <> "name" Or _
        wb.sheets("xsurvey").Range("C1") <> "label" Then
        MsgBox "Please check the survey sheet of the tool.", vbInformation
        End
    End If
    
End Sub

Sub xsurvey_column_order()
    Dim search As Range
    Dim cnt As Integer
    Dim colOrdr As Variant
    Dim indx As Integer
    
    colOrdr = Array("type", "name", "label")
    
    cnt = 1
    
    For indx = LBound(colOrdr) To UBound(colOrdr)
        Set search = rows("1:1").Find(colOrdr(indx), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not search Is Nothing Then
            If search.Column <> cnt Then
                search.EntireColumn.Cut
                ThisWorkbook.sheets("xsurvey").columns(cnt).Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
        cnt = cnt + 1
        End If
    Next indx
End Sub

Sub import_choices(tools_path As String)

    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim ImportWorkbook As Workbook
    Dim i As Long
    Dim label_col As Long

    ThisWorkbook.sheets("xchoices").Cells.Clear
        
    Set ImportWorkbook = Workbooks.Open(Filename:=tools_path)
    
    If (ImportWorkbook.Worksheets("choices").AutoFilterMode And ImportWorkbook.Worksheets("choices").FilterMode) Or _
       ImportWorkbook.Worksheets("choices").FilterMode Then
        ImportWorkbook.Worksheets("choices").ShowAllData
    End If
    
    ImportWorkbook.Worksheets("choices").Cells.NumberFormat = "@"
    
    ImportWorkbook.Worksheets("choices").UsedRange.Copy
    wb.Worksheets("xchoices").Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, SkipBlanks:=False
  
    ' Paste:=xlPasteValues

    ImportWorkbook.Close
    
'    ThisWorkbook.sheets("xchoices").Cells.NumberFormat = "@"
    Call convert_to_lower(wb.sheets("xchoices"))
    Call delete_irrelevant_columns("xchoices")
    
    ' trime all three columns:
    Dim rng As Range
    
    For i = 1 To 3
        Set rng = columns(i)
        rng.Value = Application.Trim(rng)
    Next i

    label_col = this_gen_column_number("label::english", "xchoices")
     
    If label_col > 0 Then
        wb.sheets("xchoices").Cells(1, label_col).Value = "label"
    End If
      
    If wb.sheets("xchoices").Range("A1") <> "list_name" Or wb.sheets("xchoices").Range("B1") <> "name" Or _
        wb.sheets("xchoices").Range("C1") <> "label" Then
        Call xchoices_column_order
        
    End If
    
'    ThisWorkbook.sheets("xchoices").columns("B").NumberFormat = "@"
    
    If wb.sheets("xchoices").Range("A1") <> "list_name" Or wb.sheets("xchoices").Range("B1") <> "name" Or _
        wb.sheets("xchoices").Range("C1") <> "label" Then
        MsgBox "Please check the choices sheet of the tool.", vbInformation
        End
    End If
    
End Sub

Sub xchoices_column_order()
    Dim search As Range
    Dim cnt As Integer
    Dim colOrdr As Variant
    Dim indx As Integer
    
    colOrdr = Array("list_name", "name", "label")
    
    cnt = 1
    
    For indx = LBound(colOrdr) To UBound(colOrdr)
        Set search = rows("1:1").Find(colOrdr(indx), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not search Is Nothing Then
            If search.Column <> cnt Then
                search.EntireColumn.Cut
                ThisWorkbook.sheets("xchoices").columns(cnt).Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
        cnt = cnt + 1
        End If
    Next indx
End Sub

Sub make_survey_choice()
    
    Dim s_rng As Range, c_rng As Range
    Dim last_row_survey As Long
    Dim last_row_choice As Long
    Dim last_row_xsurvey_choices As Long
    Dim i As Long
    Dim j As Long
    Dim new_row As Long
    Dim question As Variant
    Dim question_label As String
    Dim the_list As Variant
    Dim choice_list As Variant
    Dim chioce_value As String
    Dim chioce_label As String
    Dim last_row As Long

    last_row_survey = ThisWorkbook.sheets("xsurvey").Cells(rows.count, 1).End(xlUp).Row
    last_row_choice = ThisWorkbook.sheets("xchoices").Cells(rows.count, 1).End(xlUp).Row
    Debug.Print last_row_survey, last_row_choice
    Set s_rng = ThisWorkbook.sheets("xsurvey").Range("A2:C" & last_row_survey)

    Set c_rng = ThisWorkbook.sheets("xchoices").Range("A2:C" & last_row_choice)
    
    ThisWorkbook.sheets("xsurvey_choices").Cells.Clear
    ThisWorkbook.sheets("xsurvey_choices").Cells.NumberFormat = "@"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 1) = "type"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 2) = "question"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 3) = "question_label"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 4) = "choice"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 5) = "choice_label"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 6) = "question_choice"
    
    last_row_xsurvey_choices = ThisWorkbook.sheets("xsurvey_choices").Cells(rows.count, 1).End(xlUp).Row
    
    i = 0
    
    new_row = 1
    
    For Each question In s_rng.columns(2).Cells
        i = i + 1
        question_label = s_rng.columns(3).rows(i)
'        Debug.Print question, show_list_name(CStr(question)); question_type(CStr(question))
        
        If question_type(CStr(question)) = "integer" Or question_type(CStr(question)) = "decimal" Or _
            question_type(CStr(question)) = "calculate" Then
            new_row = new_row + 1
            ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 1) = question_type(CStr(question))
            ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 2) = question
            ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 3) = question_label
    
        ElseIf left(question_type(CStr(question)), 7) = "select_" Then
            the_list = show_list_name(CStr(question))
            j = 0
            For Each choice_list In c_rng.columns(1).Cells
                j = j + 1
                If the_list = choice_list Then
                    chioce_value = c_rng.columns(2).rows(j)
                    chioce_label = c_rng.columns(3).rows(j)
                    new_row = new_row + 1
                    
                    ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 1) = question_type(CStr(question))
                    ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 2) = question
                    ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 3) = question_label
                    ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 4) = chioce_value
                    ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 5) = chioce_label
                End If
        
            Next choice_list
            
        End If
        
    Next question
    
    ' final steps
    last_row = ThisWorkbook.Worksheets("xsurvey_choices").Cells(rows.count, 1).End(xlUp).Row
    
    ' make question_choice concatonation
    ThisWorkbook.Worksheets("xsurvey_choices").columns("F:F").NumberFormat = "General"
    ThisWorkbook.Worksheets("xsurvey_choices").Range("F2").FormulaR1C1 = "=RC[-4]&RC[-2]"
    ThisWorkbook.Worksheets("xsurvey_choices").Range("F2").AutoFill _
            Destination:=ThisWorkbook.Worksheets("xsurvey_choices").Range("F2:F" & last_row)

    ThisWorkbook.Worksheets("xsurvey_choices").columns("F:F").Copy
    ThisWorkbook.Worksheets("xsurvey_choices").columns("F:F").PasteSpecial _
            Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
     
End Sub

Sub delete_irrelevant_columns(sheet_name As String)
    Dim keepColumn As Boolean
    Dim currentColumn As Integer
    Dim columnHeading As String
    Dim temp_ws As Worksheet
'    Set temp_ws = Worksheets(sheet_name)
    Set temp_ws = ThisWorkbook.Worksheets(sheet_name)

    currentColumn = 1

    While currentColumn <= temp_ws.UsedRange.columns.count
        columnHeading = temp_ws.UsedRange.Cells(1, currentColumn).Value
        
        ' check whether to keep the column
        keepColumn = False
        If columnHeading = "list_name" Then keepColumn = True
        If columnHeading = "type" Then keepColumn = True
        If columnHeading = "name" Then keepColumn = True
        If columnHeading = "label::english" Then
            keepColumn = True
        Else
            If columnHeading = "label" Then keepColumn = True
        End If
        
        If keepColumn Then
            'if yes then skip to the next column
            currentColumn = currentColumn + 1
        Else
            'if no delete the column
            temp_ws.columns(currentColumn).Delete
        End If

        'lastly an escape in case the sheet has no columns left
        If (temp_ws.UsedRange.Address = "$A$1") And (temp_ws.Range("$A$1").text = "") Then Exit Sub
    Wend

End Sub

Function match_type(col_name As String) As String
    On Error Resume Next
    'Declare the variables
    Dim temp_ws As Worksheet
    Dim TableRange As Range
    Dim matchRow As Variant
    
    'Set the variables
    Set temp_ws = ThisWorkbook.Worksheets("xsurvey")
    Set TableRange = temp_ws.Range("A:B")
    
    'Use Application.Match instead of WorksheetFunction.Match to avoid errors
    matchRow = Application.Match(col_name, TableRange.columns(2), 0)
    
    'Check if the match was successful
    If Not IsError(matchRow) Then
        match_type = TableRange.Cells(matchRow, 1)
    Else
        match_type = ""
    End If
End Function

Sub add_question_label(question_name As String)
    On Error Resume Next
    Dim main_ws As Worksheet
    Dim DestSheet As Worksheet
    Dim new_col As Long
    Dim q_type As String
    Dim old_col As Long
    Dim question_col_number As Long
    Dim qeustion_col As String
    Dim qeustion_label_col As String
    Dim vaArray As Variant
    Dim key_name As Variant
    Dim last_row_choice As Long
    Dim last_row_dt As Long
    Dim last_row_choices As Long
    Dim SourceRange As Range
    Dim criteria As String
    Dim filtered_col As Long
    
    ' for using this function in other sheet the below line has been dissabled. and main_ws has been set to activeSheet.
    ' Set main_ws = sheets(find_main_data)
    Set main_ws = ActiveSheet
    q_type = match_type(question_name)
    
    If left(q_type, 19) = "select_one_external" Then
        MsgBox "Sorry! This is a select_one_external data type in the tool, " & vbCrLf & _
               "This data type is not supported!  ", vbInformation
        Exit Sub
    End If
    
    If left(q_type, 15) = "select_multiple" Then
        MsgBox "This is a select_multiple data type in the tool, " & vbCrLf & _
               "This data type is not supported!  ", vbInformation
        Exit Sub
    End If
    
    If left(q_type, 10) = "select_one" Then
        
        Call clear_filter(main_ws)
        
        ' check if there is any old label exist or not.
        ' if there is an old one it will be deleted from the dataset
        old_col = gen_column_number(question_name & "_label", main_ws.name)
        question_col_number = gen_column_number(question_name, main_ws.name)
        
       Debug.Print old_col
        
        If old_col <> 0 Then
            main_ws.columns(old_col).Delete Shift:=xlToLeft
        End If
        
        new_col = gen_column_number(question_name, main_ws.name) + 1
        Debug.Print question_name, new_col & " gen: " & gen_column_number(question_name, main_ws.name)
        
        main_ws.columns(new_col).Insert
        main_ws.Cells(1, new_col).Value = question_name & "_label"
        main_ws.columns(new_col).NumberFormat = "General"

        main_ws.Select
        
        qeustion_col = column_letter(question_name)
        qeustion_label_col = column_letter(question_name & "_label")
        
        vaArray = Split(q_type, " ")
        key_name = vaArray(UBound(vaArray))
        
        last_row_choice = ThisWorkbook.Worksheets("xchoices").Cells(rows.count, 1).End(xlUp).Row

        Set SourceRange = ThisWorkbook.Worksheets("xchoices").Range("A1:C" & last_row_choice)
        
        ' check if redeem sheet exist
        If worksheet_exists("redeem") <> True Then
            Call create_sheet(main_ws.name, "redeem")
        End If
    
        Set DestSheet = Worksheets("redeem")

        'Apply the filter on the source range
        SourceRange.AutoFilter field:=1, Criteria1:=key_name

        'Copy only the visible cells to the destination sheet
        SourceRange.SpecialCells(xlCellTypeVisible).Copy DestSheet.Range("A1")
        SourceRange.AutoFilter
        Application.CutCopyMode = False
        
'        DestSheet.Select
'        DestSheet.columns("B:B").Select
'        Selection.TextToColumns Destination:=DestSheet.Range("B1"), DataType:=xlDelimited, _
'            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
'            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
'            :=Array(1, 1), TrailingMinusNumbers:=True
'        Application.CutCopyMode = False

        ' apply lookup formula:
        main_ws.Select
        
        ' need for improvement
        last_row_dt = main_ws.Cells(rows.count, question_col_number).End(xlUp).Row
        last_row_choices = sheets("redeem").Cells(rows.count, 1).End(xlUp).Row

        main_ws.Range(qeustion_label_col & "2:" & qeustion_label_col & CStr(last_row_dt)).Formula = _
            "=VLOOKUP(" & qeustion_col & "2,'redeem'!B$2:C$" & last_row_choices & ",2,False)"
        
        ' convert formula to values:
        columns(qeustion_label_col & ":" & qeustion_label_col).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                               :=False, Transpose:=False
        
        ' replace #N/A with "":
        main_ws.columns(qeustion_label_col & ":" & qeustion_label_col).Select
        Selection.Replace What:="#N/A", replacement:="", LookAt:=xlPart, _
                          SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

        main_ws.Range(qeustion_label_col & "1").Select
    Else
        MsgBox "The label was not found!               ", vbInformation
    End If
           
    Application.DisplayAlerts = False
    
    If worksheet_exists("redeem") Then
        sheets("redeem").Delete
    End If
    
    Application.DisplayAlerts = True

End Sub


Sub add_label()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim last_col As Long
    Dim selectedRange As Range
    Dim data_col_number As Long
    
    Set selectedRange = Application.Selection
    
    If ThisWorkbook.Worksheets("xsurvey").Range("A1") = vbNullString Then
        MsgBox "Please import the tool from the setting!   ", vbInformation
        End
    End If
    
    ' check if the selected range is in one column
    If selectedRange.columns.count > 1 Then
        MsgBox "Please select one column.      ", vbInformation
        Exit Sub
    End If
    
    data_col_number = selectedRange.Column
    Application.Cells(1, data_col_number).Select
    Set selectedRange = Application.Selection
    
    Call add_question_label(selectedRange.Value)
    Application.ScreenUpdating = True
End Sub

Function question_type(col_name As String) As Variant
    On Error Resume Next
    Dim temp_ws As Worksheet
    Dim TableRange As Range
    Dim matchRow As Variant
    Dim res As String
    Set temp_ws = ThisWorkbook.Worksheets("xsurvey")
    Set TableRange = temp_ws.Range("A:B")
    
    matchRow = Application.Match(col_name, TableRange.columns(2), 0)
    
    'Check if the match was successful
    If Not IsError(matchRow) Then
        res = TableRange.Cells(matchRow, 1)
        question_type = Split(res, " ")(0)
    Else
        question_type = ""
    End If
End Function

Function show_list_name(col_name As String) As Variant
    On Error Resume Next
    Dim temp_ws As Worksheet
    Dim TableRange As Range
    Dim matchRow As Variant
    Dim res As String

    Set temp_ws = ThisWorkbook.Worksheets("xsurvey")
    Set TableRange = temp_ws.Range("A:B")
    
    matchRow = Application.Match(col_name, TableRange.columns(2), 0)
    
    'Check if the match was successful
    If Not IsError(matchRow) Then
        res = TableRange.Cells(matchRow, 1)
        show_list_name = Split(res, " ")(1)
    Else
        show_list_name = ""
    End If
End Function

Sub convert_to_lower(ws As Worksheet)
    Dim rng As Range
    Dim cell As Range

    Set rng = ws.Range("1:1")

    For Each cell In rng
        cell.Value = LCase(cell.Value)
        cell.Value = Trim(cell.Value)
        
        If left(cell.Value, 14) = "label::english" Then
             cell.Value = "label::english"
        End If
    Next cell
End Sub
