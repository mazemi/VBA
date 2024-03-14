Attribute VB_Name = "tools_module"
Option Explicit

Sub import_survey(tools_path As String)
    On Error Resume Next
    DoEvents
    setting_form.bar.Width = 20
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim ImportWorkbook As Workbook
    Dim i As Long
    Dim label_col As Long
    Dim rng As Range
    
    Set wb = ThisWorkbook
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
    DoEvents
    setting_form.bar.Width = 25
    
    Call convert_to_lower(wb.sheets("xsurvey"))
    Call delete_irrelevant_columns("xsurvey")
    
    ' trime all three columns
    For i = 1 To 3
        Set rng = Columns(i)
        rng.value = Application.Trim(rng)
    Next i

    label_col = this_gen_column_number("label::english", "xsurvey")
     
    If label_col > 0 Then
        wb.sheets("xsurvey").Cells(1, label_col).value = "label"
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
    Dim colOrder As Variant
    Dim indx As Integer
    
    colOrder = Array("type", "name", "label")
    
    cnt = 1
    
    For indx = LBound(colOrder) To UBound(colOrder)
        Set search = Rows("1:1").Find(colOrder(indx), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not search Is Nothing Then
            If search.Column <> cnt Then
                search.EntireColumn.Cut
                ThisWorkbook.sheets("xsurvey").Columns(cnt).Insert Shift:=xlToRight
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

    ImportWorkbook.Close
    DoEvents
    setting_form.bar.Width = 35
    
'    ThisWorkbook.sheets("xchoices").Cells.NumberFormat = "@"
    Call convert_to_lower(wb.sheets("xchoices"))
    Call delete_irrelevant_columns("xchoices")
    
    ' trime all three columns:
    Dim rng As Range
    
    For i = 1 To 3
        Set rng = Columns(i)
        rng.value = Application.Trim(rng)
    Next i

    label_col = this_gen_column_number("label::english", "xchoices")
     
    If label_col > 0 Then
        wb.sheets("xchoices").Cells(1, label_col).value = "label"
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
    Dim colOrder As Variant
    Dim indx As Integer
    
    colOrder = Array("list_name", "name", "label")
    
    cnt = 1
    
    For indx = LBound(colOrder) To UBound(colOrder)
        Set search = Rows("1:1").Find(colOrder(indx), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not search Is Nothing Then
            If search.Column <> cnt Then
                search.EntireColumn.Cut
                ThisWorkbook.sheets("xchoices").Columns(cnt).Insert Shift:=xlToRight
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

    last_row_survey = ThisWorkbook.sheets("xsurvey").Cells(Rows.count, 1).End(xlUp).Row
    last_row_choice = ThisWorkbook.sheets("xchoices").Cells(Rows.count, 1).End(xlUp).Row
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
    
    last_row_xsurvey_choices = ThisWorkbook.sheets("xsurvey_choices").Cells(Rows.count, 1).End(xlUp).Row
    
    i = 0
    
    new_row = 1
    
    For Each question In s_rng.Columns(2).Cells
        i = i + 1
        question_label = s_rng.Columns(3).Rows(i)
        
        If question_type(CStr(question)) = "integer" Or question_type(CStr(question)) = "decimal" Or _
            question_type(CStr(question)) = "calculate" Then
            new_row = new_row + 1
            ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 1) = question_type(CStr(question))
            ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 2) = question
            ThisWorkbook.sheets("xsurvey_choices").Cells(new_row, 3) = question_label
    
        ElseIf left(question_type(CStr(question)), 7) = "select_" Then
            the_list = show_list_name(CStr(question))
            j = 0
            For Each choice_list In c_rng.Columns(1).Cells
                j = j + 1
                If the_list = choice_list Then
                    chioce_value = c_rng.Columns(2).Rows(j)
                    chioce_label = c_rng.Columns(3).Rows(j)
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
    last_row = ThisWorkbook.Worksheets("xsurvey_choices").Cells(Rows.count, 1).End(xlUp).Row
    
    ' make question_choice concatonation
    ThisWorkbook.Worksheets("xsurvey_choices").Columns("F:F").NumberFormat = "General"
    ThisWorkbook.Worksheets("xsurvey_choices").Range("F2").FormulaR1C1 = "=RC[-4]&RC[-2]"
    ThisWorkbook.Worksheets("xsurvey_choices").Range("F2").AutoFill _
            Destination:=ThisWorkbook.Worksheets("xsurvey_choices").Range("F2:F" & last_row)

    ThisWorkbook.Worksheets("xsurvey_choices").Columns("F:F").Copy
    ThisWorkbook.Worksheets("xsurvey_choices").Columns("F:F").PasteSpecial _
            Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
     
    Call check_choice_duplicates
    
End Sub

Sub delete_irrelevant_columns(SHEET_NAME As String)
    Dim keepColumn As Boolean
    Dim currentColumn As Integer
    Dim columnHeading As String
    Dim temp_ws As Worksheet
'    Set temp_ws = Worksheets(sheet_name)
    Set temp_ws = ThisWorkbook.Worksheets(SHEET_NAME)

    currentColumn = 1

    While currentColumn <= temp_ws.UsedRange.Columns.count
        columnHeading = temp_ws.UsedRange.Cells(1, currentColumn).value
        
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
            temp_ws.Columns(currentColumn).Delete
        End If

        'lastly an escape in case the sheet has no columns left
        If (temp_ws.UsedRange.Address = "$A$1") And (temp_ws.Range("$A$1").Text = "") Then Exit Sub
    Wend

End Sub

Function match_type(col_name As String) As String
    On Error Resume Next
    Dim temp_ws As Worksheet
    Dim TableRange As Range
    Dim matchRow As Variant
    
    Set temp_ws = ThisWorkbook.Worksheets("xsurvey")
    Set TableRange = temp_ws.Range("A:B")
    
    matchRow = Application.Match(col_name, TableRange.Columns(2), 0)
    
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
        old_col = gen_column_number(question_name & "_label", main_ws.Name)
        question_col_number = gen_column_number(question_name, main_ws.Name)
        
       Debug.Print old_col
        
        If old_col <> 0 Then
            main_ws.Columns(old_col).Delete Shift:=xlToLeft
        End If
        
        new_col = gen_column_number(question_name, main_ws.Name) + 1
        Debug.Print question_name, new_col & " gen: " & gen_column_number(question_name, main_ws.Name)
        
        main_ws.Columns(new_col).Insert
        main_ws.Cells(1, new_col).value = question_name & "_label"
        main_ws.Columns(new_col).NumberFormat = "General"

        main_ws.Select
        
        qeustion_col = column_letter(question_name)
        qeustion_label_col = column_letter(question_name & "_label")
        
        vaArray = Split(q_type, " ")
        key_name = vaArray(UBound(vaArray))
        
        last_row_choice = ThisWorkbook.Worksheets("xchoices").Cells(Rows.count, 1).End(xlUp).Row

        Set SourceRange = ThisWorkbook.Worksheets("xchoices").Range("A1:C" & last_row_choice)
        
        ' check if redeem sheet exist
        If worksheet_exists("redeem") <> True Then
            Call create_sheet(main_ws.Name, "redeem")
        End If
    
        Set DestSheet = Worksheets("redeem")

        SourceRange.AutoFilter Field:=1, Criteria1:=key_name

        SourceRange.SpecialCells(xlCellTypeVisible).Copy DestSheet.Range("A1")
        SourceRange.AutoFilter
        Application.CutCopyMode = False

        main_ws.Select
        
        ' need for improvement
        last_row_dt = main_ws.Cells(Rows.count, question_col_number).End(xlUp).Row
        last_row_choices = sheets("redeem").Cells(Rows.count, 1).End(xlUp).Row

        main_ws.Range(qeustion_label_col & "2:" & qeustion_label_col & CStr(last_row_dt)).Formula = _
            "=VLOOKUP(" & qeustion_col & "2,'redeem'!B$2:C$" & last_row_choices & ",2,False)"
        
        ' convert formula to values:
        Columns(qeustion_label_col & ":" & qeustion_label_col).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                               :=False, Transpose:=False
        
        ' replace #N/A with "":
        main_ws.Columns(qeustion_label_col & ":" & qeustion_label_col).Select
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
    If selectedRange.Columns.count > 1 Then
        MsgBox "Please select one column.      ", vbInformation
        Exit Sub
    End If
    
    data_col_number = selectedRange.Column
    Application.Cells(1, data_col_number).Select
    Set selectedRange = Application.Selection
    
    Call add_question_label(selectedRange.value)
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
    
    matchRow = Application.Match(col_name, TableRange.Columns(2), 0)
    
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
    
    matchRow = Application.Match(col_name, TableRange.Columns(2), 0)
    
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
        cell.value = LCase(cell.value)
        cell.value = Trim(cell.value)
        
        If left(cell.value, 14) = "label::english" Then
             cell.value = "label::english"
        End If
    Next cell
End Sub

Sub check_choice_duplicates()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim checkRange As Range
    Dim cell As Variant
    Dim value As Variant
    Dim valueCount As Integer
    Dim c As Variant
    Dim all_arr() As Variant
    Dim arr() As Variant
    Dim i As Double
    Dim j As Double
    Dim k As Double
    Dim msg As String
    Dim has_duplicate As Boolean
    Dim res_arr() As Variant
    Dim rng As Range
    
    Set ws = ThisWorkbook.Worksheets("xsurvey_choices")
    ws.Columns("M:N").Clear
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    Set checkRange = ws.Range("F2:F" & lastRow)
    
    all_arr = ws.Range("A1").CurrentRegion.value
    arr = checkRange.value
    arr = Application.Transpose(arr)
    
    has_duplicate = False
    ReDim res_arr(1 To 1, 1 To 2)
    
    i = 2 ' exclude the heder by skipping to 2
    For Each cell In arr
        value = cell
        
        If Not IsEmpty(value) Then
            valueCount = 0
            
            For Each c In arr
                If c = value Then
                    valueCount = valueCount + 1
                End If
            Next c
            
            If valueCount > 1 Then
                has_duplicate = True
                j = j + 1
                'Debug.Print "duplicates", all_arr(i, 2), all_arr(i, 4)
                ws.Cells(j, "M") = all_arr(i, 2)
                ws.Cells(j, "N") = all_arr(i, 4)
            End If
        End If
        i = i + 1
    Next cell
       
    If has_duplicate Then
        Set rng = ws.Range("M1").CurrentRegion
        rng.RemoveDuplicates Columns:=Array(1, 2), Header:=xlNo
        Set rng = ws.Range("M1").CurrentRegion
     
        For k = 1 To rng.Rows.count
            msg = msg & vbCrLf & "question: " & rng.Cells(k, 1).value & " , choice: " & rng.Cells(k, 2).value
        Next k
                
        If msg <> vbNullString Then
            MsgBox "There are some repetitive questions or choices in your KOBO tool." & _
                " Please first check the tool and import it again." & _
                " If you would like to ignore this message, these questions won't be analyzed." & vbCrLf & _
                msg, vbInformation
        End If
    End If
    ws.Columns("M:N").Clear
End Sub
