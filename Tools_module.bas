Attribute VB_Name = "tools_module"
Sub import_survey(tools_path As String)
    '    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim wb As Workbook
    Dim ImportWorkbook As Workbook
'    Set wb = ActiveWorkbook
    Set wb = ThisWorkbook
    
    ' clear survey sheet
    ThisWorkbook.sheets("xsurvey").Cells.Clear

    Set ImportWorkbook = Workbooks.Open(Filename:=tools_path)
    
    If (ImportWorkbook.Worksheets("survey").AutoFilterMode And ImportWorkbook.Worksheets("survey").FilterMode) Or _
       ImportWorkbook.Worksheets("survey").FilterMode Then
        ImportWorkbook.Worksheets("survey").ShowAllData
    End If
    
    ImportWorkbook.Worksheets("survey").UsedRange.Copy
    wb.Worksheets("xsurvey").Range("A1").PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, transpose:=False

    ImportWorkbook.Close

    ' trime all three columns:
    Dim rng As Range
    
    For i = 1 To 3
        Set rng = columns(i)
        rng.value = Application.Trim(rng)
    Next i

    Call delete_irrelevant_columns("xsurvey")

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub import_choices(tools_path As String)
    ' On Error Resume Next
    Application.DisplayAlerts = False

    Application.ScreenUpdating = False
    Dim wb As Workbook
'    Set wb = ActiveWorkbook
    Set wb = ThisWorkbook
    Dim ImportWorkbook As Workbook

    ThisWorkbook.sheets("xchoices").Cells.Clear
        
    Set ImportWorkbook = Workbooks.Open(Filename:=tools_path)
    
    If (ImportWorkbook.Worksheets("choices").AutoFilterMode And ImportWorkbook.Worksheets("choices").FilterMode) Or _
       ImportWorkbook.Worksheets("choices").FilterMode Then
        ImportWorkbook.Worksheets("choices").ShowAllData
    End If
    
    ImportWorkbook.Worksheets("choices").UsedRange.Copy
    wb.Worksheets("xchoices").Range("A1").PasteSpecial _
        Paste:=xlPasteValues, SkipBlanks:=False

    ImportWorkbook.Close

    Call delete_irrelevant_columns("xchoices")

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub make_survey_choice()
    On Error Resume Next
    Dim survey_df As New Array2D
    Dim choices_df As New Array2D
    Dim full_df As New Array2D
    Dim int_df As New Array2D, dec_df As New Array2D, calc_df As New Array2D
    Dim select_one_df As New Array2D, select_multiple_df As New Array2D
    Dim good_df As New Array2D
    Dim last_row As Long
    
    full_df.Data = Null
    
    full_df.insertColumnsBlank 1, 4
    
    Dim s_rng As Range, c_rng As Range
    
    ' edited
    Set s_rng = ThisWorkbook.sheets("xsurvey").Range("A1").CurrentRegion
    Set c_rng = ThisWorkbook.sheets("xchoices").Range("A1").CurrentRegion
    
    survey_df.dataFromRange rg:=s_rng, removeHeader:=True
    survey_df.insertColumnsBlank survey_df.columnCount + 1
    choices_df.dataFromRange rg:=c_rng, removeHeader:=True
    
    For i = 1 To survey_df.RowCount
        If InStr(survey_df.value(i, 1), " ") Then
            Leftx = left(survey_df.value(i, 1), InStrRev(survey_df.value(i, 1), " ") - 1)
            Rightx = Right(survey_df.value(i, 1), Len(survey_df.value(i, 1)) - InStrRev(survey_df.value(i, 1), " "))
            survey_df.value(i, 1) = Leftx
            survey_df.value(i, survey_df.columnCount) = Rightx
        End If
    Next i
    
    Dim k
    k = 1
    
    While k <= survey_df.RowCount
        M = full_df.RowCount
        
        For j = 1 To choices_df.RowCount
            
            If choices_df.value(j, 1) = survey_df.value(k, 4) Then
                n = full_df.RowCount
                full_df.value(n, 1) = survey_df.value(k, 1)
                full_df.value(n, 2) = survey_df.value(k, 2)
                full_df.value(n, 3) = survey_df.value(k, 3)
                full_df.value(n, 4) = choices_df.value(j, 2)
                full_df.value(n, 5) = choices_df.value(j, 3)
                full_df.insertRowsBlank full_df.RowCount + 1
            End If
           
        Next
        
        If survey_df.value(k, 1) <> "" Then
            full_df.value(M, 1) = survey_df.value(k, 1)
            full_df.value(M, 2) = survey_df.value(k, 2)
            full_df.value(M, 3) = survey_df.value(k, 3)
            full_df.insertRowsBlank full_df.RowCount + 1
        End If
        k = k + 1
    Wend
    
    int_df.Data = full_df.Filter("integer")
    dec_df.Data = full_df.Filter("decimal")
    calc_df.Data = full_df.Filter("calculate")
    select_one_df.Data = full_df.Filter("select_one")
    select_multiple_df.Data = full_df.Filter("select_multiple")
    
    full_df.Data = full_df.Filter("select_one")
    
    full_df.insertRows int_df.RowCount, int_df.Data
    
    good_df.Data = int_df.Data
    good_df.insertRows calc_df.RowCount, calc_df.Data
    good_df.insertRows select_one_df.RowCount, select_one_df.Data
    good_df.insertRows select_multiple_df.RowCount, select_multiple_df.Data
    
    ThisWorkbook.sheets("xsurvey_choices").Cells.Clear
    
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 1) = "type"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 2) = "question"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 3) = "question_label"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 4) = "choice"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 5) = "choice_label"
    ThisWorkbook.sheets("xsurvey_choices").Cells(1, 6) = "question_choice"

    good_df.writeDataToRange ThisWorkbook.sheets("xsurvey_choices").Range("A2")
    
    last_row = ThisWorkbook.Worksheets("xsurvey_choices").Cells(rows.count, 1).End(xlUp).row
    
    ' make question_choice concatonation
    ThisWorkbook.Worksheets("xsurvey_choices").Range("F2").FormulaR1C1 = "=RC[-4]&RC[-2]"
    ThisWorkbook.Worksheets("xsurvey_choices").Range("F2").AutoFill _
            Destination:=ThisWorkbook.Worksheets("xsurvey_choices").Range("F2:F" & last_row)

    ThisWorkbook.Worksheets("xsurvey_choices").columns("F:F").Copy
    ThisWorkbook.Worksheets("xsurvey_choices").columns("F:F").PasteSpecial _
            Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, transpose:=False
                                            

    Application.CutCopyMode = False
    
    Set survey_df = Nothing
    Set choices_df = Nothing
    Set full_df = Nothing
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
        columnHeading = temp_ws.UsedRange.Cells(1, currentColumn).value
        
        ' check whether to keep the column
        keepColumn = False
        If columnHeading = "list_name" Then keepColumn = True
        If columnHeading = "type" Then keepColumn = True
        If columnHeading = "name" Then keepColumn = True
        If columnHeading = "label::English" Then
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
        If (temp_ws.UsedRange.Address = "$A$1") And (temp_ws.Range("$A$1").Text = "") Then Exit Sub
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

        Dim SourceRange As Range
        Dim DestSheet As Worksheet
        Dim criteria As String
        Dim filtered_col As Long
        
        Call clear_filter(main_ws)
        
        ' check if there is any old label exist or not.
        ' if there is an old one it will be deleted from the dataset
        old_col = gen_column_number(question_name & "_label", main_ws.Name)
        question_col_number = gen_column_number(question_name, main_ws.Name)
        
        If old_col <> "" Then
            main_ws.columns(old_col).Delete Shift:=xlToLeft
        End If
        
        main_ws.columns(column_number(question_name) + 1).Insert
        main_ws.Cells(1, column_number(question_name) + 1).value = question_name & "_label"
        
        main_ws.Select
        
        qeustion_col = column_letter(question_name)
        qeustion_label_col = column_letter(question_name & "_label")
        
        vaArray = Split(q_type, " ")
        key_name = vaArray(UBound(vaArray))
        
        last_row_choice = ThisWorkbook.Worksheets("xchoices").Cells(rows.count, 1).End(xlUp).row

        Set SourceRange = ThisWorkbook.Worksheets("xchoices").Range("A1:C" & last_row_choice)
        
        ' check if redeem sheet exist
        If worksheet_exists("redeem") <> True Then
            Call create_sheet(main_ws.Name, "redeem")
        End If
    
        Set DestSheet = Worksheets("redeem")

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
        
        ' need for improvement
'        last_row_dt = main_ws.Cells(rows.count, 1).End(xlUp).row
        last_row_dt = main_ws.Cells(rows.count, question_col_number).End(xlUp).row
        
        last_row_choices = sheets("redeem").Cells(rows.count, 1).End(xlUp).row

        main_ws.Range(qeustion_label_col & "2:" & qeustion_label_col & CStr(last_row_dt)).Formula = _
            "=VLOOKUP(" & qeustion_col & "2,'redeem'!B$2:C$" & last_row_choices & ",2,False)"
        
        ' convert formula to values:
        columns(qeustion_label_col & ":" & qeustion_label_col).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                               :=False, transpose:=False
        
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

Sub lookup_label(question As String)
    On Error Resume Next
    Dim main_ws As Worksheet
    
    Set main_ws = ActiveWorkbook.ActiveSheet
    qeustion_col_number = column_number(question)
    qeustion_col = column_letter(question)
    qeustion_label_col = column_letter(question_name & "_label")

    last_row = main_ws.Cells(rows.count, qeustion_col_number).End(xlUp).row
    last_row_choices = sheets("redeem").Cells(rows.count, 1).End(xlUp).row

    main_ws.Range(qeustion_label_col & "2:" & qeustion_label_col & CStr(last_row)).Formula = _
            "=VLOOKUP(" & qeustion_col & "2,'redeem'!B$2:C$" & last_row_choices & ",2,False)"
End Sub

Sub add_label()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim last_col As Long
    Dim selectedRange As Range
    Set selectedRange = Application.Selection
    
    ' check if the tool exists
'    If Not worksheet_exists("survey") Or Not worksheet_exists("choices") Then
    If ThisWorkbook.Worksheets("xsurvey").Range("A1") = vbNullString Then
        MsgBox "Please import the tool from the setting!   ", vbInformation
        End
    End If
    
    ' check if the selected range is in one column
    If selectedRange.columns.count > 1 Then
        MsgBox "Please select one column.      ", vbInformation
        Exit Sub
    End If
    
    data_col_number = selectedRange.column
    Application.Cells(1, data_col_number).Select
    Set selectedRange = Application.Selection
    
    Call add_question_label(selectedRange.value)
    Application.ScreenUpdating = True
End Sub



Function question_type(col_name As String) As Variant
    On Error Resume Next
    'Declare the variables
    Dim temp_ws As Worksheet
    Dim TableRange As Range
    Dim matchRow As Variant

    'Set the variables
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


