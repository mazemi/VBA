Attribute VB_Name = "Tools_module"
Option Explicit

Sub newImportTool(the_path As String, sheetName As String)
    Dim externalWorkbook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim headerArr As Variant
    Dim sourceData As Variant
    Dim targetData As Variant
    Dim sourceHeaderRow As Range
    Dim headerIndex As Long
    Dim lastRow As Long
    Dim requiredColumns As Variant
    Dim headerData() As Variant
    
    If sheetName = "survey" Then
        requiredColumns = Array("type", "name", "label::english")
    Else
        requiredColumns = Array("list_name", "name", "label::english")
    End If
        
    Set externalWorkbook = Workbooks.Open(the_path)
    Set sourceSheet = externalWorkbook.sheets(sheetName)
    
    If sourceSheet.AutoFilterMode Then
        sourceSheet.AutoFilterMode = False
    End If

    sourceSheet.Rows.Hidden = False
    
    Set sourceHeaderRow = sourceSheet.Rows(1)
    headerArr = sourceHeaderRow.value
    
    ReDim headerData(1 To 1, 1 To UBound(headerArr, 2)) As Variant
    Dim i As Long
    For i = 1 To UBound(headerArr, 2)
        headerData(1, i) = LCase(Trim(headerArr(1, i)))
    Next i
    
    headerIndex = Application.Match(requiredColumns(0), headerData, 0)
    
    If Not IsError(headerIndex) Then
        lastRow = sourceSheet.Cells(sourceSheet.Rows.count, headerIndex).End(xlUp).Row
        sourceData = sourceSheet.Range(sourceSheet.Cells(1, headerIndex), sourceSheet.Cells(lastRow, headerIndex)).value
        
        Set targetSheet = ThisWorkbook.sheets("x" & sheetName)
        targetSheet.Cells.Clear
        targetSheet.Cells(1, 1).Resize(lastRow, 1).value = sourceData
        
    Else
        Debug.Print "Column 1 not found."
    End If
    
    headerIndex = Application.Match(requiredColumns(1), headerData, 0)
    
    If Not IsError(headerIndex) Then
        sourceData = sourceSheet.Range(sourceSheet.Cells(1, headerIndex), sourceSheet.Cells(lastRow, headerIndex)).value
        targetSheet.Cells(1, 2).Resize(lastRow, 1).value = sourceData
    Else
        Debug.Print "Column 2 not found."
    End If
    
    headerIndex = 0
    
    For i = 1 To UBound(headerData, 2)
        If left(headerData(1, i), 14) = requiredColumns(2) Then
            headerIndex = i
            Exit For
        End If
    Next i
    
    If headerIndex = 0 Then
        For i = 1 To UBound(headerData, 2)
            If left(headerData(1, i), 5) = "label" Then
                headerIndex = i
                Exit For
            End If
        Next i
    End If
      
    If Not IsError(headerIndex) Then
        sourceData = sourceSheet.Range(sourceSheet.Cells(1, headerIndex), sourceSheet.Cells(lastRow, headerIndex)).value
        targetSheet.Cells(1, 3).Resize(lastRow, 1).value = sourceData
    Else
        Debug.Print "Column 3 not found."
    End If
    
    externalWorkbook.Close SaveChanges:=False
    
    If sheetName = "survey" Then
        targetSheet.Range("A1:C1") = Array("type", "name", "label")
    ElseIf sheetName = "choices" Then
        targetSheet.Range("A1:C1") = Array("list_name", "name", "label")
    End If
    
    Call DeleteEmptyRows("x" & sheetName)

End Sub

Sub DeleteEmptyRows(xsheet As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.sheets(xsheet)
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    For i = lastRow To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(i)) = 0 Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub

Sub make_survey_choice_new()
    
    Dim s_ws As Worksheet, c_ws As Worksheet, sc_ws As Worksheet
    Dim s_rng As Range, c_rng As Range
    Dim last_row_survey As Long, last_row_choice As Long
    Dim i As Long, j As Long, new_row As Long, last_row As Long
    Dim question As Range
    Dim question_label As String
    Dim the_list As String
    Dim choice_list As Range
    Dim choice_value As String, choice_label As String

    Set s_ws = ThisWorkbook.sheets("xsurvey")
    Set c_ws = ThisWorkbook.sheets("xchoices")
    Set sc_ws = ThisWorkbook.sheets("xsurvey_choices")

    last_row_survey = s_ws.Cells(s_ws.Rows.count, 1).End(xlUp).Row
    last_row_choice = c_ws.Cells(c_ws.Rows.count, 1).End(xlUp).Row
    
    Set s_rng = s_ws.Range("A2:C" & last_row_survey)
    Set c_rng = c_ws.Range("A2:C" & last_row_choice)
    
    sc_ws.Cells.Clear
    
    With sc_ws.Range("A1:F1")
        .value = Array("type", "question", "question_label", "choice", "choice_label", "question_choice")
        .NumberFormat = "@"
    End With

    new_row = 1
    
    For Each question In s_rng.Columns(2).Cells
        i = i + 1
        question_label = s_rng.Cells(i, 3).value
        
        If question_type(CStr(question)) Like "integer|decimal|calculate" Then
            new_row = new_row + 1
            With sc_ws.Rows(new_row)
                .Cells(1).value = question_type(CStr(question))
                .Cells(2).value = question.value
                .Cells(3).value = question_label
            End With
        ElseIf left(question_type(CStr(question)), 7) = "select_" Then
            the_list = show_list_name(CStr(question))
            For Each choice_list In c_rng.Columns(1).Cells
                j = j + 1
                If the_list = choice_list.value Then
                    choice_value = c_rng.Cells(j, 2).value
                    choice_label = c_rng.Cells(j, 3).value
                    new_row = new_row + 1
                    With sc_ws.Rows(new_row)
                        .Cells(1).value = question_type(CStr(question))
                        .Cells(2).value = question.value
                        .Cells(3).value = question_label
                        .Cells(4).value = choice_value
                        .Cells(5).value = choice_label
                    End With
                End If
            Next choice_list
        End If
    Next question
    
    last_row = sc_ws.Cells(sc_ws.Rows.count, 1).End(xlUp).Row
    
    ' Concatenate type and question_label to create question_choice
    With sc_ws.Range("F2:F" & last_row)
        .Formula = "=A2&C2"
        .value = .value
    End With

    Call check_choice_duplicates

End Sub

Sub make_survey_choice()
    Dim s_ws As Worksheet, c_ws As Worksheet, sc_ws As Worksheet
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

    Set s_ws = ThisWorkbook.sheets("xsurvey")
    Set c_ws = ThisWorkbook.sheets("xchoices")
    Set sc_ws = ThisWorkbook.sheets("xsurvey_choices")

    last_row_survey = s_ws.Cells(s_ws.Rows.count, 1).End(xlUp).Row
    last_row_choice = c_ws.Cells(c_ws.Rows.count, 1).End(xlUp).Row
    
    Set s_rng = s_ws.Range("A2:C" & last_row_survey)
    Set c_rng = c_ws.Range("A2:C" & last_row_choice)
    
    sc_ws.Cells.Clear
    
    With sc_ws.Range("A1:F1")
        .value = Array("type", "question", "question_label", "choice", "choice_label", "question_choice")
        .NumberFormat = "@"
    End With
    
    last_row_xsurvey_choices = sc_ws.Cells(Rows.count, 1).End(xlUp).Row
    
    i = 0
    
    new_row = 1
    
    For Each question In s_rng.Columns(2).Cells
        i = i + 1
        question_label = s_rng.Columns(3).Rows(i)
        
        If question_type(CStr(question)) = "integer" Or question_type(CStr(question)) = "decimal" Or _
            question_type(CStr(question)) = "calculate" Then
            new_row = new_row + 1
            With sc_ws
                .Cells(new_row, 1) = question_type(CStr(question))
                .Cells(new_row, 2) = question
                .Cells(new_row, 3) = question_label
            End With
    
        ElseIf left(question_type(CStr(question)), 7) = "select_" Then
            the_list = show_list_name(CStr(question))
            j = 0
            For Each choice_list In c_rng.Columns(1).Cells
                j = j + 1
                If the_list = choice_list Then
                    chioce_value = c_rng.Columns(2).Rows(j)
                    chioce_label = c_rng.Columns(3).Rows(j)
                    new_row = new_row + 1
                    With sc_ws
                        .Cells(new_row, 1) = question_type(CStr(question))
                        .Cells(new_row, 2) = question
                        .Cells(new_row, 3) = question_label
                        .Cells(new_row, 4) = chioce_value
                        .Cells(new_row, 5) = chioce_label
                    End With
                End If
        
            Next choice_list
            
        End If
        
    Next question
    
    last_row = sc_ws.Cells(Rows.count, 1).End(xlUp).Row
    
    With sc_ws.Range("F2:F" & last_row)
        .NumberFormat = "General"
        .Formula = "=B2&D2"
        .value = .value
    End With
    
    Application.CutCopyMode = False
     
    Call check_choice_duplicates
    
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

        Set SourceRange = Nothing

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
    Dim header_value As String
    
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
    header_value = ActiveSheet.Cells(1, data_col_number).value
    Call add_question_label(header_value)
    
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
        rng.RemoveDuplicates Columns:=Array(1, 2), header:=xlNo
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
