Attribute VB_Name = "Public_module"
Global ISSUE_TEXT As String
Global PATTERN_CHECK_ACTION As Boolean
Global DATA_SHEET As String
Global PLAN_NUMBER As Long
Global SAMPLE_SHEET As String
Global DATA_STRATA As String
Global SAMPLE_STRATA As String
Global SAMPLE_POPULATION As String
Global CANCEL_PROCESS As Boolean
Global CURRENT_WORK_BOOK As Workbook
Global CHART_COUNT As Long
Global NUMERIC_CHART As Boolean
Global CHOSEN_CHART As Integer

Function InitializeExcludedSheets() As Variant
    Dim internalSheets() As Variant
    internalSheets = Array("result", "log_book", "analysis_list", "disaggregation_setting", _
                            "overall", "survey", "keen", "keen2", "indi_list", "temp_sheet", _
                            "choices", "datamerge", "dm_backend")
                            
     InitializeExcludedSheets = internalSheets
End Function

Function IsInArray(valToFind As String, arr As Variant) As Boolean
    Dim element As Variant
    For Each element In arr
        If element = valToFind Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function

Public Function number_to_letter(col_num As Long, input_ws As Worksheet) As String
    On Error Resume Next
    Dim vArr
    vArr = Split(input_ws.Cells(1, col_num).Address(True, False), "$")
    number_to_letter = vArr(0)
End Function

Function letter_to_number(col_name As String, input_ws As Worksheet)
    On Error Resume Next
    letter_to_number = input_ws.Range(col_name & 1).Column
End Function

Public Function worksheet_exists(sName As String) As Boolean
    On Error Resume Next
    worksheet_exists = Evaluate("ISREF('" & sName & "'!A1)")
End Function

Public Function column_number(column_value As String) As Long
    On Error Resume Next
    Dim colNum As Long
    Dim worksheetName As String
    
    worksheetName = ActiveSheet.Name
    
    colNum = Application.Match(column_value, ActiveWorkbook.sheets(worksheetName).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        column_number = colNum
    Else
        column_number = 0
    End If
    
End Function

Public Function column_letter(column_value As String) As String
    On Error Resume Next
    Dim colNum As Long
    Dim vArr
    worksheetName = ActiveSheet.Name

    colNum = Application.Match(column_value, ActiveWorkbook.sheets(worksheetName).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        column_letter = Replace(Cells(1, colNum).Address(False, False), "1", "")
    Else
        column_letter = ""                       '
    End If
End Function

Public Function gen_column_number(column_value As String, SHEET_NAME As String) As Long
    On Error Resume Next
    Dim colNum As Long
    Dim worksheetName As String

    colNum = Application.Match(column_value, sheets(SHEET_NAME).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        gen_column_number = colNum
    Else
        gen_column_number = 0
    End If
    
End Function

Public Function this_gen_column_number(column_value As String, SHEET_NAME As String) As Long
    On Error Resume Next
    Dim colNum As Long
    Dim worksheetName As String

    colNum = Application.Match(column_value, ThisWorkbook.sheets(SHEET_NAME).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        this_gen_column_number = colNum
    Else
        this_gen_column_number = 0
    End If
    
End Function

Public Function gen_column_letter(column_value As String, SHEET_NAME As String) As String
    On Error Resume Next
    Dim colNum As Long
    Dim vArr

    colNum = Application.Match(column_value, sheets(SHEET_NAME).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        gen_column_letter = Replace(sheets(SHEET_NAME).Cells(1, colNum).Address(False, False), "1", "")
    Else
        gen_column_letter = ""
    End If
End Function

Public Function data_column_letter(column_value As String) As String
    On Error Resume Next
    Dim colNum As Long
    Dim vArr
    Dim ws_name As String
    ws_name = find_main_data
    
    colNum = Application.Match(column_value, sheets(ws_name).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        data_column_letter = Replace(sheets(ws_name).Cells(1, colNum).Address(False, False), "1", "")
    Else
        data_column_letter = ""
    End If
End Function

Public Function find_uuid_coln() As Long
    On Error Resume Next
    Dim colNum As Long
    Dim worksheetName As String

    colNum = Application.Match("_uuid", sheets(find_main_data).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        find_uuid_coln = colNum
    Else
        find_uuid_coln = 0
    End If
End Function

Public Sub create_sheet(sheet_name_base As String, new_sheet_name As String)
    On Error Resume Next
    sheets.Add(after:=sheets(sheet_name_base)).Name = new_sheet_name
End Sub

Function unmatched_elements(array1 As Variant, array2 As Variant, check_both As Boolean) As Collection
    On Error Resume Next
    Dim arr1() As Variant
    Dim arr2() As Variant
    Dim unmatched As New Collection
    Dim i As Long
        
    With Application
        arr1 = .Transpose(array1)
    End With
    
    With Application
        arr2 = .Transpose(array2)
    End With
    
    ' Find elements in arr1 that are not in arr2
    For i = LBound(arr1) To UBound(arr1)
        If Not is_in_array(arr1(i), arr2) Then
            unmatched.Add arr1(i)
        End If
    Next i
    
    If check_both Then
        ' Find elements in arr2 that are not in arr1
        For i = LBound(arr2) To UBound(arr2)
            If Not is_in_array(arr2(i), arr1) Then
                unmatched.Add arr2(i)
            End If
        Next i
    End If
    
    Set unmatched_elements = unmatched
    
End Function

Function is_in_array(val As Variant, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If val = arr(i) Then
            is_in_array = True
            Exit Function
        End If
    Next i
    is_in_array = False
End Function

Sub clear_filter(ws As Worksheet)
    On Error Resume Next
    Dim filtered_col As Long

    If ws.FilterMode Then
        With ws.AutoFilter
            For filtered_col = 1 To .Filters.count
                If .Filters(filtered_col).On Then
                    ws.AutoFilter.Sort.SortFields.Clear
                    ws.ShowAllData
                End If
            Next filtered_col
        End With
        ws.UsedRange.AutoFilter
    End If
    
End Sub

Sub add_auto_filter()
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    End If
End Sub

Sub remove_auto_filter()
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If
End Sub

Sub clear_active_filter()
    On Error Resume Next
    If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
End Sub

' This function returns a collection of worksheet names in the workbook
Function sheet_list() As Collection
    On Error Resume Next
    
    Dim sheets As Collection
    Set sheets = New Collection
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        sheets.Add ws.Name
    Next ws
    Set sheet_list = sheets
End Function

Function unique_values(rng As Range) As Collection
    On Error Resume Next
    Dim d As Object, c As Range, h, Tmp As String
    Dim unique_collection As New Collection
    
    Set d = CreateObject("scripting.dictionary")
    For Each c In rng
        Tmp = Trim(c.value)
        If Len(Tmp) > 0 Then d(Tmp) = d(Tmp) + 1
    Next c

    For Each h In d.Keys
        'Debug.Print h
        unique_collection.Add CStr(h)
    Next h
    Set unique_values = unique_collection
End Function

Function find_main_data() As String
    On Error Resume Next
    Dim dt As String
    dt = ""
    dt = GetRegistrySetting("ramSetting", "dataReg")

    If dt <> "" Then
        If worksheet_exists(dt) Then
            find_main_data = dt
            Exit Function
        End If
    End If
    
    select_data_form.Show
    find_main_data = GetRegistrySetting("ramSetting", "dataReg")
    
End Function

Function replace_char(str As String)
    On Error Resume Next
    Dim i As Long
    Dim char As Variant
    Dim new_str As String
    Dim char_set As String
    
    new_str = str
    char_set = "!,@,#,%,^,&,*,~,`"
    
    For Each char In Split(char_set, ",")
        new_str = Replace(new_str, char, "_")
    Next

    replace_char = new_str
    
End Function

Public Function alpha_numeric_only(strSource As String) As String
    Dim i As Integer
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 48 To 57, 65 To 90, 97 To 122: 'include 32 if you want to include space
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    
    If Len(strResult) < 15 Then
        alpha_numeric_only = strResult
    Else
        alpha_numeric_only = left(strResult, 15)
    End If
   
End Function

Sub remove_empty_col()
    Dim dt_ws As Worksheet
    Dim i As Long
    Dim last_col As Long
    Dim colle As New Collection
    
    Set dt_ws = sheets(find_main_data)
    
    last_col = dt_ws.Cells(1, Columns.count).End(xlToLeft).Column
    
    For i = 1 To last_col
        If WorksheetFunction.CountA(dt_ws.Columns(i)) = 0 Then
            colle.Add i
        End If
    Next

    For j = colle.count To 1 Step -1
        dt_ws.Columns(colle.item(j)).Delete
    Next j
End Sub

Function no_value(question As String) As Boolean

    Dim dt_ws As Worksheet
    Dim question_col As Long
    Dim last_col As Long

    Set dt_ws = sheets(find_main_data)
    
    question_col = gen_column_number(question, dt_ws.Name)
    
    If question_col = 0 Then
        no_value = True
        Exit Function
    End If
    
    If WorksheetFunction.CountA(dt_ws.Columns(question_col)) = 1 Then
        no_value = True
    Else
        no_value = False
    End If

End Function

Sub create_log_shortcut()
    On Error Resume Next
    Application.OnKey "+^{M}", "show_issue"
End Sub

Sub delete_log_shortcut()
    On Error Resume Next
    Application.OnKey "^{M}"
End Sub

' return the label of main measurement
Function var_label(var As String) As String
    On Error GoTo ErrorHandler
    
    Dim last_row_survey As Long
    Dim v_label As String
    
    last_row_survey = ThisWorkbook.Worksheets("xsurvey").Cells(Rows.count, 1).End(xlUp).Row
    v_label = WorksheetFunction.index(ThisWorkbook.sheets("xsurvey").Range("C2:C" & last_row_survey), _
            WorksheetFunction.Match(var, ThisWorkbook.sheets("xsurvey").Range("B2:B" & last_row_survey), 0))
                
    If v_label = vbNullString Then
        var_label = var
        
    Else
        var_label = v_label
    End If
    Exit Function
                
ErrorHandler:
    var_label = var
    
End Function

' return the label of choice, if not not found return the original choice value
Function choice_label(question As String, choice As String) As String

    On Error GoTo ErrorHandler
    
    Dim ws_sc As Worksheet
    Set ws_sc = ThisWorkbook.sheets("xsurvey_choices")
    Dim last_row_xsurvey_choices As Long
    Dim question_choice As String
    question_choice = question & choice
    
    last_row_xsurvey_choices = ws_sc.Cells(Rows.count, 1).End(xlUp).Row
    choice_label = WorksheetFunction.index(ws_sc.Range("E2:E" & last_row_xsurvey_choices), _
                        WorksheetFunction.Match(question_choice, ws_sc.Range("F2:F" & last_row_xsurvey_choices), 0))

    Exit Function

ErrorHandler:
    choice_label = choice

End Function

Sub extract_choice(str As String)
    On Error Resume Next
    Dim ws As Worksheet
    Dim rng As Range
    
    If ThisWorkbook.Worksheets("xsurvey").Range("A1") = vbNullString Then
        MsgBox "Please import the KOBO tools.    ", vbInformation
        End
    End If
    
    Set ws = ThisWorkbook.sheets("xsurvey_choices")
    ws.Columns("H:K").Clear
    Set rng = ws.Range("A1").CurrentRegion
    
    ws.Cells(1, "H") = "question"
    ws.Cells(1, "K") = "choice"
    ws.Cells(2, "H") = "'=" & str
    
    rng.AdvancedFilter xlFilterCopy, ws.Range("H1:H2"), ws.Range("K1"), True
        
End Sub

Sub remove_NA()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = sheets(Public_module.DATA_SHEET)
    
    ws.Cells.Replace What:="NA", replacement:="", LookAt:=xlWhole, SearchOrder:=xlByColumns, _
            MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False

End Sub

Sub remove_tmp()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.StatusBar = False
    
    If worksheet_exists("keen") Then
        sheets("keen").Visible = xlSheetHidden
        sheets("keen").Delete
    End If
    
    If worksheet_exists("keen2") Then
        sheets("keen2").Visible = xlSheetHidden
        sheets("keen2").Delete
    End If
    
    If worksheet_exists("temp_sheet") Then
        sheets("temp_sheet").Visible = xlSheetHidden
        sheets("temp_sheet").Delete
    End If
    
    If worksheet_exists("redeem") Then
        sheets("redeem").Visible = xlSheetHidden
        sheets("redeem").Delete
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Function check_empty_cells(ws As Worksheet, col_num As Long) As Boolean
    Dim lastRow As Long
    Dim columnToCheck As Integer
    Dim i As Long
    lastRow = ws.Cells(ws.Rows.count, find_uuid_coln).End(xlUp).Row

    For i = 1 To lastRow
        If IsEmpty(ws.Cells(i, col_num).value) Then
            check_empty_cells = True
            Exit Function
        End If
    Next i
    check_empty_cells = False
End Function

Function check_exist_dis_levels() As String
    Dim ws As Worksheet
    Dim dt_ws As Worksheet
    Dim rng As Range
    Dim c As Range
    Dim last_row As Long
    Dim col_number As Long
    Dim check_dis As Boolean
    Dim header_arr() As Variant
    Dim v As Variant
    
    Set ws = sheets("disaggregation_setting")
    Set dt_ws = sheets(find_main_data)
    last_row = ws.Cells(Rows.count, 1).End(xlUp).Row
    Set rng = ws.Range("A2:A" & last_row)
    
    header_arr = dt_ws.Range(dt_ws.Cells(1, 1), dt_ws.Cells(1, 1).End(xlToRight)).Value2
     
    For Each c In rng
        If c <> "ALL" Then
            col_number = gen_column_number(CStr(c), find_main_data)
            If col_number = 0 Then
                Debug.Print CStr(c)
                check_exist_dis_levels = CStr(c)
                Exit Function
            End If
        End If
    Next c
    check_exist_dis_levels = vbNullString
End Function

Function check_null_dis_levels() As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim c As Range
    Dim last_row As Long
    Dim col_number As Long
    Dim check_dis As Boolean
    Set ws = sheets("disaggregation_setting")
    last_row = ws.Cells(Rows.count, 1).End(xlUp).Row
    Set rng = ws.Range("A2:A" & last_row)
    
    For Each c In rng
        If c <> "ALL" Then
            col_number = gen_column_number(CStr(c), find_main_data)
            check_dis = check_empty_cells(sheets(find_main_data), col_number)
            If check_dis Then
                check_null_dis_levels = CStr(c)
                Exit Function
            End If
        End If
    Next c
    check_null_dis_levels = vbNullString
End Function

Function ColumnNumberToLetter(colNumber As Integer) As String
    Dim dividend As Integer
    Dim columnLetter As String
    Dim modulo As Integer
    
    columnLetter = ""
    dividend = colNumber
    
    Do
        modulo = (dividend - 1) Mod 26
        columnLetter = Chr(65 + modulo) & columnLetter
        dividend = (dividend - modulo) \ 26
    Loop While dividend > 0
    
    ColumnNumberToLetter = columnLetter
End Function

Sub ListAllThisSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print ws.Name
    Next ws
End Sub

Sub ListAllSheets()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Debug.Print ws.Name
    Next ws
End Sub

Sub ListNonActiveWorkbooks()
    Dim wb As Workbook
    Dim i As Integer
    
    i = 1
    For Each wb In Workbooks
        If Not wb Is ActiveWorkbook Then
            Debug.Print wb.Name
            i = i + 1
        End If
    Next wb
End Sub

Sub ExportAllModulesAndForms()
    Dim VBComp As Object
    Dim VBProj As Object
    Dim FileName As String
    Dim filePath As String
    
    filePath = ActiveWorkbook.path
    filePath = filePath & "\codes\"
    
    Call CheckAndCreateDirectory(filePath)
    
    Set VBProj = ThisWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        Debug.Print VBComp.Type
        If VBComp.Type = 1 Then 'Standard module
            FileName = filePath & VBComp.Name & ".bas"
            VBComp.Export FileName
        ElseIf VBComp.Type = 2 Then 'Class module
            FileName = filePath & VBComp.Name & ".cls"
            VBComp.Export FileName
        ElseIf VBComp.Type = 3 Then 'UserForm
            FileName = filePath & VBComp.Name & ".frm"
            VBComp.Export FileName
        ElseIf VBComp.Type = 100 Then
            FileName = filePath & VBComp.Name & ".cls"
            VBComp.Export FileName
        End If
    Next VBComp
    MsgBox "All modules and forms exported successfully!", vbInformation
End Sub

Sub initialTheApp()
    Dim strFolder As String
    Dim strFile As String
    Dim VBProj As Object
    Dim VBComp As Object
    Dim FileName As String
    
    Dim filePath As String
    filePath = ThisWorkbook.path
    strFolder = ThisWorkbook.path & "\codes\"
    If Dir(strFolder, vbDirectory) = "" Then
        MsgBox "Folder not found!", vbExclamation
        Exit Sub
    End If
    
    Set VBProj = ThisWorkbook.VBProject

    strFile = Dir(strFolder & "*.*")
    Do While strFile <> ""
        If Right(strFile, 4) = ".bas" Or Right(strFile, 4) = ".cls" Or Right(strFile, 4) = ".frm" Then
            VBProj.VBComponents.Import strFolder & strFile
        End If
        strFile = Dir
    Loop
    Call CreateMainSheets
    MsgBox "All modules and forms imported successfully!", vbInformation
End Sub

Sub CreateMainSheets()
    Dim ws As Worksheet
    
    Set ws = sheets(1)
    ws.Name = "xsurvey"
    
    Set ws = ThisWorkbook.sheets.Add(after:=ThisWorkbook.sheets(ThisWorkbook.sheets.count))
    ws.Name = "xchoices"
    
    Set ws = ThisWorkbook.sheets.Add(after:=ThisWorkbook.sheets(ThisWorkbook.sheets.count))
    ws.Name = "xsurvey_choices"

    Set ws = ThisWorkbook.sheets.Add(after:=ThisWorkbook.sheets(ThisWorkbook.sheets.count))
    ws.Name = "xlogical_checks"
End Sub

Sub CheckAndCreateDirectory(pathDir As String)
    Dim myPath As String
    Dim checkDir As String
    
    checkDir = Dir(pathDir, vbDirectory)
    
    If checkDir = "" Then
        MkDir pathDir
    End If
End Sub

Private Function show_sheet()
    sheets("temp_sheet").Visible = True
End Function

Sub set_basic_config()
    Application.ReferenceStyle = xlA1
    Application.Calculation = xlAutomatic
End Sub

Sub SaveAsAddIn()
    Dim savePath As String
    savePath = Application.GetSaveAsFilename(FileFilter:="Excel Add-In (*.xlam), *.xlam")
    
    ' Check if the user canceled the save dialog
    If savePath <> "False" Then
        ' Save the current workbook as an add-in
        ThisWorkbook.SaveAs FileName:=savePath, FileFormat:=xlAddIn
    End If
End Sub

