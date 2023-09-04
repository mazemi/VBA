Attribute VB_Name = "logical_check_module"
Option Explicit

Sub auto_check()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim plan_ws As Worksheet
    Dim rng As Range
    Dim filtered_rng As Range
    Dim cr_rng As Range
    Dim last_dt As Long
    Dim plan_row As Long
    Dim condition1 As String
    Dim condition2 As String
    Dim condition1_col As String
    Dim condition2_col As String
    Dim i As Long
    Dim col_n1 As Long
    Dim col_n2 As Long
    Dim uuid_coln As Long
    
    If ThisWorkbook.sheets("xlogical_checks").Range("A1") = "" Then
        MsgBox "There is not any logical check!   ", vbInformation
        Exit Sub
    End If
    
    Set plan_ws = ThisWorkbook.sheets("xlogical_checks")
    Set ws = sheets(find_main_data)
    
    If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
        ws.ShowAllData
    End If
        
    Set rng = ws.Range("A1").CurrentRegion
    
    uuid_coln = gen_column_number("_uuid", find_main_data)
    
    If uuid_coln = 0 Then
        MsgBox "The '_uuid' column dose not exist!  ", vbInformation
        Exit Sub
    End If
    
    Call remove_empty_col
    
    last_dt = ws.Cells(rows.count, uuid_coln).End(xlUp).row
    
    plan_row = plan_ws.Cells(rows.count, 1).End(xlUp).row
    
    plan_ws.Range("M1") = Null
    plan_ws.Range("M2") = Null
    plan_ws.Range("N1") = Null
    plan_ws.Range("N2") = Null
    plan_ws.Range("N3") = Null
        
    ws.Activate
    
    For i = 1 To plan_row
        Application.ScreenUpdating = False
        If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
            ws.ShowAllData
        End If
        
        Public_module.PATTERN_CHECK_ACTION = True
        Public_module.ISSUE_TEXT = plan_ws.Cells(i, "H")
        
        ' numeric conversion
        condition1 = plan_ws.Cells(i, "C")
        condition2 = plan_ws.Cells(i, "G")
        
        Dim rng1 As Range, cel1 As Range
        Dim rng2 As Range, cel2 As Range
        
        If IsNumeric(condition1) Then
            condition1_col = gen_column_letter(plan_ws.Cells(i, "A"), find_main_data)
            If condition1_col <> vbNullString Then
                Set rng1 = ws.Range(condition1_col & "2:" & condition1_col & last_dt)
                For Each cel1 In rng1.Cells
                    If Len(cel1.value) <> 0 And IsNumeric(cel1.value) Then
                        cel1.value = CSng(cel1.value)
                    End If
                Next cel1
            End If
        End If
        
        If IsNumeric(condition2) Then
            condition2_col = gen_column_letter(plan_ws.Cells(i, "E"), find_main_data)
            If condition2_col <> vbNullString Then
                Set rng2 = ws.Range(condition2_col & "2:" & condition2_col & last_dt)
                For Each cel2 In rng2.Cells
                    If Len(cel2.value) <> 0 And IsNumeric(cel2.value) Then
                        cel2.value = CSng(cel2.value)
                    End If
                Next cel2
            End If
        End If
        
        ' case 1
        If plan_ws.Cells(i, "D") = "" Then
            col_n1 = gen_column_number(plan_ws.Cells(i, "A"), find_main_data)
            
            If col_n1 = 0 Then
                GoTo resumeLoop
            End If
            
            rng.Sort key1:=rng.Cells(1, col_n1), Order1:=xlAscending, Header:=xlYes
            rng.AutoFilter col_n1, give_operator("B" & i) & plan_ws.Cells(i, "C")
            
            If Not IsError(ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select) Then
                Call auto_pattern_check(number_to_letter(col_n1, ws))
            End If
            
        ' case 2
        ElseIf plan_ws.Cells(i, "D") = "and" And plan_ws.Cells(i, "A") = plan_ws.Cells(i, "E") Then
            col_n1 = gen_column_number(plan_ws.Cells(i, "A"), find_main_data)
            
            If col_n1 = 0 Then
                GoTo resumeLoop
            End If
            
            rng.Sort key1:=rng.Cells(1, col_n1), Order1:=xlAscending, Header:=xlYes
            rng.AutoFilter col_n1, give_operator("B" & i) & plan_ws.Cells(i, "C"), xlAnd, _
                give_operator("F" & i) & plan_ws.Cells(i, "G")
                
            If Not IsError(ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select) Then
                Call auto_pattern_check(number_to_letter(col_n1, ws))
            End If
            
        ' case 3
        ElseIf plan_ws.Cells(i, "D") = "or" And plan_ws.Cells(i, "A") = plan_ws.Cells(i, "E") Then
            
            col_n1 = gen_column_number(plan_ws.Cells(i, "A"), find_main_data)

            If col_n1 = 0 Then
                GoTo resumeLoop
            End If
  
            rng.Sort key1:=rng.Cells(1, col_n1), Order1:=xlAscending, Header:=xlYes
            rng.AutoFilter col_n1, give_operator("B" & i) & plan_ws.Cells(i, "C")
            
            If Not IsError(ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select) Then
                Call auto_pattern_check(number_to_letter(col_n1, ws))
            End If
            
            If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
                ws.ShowAllData
            End If
        
            rng.AutoFilter col_n1, give_operator("F" & i) & plan_ws.Cells(i, "G")
            
            If Not IsError(ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select) Then
                Call auto_pattern_check(number_to_letter(col_n1, ws))
            End If
            
        ' use advancefilter
        ' case 4
        ElseIf plan_ws.Cells(i, "D") = "and" And plan_ws.Cells(i, 1) <> plan_ws.Cells(i, 4) Then
            col_n1 = gen_column_number(plan_ws.Cells(i, "A"), find_main_data)
            col_n2 = gen_column_number(plan_ws.Cells(i, "E"), find_main_data)
            
            If col_n1 = 0 Or col_n2 = 0 Then
                GoTo resumeLoop
            End If
            
            plan_ws.Range("M1") = plan_ws.Cells(i, "A")
            plan_ws.Range("M2") = give_operator("B" & i) & plan_ws.Cells(i, "C")
            plan_ws.Range("N1") = plan_ws.Cells(i, "E")
            plan_ws.Range("N2") = give_operator("F" & i) & plan_ws.Cells(i, "G")
            
            Set cr_rng = plan_ws.Range("M1").CurrentRegion
            rng.AdvancedFilter xlFilterInPlace, cr_rng

            rng.Sort key1:=rng.Cells(1, col_n1), Order1:=xlAscending, Header:=xlYes
            If Not IsError(ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select) Then
                Call auto_pattern_check(number_to_letter(col_n1, ws))
            End If
            
            rng.Sort key1:=rng.Cells(1, col_n2), Order1:=xlAscending, Header:=xlYes
            If Not IsError(ws.Range(ws.Cells(2, col_n2), ws.Cells(last_dt, col_n2)).SpecialCells(xlCellTypeVisible).Select) Then
                Call auto_pattern_check(number_to_letter(col_n2, ws))
            End If
            
        ' use advancefilter
        ' case 5
        ElseIf plan_ws.Cells(i, "D") = "or" And plan_ws.Cells(i, 1) <> plan_ws.Cells(i, 4) Then
            col_n1 = gen_column_number(plan_ws.Cells(i, "A"), find_main_data)
            col_n2 = gen_column_number(plan_ws.Cells(i, "E"), find_main_data)
            
            If col_n1 = 0 Or col_n2 = 0 Then
                GoTo resumeLoop
            End If
            
            plan_ws.Range("M1") = plan_ws.Cells(i, "A")
            plan_ws.Range("M2") = give_operator("B" & i) & plan_ws.Cells(i, "C")
            plan_ws.Range("N1") = plan_ws.Cells(i, "E")
            plan_ws.Range("N3") = give_operator("F" & i) & plan_ws.Cells(i, "G")
            
            Set cr_rng = plan_ws.Range("M1").CurrentRegion
            rng.AdvancedFilter xlFilterInPlace, cr_rng
            
            rng.Sort key1:=rng.Cells(1, col_n1), Order1:=xlAscending, Header:=xlYes
            If Not IsError(ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select) Then
                Call auto_pattern_check(number_to_letter(col_n1, ws))
            End If
            
            rng.Sort key1:=rng.Cells(1, col_n2), Order1:=xlAscending, Header:=xlYes
            If Not IsError(ws.Range(ws.Cells(2, col_n2), ws.Cells(last_dt, col_n2)).SpecialCells(xlCellTypeVisible).Select) Then
                Call auto_pattern_check(number_to_letter(col_n2, ws))
            End If
            
        End If

        plan_ws.Range("M1") = Null
        plan_ws.Range("M2") = Null
        plan_ws.Range("N1") = Null
        plan_ws.Range("N2") = Null
        plan_ws.Range("N3") = Null
            
resumeLoop:

    Next
    
    If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
        ws.ShowAllData
    End If
    
    Application.ScreenUpdating = True
    Debug.Print "done"
End Sub

Sub single_check(p_row As Long)
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim plan_ws As Worksheet
    Dim rng As Range
    Dim filtered_rng As Range
    Dim cr_rng As Range
    Dim last_dt As Long
    Dim plan_row As Long
    Dim condition1 As String
    Dim condition2 As String
    Dim condition1_col As String
    Dim condition2_col As String
    Dim i As Long
    Dim col_n1 As Long
    Dim col_n2 As Long
    Dim uuid_coln As Long
    
    If ThisWorkbook.sheets("xlogical_checks").Range("A1") = "" Then
        MsgBox "There is not any logical check!   ", vbInformation
        Exit Sub
    End If
    
    Set plan_ws = ThisWorkbook.sheets("xlogical_checks")
    Set ws = sheets(find_main_data)
    
    If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
        ws.ShowAllData
    End If
        
    Set rng = ws.Range("A1").CurrentRegion
    
    uuid_coln = gen_column_number("_uuid", find_main_data)
    
    If uuid_coln = 0 Then
        MsgBox "The '_uuid' column dose not exist!  ", vbInformation
        Exit Sub
    End If
    
    Call remove_empty_col
    
    last_dt = ws.Cells(rows.count, uuid_coln).End(xlUp).row
    
    plan_row = plan_ws.Cells(rows.count, 1).End(xlUp).row
    
    ws.Activate
    
    Application.ScreenUpdating = False
    If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
        ws.ShowAllData
    End If
    
    Public_module.PATTERN_CHECK_ACTION = True
    Public_module.ISSUE_TEXT = plan_ws.Cells(p_row, "H")
    
    ' numeric conversion
    condition1 = plan_ws.Cells(p_row, "C")
    condition2 = plan_ws.Cells(p_row, "G")
    
    Dim rng1 As Range, cel1 As Range
    Dim rng2 As Range, cel2 As Range
    
    If IsNumeric(condition1) Then
        condition1_col = gen_column_letter(plan_ws.Cells(p_row, "A"), find_main_data)
        If condition1_col <> vbNullString Then
            Set rng1 = ws.Range(condition1_col & "2:" & condition1_col & last_dt)
            For Each cel1 In rng1.Cells
                If Len(cel1.value) <> 0 And IsNumeric(cel1.value) Then
                    cel1.value = CSng(cel1.value)
                End If
            Next cel1
        End If
    End If
    
    If IsNumeric(condition2) Then
        condition2_col = gen_column_letter(plan_ws.Cells(p_row, "E"), find_main_data)
        If condition2_col <> vbNullString Then
            Set rng2 = ws.Range(condition2_col & "2:" & condition2_col & last_dt)
            For Each cel2 In rng2.Cells
                If Len(cel2.value) <> 0 And IsNumeric(cel2.value) Then
                    cel2.value = CSng(cel2.value)
                End If
            Next cel2
        End If
    End If
    
    ' case 1
    If plan_ws.Cells(p_row, "D") = "" Then
        col_n1 = gen_column_number(plan_ws.Cells(p_row, "A"), find_main_data)
        
        If col_n1 = 0 Then
            GoTo resumeLoop
        End If
        
        rng.Sort key1:=rng.Cells(1, col_n1), Order1:=xlAscending, Header:=xlYes
        rng.AutoFilter col_n1, give_operator("B" & p_row) & plan_ws.Cells(p_row, "C")
        
'        If count_rows > 1 Then
'            ws.Cells(2, col_n1).Select
'            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
            ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select
'            Call pattern_check(True)
'        End If

    ' case 2
    ElseIf plan_ws.Cells(p_row, "D") = "and" And plan_ws.Cells(p_row, "A") = plan_ws.Cells(p_row, "E") Then
        col_n1 = gen_column_number(plan_ws.Cells(p_row, "A"), find_main_data)
        
        If col_n1 = 0 Then
            GoTo resumeLoop
        End If
        
        rng.Sort key1:=rng.Cells(1, col_n1), Order1:=xlAscending, Header:=xlYes
        rng.AutoFilter col_n1, give_operator("B" & p_row) & plan_ws.Cells(p_row, "C"), xlAnd, _
            give_operator("F" & p_row) & plan_ws.Cells(p_row, "G")
            
'        If count_rows > 1 Then
'            ws.Cells(2, col_n1).Select
'            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
            ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select
'            Call pattern_check(True)
'        End If
        
    ' case 3
    ElseIf plan_ws.Cells(p_row, "D") = "or" And plan_ws.Cells(p_row, "A") = plan_ws.Cells(p_row, "E") Then
        
        col_n1 = gen_column_number(plan_ws.Cells(p_row, "A"), find_main_data)

        If col_n1 = 0 Then
            GoTo resumeLoop
        End If

        rng.Sort key1:=rng.Cells(1, col_n1), Order1:=xlAscending, Header:=xlYes
        rng.AutoFilter col_n1, give_operator("B" & p_row) & plan_ws.Cells(p_row, "C")
        
'        If count_rows > 1 Then
'            ws.Cells(2, col_n1).Select
'            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
            ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select
'            Call pattern_check(True)
'        End If
        
        If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
            ws.ShowAllData
        End If
    
        rng.AutoFilter col_n1, give_operator("F" & p_row) & plan_ws.Cells(p_row, "G")
        
'        If count_rows > 1 Then
'            ws.Cells(2, col_n1).Select
'            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
            ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select
'            Call pattern_check(True)
'        End If
        
    ' use advancefilter
    ' case 4
    ElseIf plan_ws.Cells(p_row, "D") = "and" And plan_ws.Cells(p_row, 1) <> plan_ws.Cells(p_row, 4) Then
        col_n1 = gen_column_number(plan_ws.Cells(p_row, "A"), find_main_data)
        col_n2 = gen_column_number(plan_ws.Cells(p_row, "E"), find_main_data)
        
        If col_n1 = 0 Or col_n2 = 0 Then
            GoTo resumeLoop
        End If
        
        plan_ws.Range("M1") = plan_ws.Cells(p_row, "A")
        plan_ws.Range("M2") = give_operator("B" & p_row) & plan_ws.Cells(p_row, "C")
        plan_ws.Range("N1") = plan_ws.Cells(p_row, "E")
        plan_ws.Range("N2") = give_operator("F" & p_row) & plan_ws.Cells(p_row, "G")
        
        Set cr_rng = plan_ws.Range("M1").CurrentRegion
        rng.AdvancedFilter xlFilterInPlace, cr_rng
        
'        If count_rows > 1 Then
            rng.Sort key1:=rng.Cells(1, col_n1), Order1:=xlAscending, Header:=xlYes
'            ws.Cells(2, col_n1).Select
'            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
            ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select
'            Call pattern_check(True)
            
            rng.Sort key1:=rng.Cells(1, col_n2), Order1:=xlAscending, Header:=xlYes
'            ws.Cells(2, col_n2).Select
'            ws.Range(ws.Cells(2, col_n2), Selection.End(xlDown)).Select
            ws.Range(ws.Cells(2, col_n2), ws.Cells(last_dt, col_n2)).SpecialCells(xlCellTypeVisible).Select
'            Call pattern_check(True)
'        End If
        
    ' use advancefilter
    ' case 5
    ElseIf plan_ws.Cells(p_row, "D") = "or" And plan_ws.Cells(p_row, 1) <> plan_ws.Cells(p_row, 4) Then
        col_n1 = gen_column_number(plan_ws.Cells(p_row, "A"), find_main_data)
        col_n2 = gen_column_number(plan_ws.Cells(p_row, "E"), find_main_data)
        
        If col_n1 = 0 Or col_n2 = 0 Then
            GoTo resumeLoop
        End If
        
        plan_ws.Range("M1") = plan_ws.Cells(p_row, "A")
        plan_ws.Range("M2") = give_operator("B" & p_row) & plan_ws.Cells(p_row, "C")
        plan_ws.Range("N1") = plan_ws.Cells(p_row, "E")
        plan_ws.Range("N3") = give_operator("F" & p_row) & plan_ws.Cells(p_row, "G")
        
        Set cr_rng = plan_ws.Range("M1").CurrentRegion
        rng.AdvancedFilter xlFilterInPlace, cr_rng
        
'        If count_rows > 1 Then
            rng.Sort key1:=rng.Cells(1, col_n1), Order1:=xlAscending, Header:=xlYes
'            ws.Cells(2, col_n1).Select
'            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
            ws.Range(ws.Cells(2, col_n1), ws.Cells(last_dt, col_n1)).SpecialCells(xlCellTypeVisible).Select
'            Call pattern_check(True)
            
            rng.Sort key1:=rng.Cells(1, col_n2), Order1:=xlAscending, Header:=xlYes
'            ws.Cells(2, col_n2).Select
'            ws.Range(ws.Cells(2, col_n2), Selection.End(xlDown)).Select
            ws.Range(ws.Cells(2, col_n2), ws.Cells(last_dt, col_n2)).SpecialCells(xlCellTypeVisible).Select
'            Call pattern_check(True)
'        End If
        
    End If

    plan_ws.Range("M1") = Null
    plan_ws.Range("M2") = Null
    plan_ws.Range("N1") = Null
    plan_ws.Range("N2") = Null
    plan_ws.Range("N3") = Null
            
resumeLoop:

    Debug.Print col_n1, col_n2
    ws.Activate
    ActiveWindow.ScrollRow = 1
    If col_n1 > 0 And col_n2 = 0 Then
        If col_n1 > 1 Then
            ActiveWindow.ScrollColumn = col_n1 - 1
        Else
            ActiveWindow.ScrollColumn = col_n1
        End If
    ElseIf col_n2 > 0 Then
        If col_n2 > 1 Then
            ActiveWindow.ScrollColumn = col_n2 - 1
        Else
           ActiveWindow.ScrollColumn = col_n2
        End If
    End If
    
    Application.ScreenUpdating = True
End Sub

Private Function give_operator(str As String) As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets("xlogical_checks")
    
    Select Case ws.Range(str).value
        Case "is equal"
            give_operator = ""
        Case "is not equal"
            give_operator = "<>"
        Case "is empty"
            give_operator = "="
        Case "is not empty"
            give_operator = "<>"
        Case "is greater than"
            give_operator = ">"
        Case "is greater than or equal"
            give_operator = ">="
        Case "is less than"
            give_operator = "<"
        Case "is less than or equal"
            give_operator = "<="
        Case Else
            give_operator = vbNullString
    End Select

End Function

Function count_rows() As Long
    On Error GoTo errHandler
    Dim ws As Worksheet
    Dim uuid_col As Long
    Dim rows_n As Long
    uuid_col = gen_column_number("_uuid", find_main_data)
    Set ws = sheets(find_main_data)
    rows_n = ws.AutoFilter.Range.columns(uuid_col).SpecialCells(xlCellTypeVisible).Cells.count
    Debug.Print rows_n
    count_rows = rows_n
    Exit Function
    
errHandler:
    count_rows = 0

End Function

Sub import_plan()
    On Error Resume Next
    Dim ws As Worksheet, strFile As String
    
    strFile = Application.GetOpenFilename("Plan Files (*.plan),*.plan", , "Please select a cleaning plan...")
    
    If strFile = "False" Then
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.sheets("xlogical_checks")
    ws.Cells.Clear
    
    With ws.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh
    End With
End Sub

Sub export_plan()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbDest As Workbook
    Dim fName As String
    Dim last_row  As Long
    Dim path As String
    
    If ThisWorkbook.sheets("xlogical_checks").Range("A1") = vbNullString Then
        MsgBox "The logical check dose not exist!     ", vbInformation
        Exit Sub
    End If
    Application.DisplayAlerts = False
    
    Set wbSource = ThisWorkbook
    Set wsSource = ThisWorkbook.sheets("xlogical_checks")
    last_row = wsSource.Cells(rows.count, 1).End(xlUp).row

    path = Application.GetSaveAsFilename( _
           FileFilter:="Plan Files (*.plan), *.plan", _
           title:="Save the cleaning plan", _
           InitialFileName:="logical_ckeck")
             
    If path = "" Then End
    
    Workbooks.Add
    
    With ActiveWorkbook
        wsSource.Range("A1:H" & last_row).Copy .sheets(1).Range("A1")
        .SaveAs path, FileFormat:=xlCurrentPlatformText
        .Close True
    End With
    
    Application.DisplayAlerts = True
    
End Sub





