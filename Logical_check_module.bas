Attribute VB_Name = "logical_check_module"
Option Explicit

Sub single_check(r As Long)
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim plan_ws As Worksheet
    Dim rng As Range
    Dim filtered_rng1 As Range
    Dim filtered_rng2 As Range
    Dim cr_rng As Range
    Dim new_row As Long
    Dim plan_row As Long
    Dim condition1 As String
    Dim condition2 As String
    Dim condition1_col As String
    Dim condition2_col As String
    Dim col_n1 As Long
    Dim col_n2 As Long
    Dim uuid_coln As Long
    
    If ThisWorkbook.sheets("logical_checks").Range("A1") = "" Then
        MsgBox "There is not any logical check!   ", vbInformation
        Exit Sub
    End If
    
    Set plan_ws = ThisWorkbook.sheets("logical_checks")
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
    
    col_n1 = 0
    col_n2 = 0
    new_row = ws.Cells(rows.count, uuid_coln).End(xlUp).row
    
    plan_row = plan_ws.Cells(rows.count, 1).End(xlUp).row
    
    ws.Activate
    
    If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
        ws.ShowAllData
    End If
    
    Public_module.PATTERN_CHECK_ACTION = True
    Public_module.ISSUE_TEXT = plan_ws.Cells(r, 6)
    
    ' numeric conversion
    condition1 = remove_operator(plan_ws.Cells(r, 2))
    condition2 = remove_operator(plan_ws.Cells(r, 5))
    
    Dim rng1 As Range, cel1 As Range
    Dim rng2 As Range, cel2 As Range
    
    If IsNumeric(condition1) Then
        condition1_col = gen_column_letter(plan_ws.Cells(r, 1), find_main_data)
        If condition1_col > 0 Then
            Set rng1 = ws.Range(condition1_col & "2:" & condition1_col & ws.Cells(rows.count, uuid_coln).End(xlUp).row)
            For Each cel1 In rng1.Cells
                If Len(cel1.value) <> 0 Then
                    cel1.value = CSng(cel1.value)
                End If
            Next cel1
        End If
    End If
    
    If IsNumeric(condition2) Then
        condition2_col = gen_column_letter(plan_ws.Cells(r, 4), find_main_data)
        If condition2_col > 0 Then
            Set rng2 = ws.Range(condition2_col & "2:" & condition2_col & ws.Cells(rows.count, uuid_coln).End(xlUp).row)
            For Each cel2 In rng2.Cells
                If Len(cel2.value) <> 0 Then
                    cel2.value = CSng(cel2.value)
                End If
            Next cel2
        End If
    End If
    
    ' add to logbook based on the case:
    If plan_ws.Cells(r, 3) = "" Then
        col_n1 = column_number(plan_ws.Cells(r, 1))
        rng.AutoFilter col_n1, plan_ws.Cells(r, 2)
        
        If count_rows > 1 Then
            ws.Cells(2, col_n1).Select
            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
        End If
        
    ElseIf plan_ws.Cells(r, 3) = "and" And plan_ws.Cells(r, 1) = plan_ws.Cells(r, 4) Then
        col_n1 = column_number(plan_ws.Cells(r, 1))
        rng.AutoFilter col_n1, plan_ws.Cells(r, 2), xlAnd, plan_ws.Cells(r, 5)
        
        If count_rows > 1 Then
            ws.Cells(2, col_n1).Select
            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
        End If
        
    ElseIf plan_ws.Cells(r, 3) = "or" And plan_ws.Cells(r, 1) = plan_ws.Cells(r, 4) Then
        col_n1 = column_number(plan_ws.Cells(r, 1))
        rng.AutoFilter col_n1, plan_ws.Cells(r, 2), xlOr, plan_ws.Cells(r, 5)
        
        If count_rows > 1 Then
            ws.Cells(2, col_n1).Select
            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
        End If
        
        ' use advancefilter
    ElseIf plan_ws.Cells(r, 3) = "and" And plan_ws.Cells(r, 1) <> plan_ws.Cells(r, 4) Then
        col_n1 = column_number(plan_ws.Cells(r, 1))
        col_n2 = column_number(plan_ws.Cells(r, 4))
        
        plan_ws.Range("M1") = plan_ws.Cells(r, 1)
        plan_ws.Range("M2") = plan_ws.Cells(r, 2)
        plan_ws.Range("N1") = plan_ws.Cells(r, 4)
        plan_ws.Range("N2") = plan_ws.Cells(r, 5)
        
        Set cr_rng = plan_ws.Range("M1").CurrentRegion
        rng.AdvancedFilter xlFilterInPlace, cr_rng
        
'        If count_rows > 1 Then
        
            Set filtered_rng1 = ws.Range(ws.Cells(2, col_n1), ws.Range(ws.Cells(rows.count, col_n1))).End(xlUp)
            ' Check if r is only 1 cell
            If filtered_rng1.count = 1 Then
                filtered_rng1.Select
            Else
                Set filtered_rng1 = filtered_rng1.SpecialCells(xlCellTypeVisible)
                filtered_rng1.Select
            End If
            
            Set filtered_rng2 = ws.Range(ws.Cells(2, col_n2), ws.Range(ws.Cells(rows.count, col_n2))).End(xlUp)
            ' Check if r is only 1 cell
            If filtered_rng2.count = 1 Then
                filtered_rng2.Select
            Else
                Set filtered_rng2 = filtered_rng2.SpecialCells(xlCellTypeVisible)
                filtered_rng2.Select
            End If
        
'            ws.Cells(2, col_n1).Select
'            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
'
'            ws.Cells(2, col_n2).Select
'            ws.Range(ws.Cells(2, col_n2), Selection.End(xlDown)).Select
'        End If
        
        ' use advancefilter
    ElseIf plan_ws.Cells(r, 3) = "or" And plan_ws.Cells(r, 1) <> plan_ws.Cells(r, 4) Then
        col_n1 = column_number(plan_ws.Cells(r, 1))
        col_n2 = column_number(plan_ws.Cells(r, 4))
        
        plan_ws.Range("M1") = plan_ws.Cells(r, 1)
        plan_ws.Range("M2") = plan_ws.Cells(r, 2)
        plan_ws.Range("N1") = plan_ws.Cells(r, 4)
        plan_ws.Range("N3") = plan_ws.Cells(r, 5)
        
        Set cr_rng = plan_ws.Range("M1").CurrentRegion
        rng.AdvancedFilter xlFilterInPlace, cr_rng
        
'        If count_rows > 1 Then
            Set filtered_rng1 = ws.Range(ws.Cells(2, col_n1), ws.Range(ws.Cells(rows.count, col_n1))).End(xlUp)
            ' Check if r is only 1 cell
            If filtered_rng1.count = 1 Then
                filtered_rng1.Select
            Else
                Set filtered_rng1 = filtered_rng1.SpecialCells(xlCellTypeVisible)
                filtered_rng1.Select
            End If
            
            Set filtered_rng2 = ws.Range(ws.Cells(2, col_n2), ws.Range(ws.Cells(rows.count, col_n2))).End(xlUp)
            ' Check if r is only 1 cell
            If filtered_rng2.count = 1 Then
                filtered_rng2.Select
            Else
                Set filtered_rng2 = filtered_rng2.SpecialCells(xlCellTypeVisible)
                filtered_rng2.Select
            End If
        
'            ws.Cells(2, col_n1).Select
'            ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
'
'            ws.Cells(2, col_n2).Select
'            ws.Range(ws.Cells(2, col_n2), Selection.End(xlDown)).Select
'        End If
        
    End If
    
    plan_ws.Range("M1") = Null
    plan_ws.Range("M2") = Null
    plan_ws.Range("N1") = Null
    plan_ws.Range("N2") = Null
    plan_ws.Range("N3") = Null
    
    If col_n1 > 0 And col_n2 = 0 Then
        If col_n1 > 1 Then
            ActiveWindow.ScrollColumn = col_n1 - 1
        Else
            ActiveWindow.ScrollColumn = col_n1
        End If
        ActiveWindow.ScrollRow = 1
    ElseIf col_n2 > 0 Then
        If col_n2 > 1 Then
            ActiveWindow.ScrollColumn = col_n2 - 1
        Else
           ActiveWindow.ScrollColumn = col_n2
        End If
        ActiveWindow.ScrollRow = 1
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Sub auto_check()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim plan_ws As Worksheet
    Dim rng As Range
    Dim filtered_rng As Range
    Dim cr_rng As Range
    Dim new_row As Long
    Dim plan_row As Long
    Dim condition1 As String
    Dim condition2 As String
    Dim condition1_col As String
    Dim condition2_col As String
    Dim i As Long
    Dim col_n1 As Long
    Dim col_n2 As Long
    Dim uuid_coln As Long
    
    If ThisWorkbook.sheets("logical_checks").Range("A1") = "" Then
        MsgBox "There is not any logical check!   ", vbInformation
        Exit Sub
    End If
    
    Set plan_ws = ThisWorkbook.sheets("logical_checks")
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
    
    new_row = ws.Cells(rows.count, uuid_coln).End(xlUp).row
    
    plan_row = plan_ws.Cells(rows.count, 1).End(xlUp).row
    
    ws.Activate
    
    For i = 1 To plan_row
        Application.ScreenUpdating = False
        If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
            ws.ShowAllData
        End If
        
        Public_module.PATTERN_CHECK_ACTION = True
        Public_module.ISSUE_TEXT = plan_ws.Cells(i, 6)
        
        ' numeric conversion
        condition1 = remove_operator(plan_ws.Cells(i, 2))
        condition2 = remove_operator(plan_ws.Cells(i, 5))
        
        Dim rng1 As Range, cel1 As Range
        Dim rng2 As Range, cel2 As Range
        
        If IsNumeric(condition1) Then
            condition1_col = gen_column_letter(plan_ws.Cells(i, 1), find_main_data)
            If condition1_col > 0 Then
                Set rng1 = ws.Range(condition1_col & "2:" & condition1_col & ws.Cells(rows.count, uuid_coln).End(xlUp).row)
                For Each cel1 In rng1.Cells
                    If Len(cel1.value) <> 0 Then
                        cel1.value = CSng(cel1.value)
                    End If
                Next cel1
            End If
        End If
        
        If IsNumeric(condition2) Then
            condition2_col = gen_column_letter(plan_ws.Cells(i, 4), find_main_data)
            If condition2_col > 0 Then
                Set rng2 = ws.Range(condition2_col & "2:" & condition2_col & ws.Cells(rows.count, uuid_coln).End(xlUp).row)
                For Each cel2 In rng2.Cells
                    If Len(cel2.value) <> 0 Then
                        cel2.value = CSng(cel2.value)
                    End If
                Next cel2
            End If
        End If
        
        If plan_ws.Cells(i, 3) = "" Then
            col_n1 = column_number(plan_ws.Cells(i, 1))
            rng.AutoFilter col_n1, plan_ws.Cells(i, 2)
            
            If count_rows > 1 Then
                ws.Cells(2, col_n1).Select
                ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
                Call pattern_check(True)
            End If
            
        ElseIf plan_ws.Cells(i, 3) = "and" And plan_ws.Cells(i, 1) = plan_ws.Cells(i, 4) Then
            col_n1 = column_number(plan_ws.Cells(i, 1))
            rng.AutoFilter col_n1, plan_ws.Cells(i, 2), xlAnd, plan_ws.Cells(i, 5)
            
            If count_rows > 1 Then
                ws.Cells(2, col_n1).Select
                ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
                Call pattern_check(True)
            End If
            
        ElseIf plan_ws.Cells(i, 3) = "or" And plan_ws.Cells(i, 1) = plan_ws.Cells(i, 4) Then
            col_n1 = column_number(plan_ws.Cells(i, 1))
            rng.AutoFilter col_n1, plan_ws.Cells(i, 2), xlOr, plan_ws.Cells(i, 5)
            
            If count_rows > 1 Then
                ws.Cells(2, col_n1).Select
                ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
                Call pattern_check(True)
            End If
            
            ' use advancefilter
        ElseIf plan_ws.Cells(i, 3) = "and" And plan_ws.Cells(i, 1) <> plan_ws.Cells(i, 4) Then
            col_n1 = column_number(plan_ws.Cells(i, 1))
            col_n2 = column_number(plan_ws.Cells(i, 4))
            
            plan_ws.Range("M1") = plan_ws.Cells(i, 1)
            plan_ws.Range("M2") = plan_ws.Cells(i, 2)
            plan_ws.Range("N1") = plan_ws.Cells(i, 4)
            plan_ws.Range("N2") = plan_ws.Cells(i, 5)
            
            Set cr_rng = plan_ws.Range("M1").CurrentRegion
            rng.AdvancedFilter xlFilterInPlace, cr_rng
            
            If count_rows > 1 Then
                ws.Cells(2, col_n1).Select
                ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
                Call pattern_check(True)
                
                ws.Cells(2, col_n2).Select
                ws.Range(ws.Cells(2, col_n2), Selection.End(xlDown)).Select
                Call pattern_check(True)
            End If
            
            ' use advancefilter
        ElseIf plan_ws.Cells(i, 3) = "or" And plan_ws.Cells(i, 1) <> plan_ws.Cells(i, 4) Then
            col_n1 = column_number(plan_ws.Cells(i, 1))
            col_n2 = column_number(plan_ws.Cells(i, 4))
            
            plan_ws.Range("M1") = plan_ws.Cells(i, 1)
            plan_ws.Range("M2") = plan_ws.Cells(i, 2)
            plan_ws.Range("N1") = plan_ws.Cells(i, 4)
            plan_ws.Range("N3") = plan_ws.Cells(i, 5)
            
            Set cr_rng = plan_ws.Range("M1").CurrentRegion
            rng.AdvancedFilter xlFilterInPlace, cr_rng
            
            If count_rows > 1 Then
                ws.Cells(2, col_n1).Select
                ws.Range(ws.Cells(2, col_n1), Selection.End(xlDown)).Select
                Call pattern_check(True)
                
                ws.Cells(2, col_n2).Select
                ws.Range(ws.Cells(2, col_n2), Selection.End(xlDown)).Select
                Call pattern_check(True)
            End If
            
        End If
        
        plan_ws.Range("M1") = Null
        plan_ws.Range("M2") = Null
        plan_ws.Range("N1") = Null
        plan_ws.Range("N2") = Null
        plan_ws.Range("N3") = Null
            
    Next
    
    If (ws.AutoFilterMode And ws.FilterMode) Or ws.FilterMode Then
        ws.ShowAllData
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Function count_rows() As Long

    Dim ws As Worksheet
    Dim uuid_col As Long
    Dim rows_n As Long
    uuid_col = gen_column_number("_uuid", find_main_data)
    Set ws = sheets(find_main_data)
    rows_n = ws.AutoFilter.Range.columns(uuid_col).SpecialCells(xlCellTypeVisible).Cells.count
    Debug.Print rows_n
    count_rows = rows_n

End Function

Sub add_logical_chek()
    On Error Resume Next
    Dim main_rng As Range
    Dim c_rng As Range
    Dim n As Long
    Dim c As String
    Dim q As String
    
    q = ThisWorkbook.sheets("logical_checks").Cells(2, 1)
    
    n = column_number(q)
    c = ThisWorkbook.sheets("logical_checks").Cells(2, 2)
    
    Set main_rng = sheets("uu").Range("A1").CurrentRegion
    sheets("uu").Activate
    
    If sheets("uu").FilterMode = True Then
        sheets("uu").ShowAllData
    End If

    main_rng.AutoFilter n, c

End Sub

Sub import_plan()
    On Error Resume Next
    Dim ws As Worksheet, strFile As String
    
    strFile = Application.GetOpenFilename("Plan Files (*.plan),*.plan", , "Please select a cleaning plan...")
    
    If strFile = "False" Then
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.sheets("logical_checks")
    ws.Cells.Clear
    
    With ws.QueryTables.Add(Connection:="TEXT;" & strFile, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh
    End With
End Sub

Sub export_planxx()
    On Error Resume Next
    Dim base_path As String
    Dim path As String
'
'    base_path = ActiveWorkbook.path
    Application.DisplayAlerts = False
    
    If ThisWorkbook.sheets("logical_checks").Range("A1") = vbNullString Then
        MsgBox "The logical check dose not exist!     ", vbInformation
        End
    End If
    
'    ThisWorkbook.sheets("logical_checks").Copy
'    ThisWorkbook.sheets ("logical_checks")
    'FileFilter:="Plan Files (*.plan), *.plan",
    
    path = Application.GetSaveAsFilename( _
           FileFilter:="Plan Files (*.txt), *.txt", _
           title:="Save the cleaning plan", _
           InitialFileName:="logical_ckeck")
            
    Workbooks.Add
    With ActiveWorkbook
        ThisWorkbook.sheets("logical_checks").Copy .sheets(1).Range("A1")
        .SaveAs path, FileFormat:=xlCurrentPlatformText
        .Close False
    End With
                
'    ActiveWorkbook.SaveAs Filename:=path, FileFormat:=xlCurrentPlatformText
'    ActiveWorkbook.Close
    
    Application.DisplayAlerts = True

End Sub

Sub export_plan()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbDest As Workbook
    Dim fName As String
    Dim last_row  As Long
    Dim path As String
    
    If ThisWorkbook.sheets("logical_checks").Range("A1") = vbNullString Then
        MsgBox "The logical check dose not exist!     ", vbInformation
        Exit Sub
    End If
    Application.DisplayAlerts = False
    
    Set wbSource = ThisWorkbook
    Set wsSource = ThisWorkbook.sheets("logical_checks")
    last_row = wsSource.Cells(rows.count, 1).End(xlUp).row

    path = Application.GetSaveAsFilename( _
           FileFilter:="Plan Files (*.plan), *.plan", _
           title:="Save the cleaning plan", _
           InitialFileName:="logical_ckeck")
             
    If path = "" Then End
    
    Workbooks.Add
    
    With ActiveWorkbook
        wsSource.Range("A1:F" & last_row).Copy .sheets(1).Range("A1")
        .SaveAs path, FileFormat:=xlCurrentPlatformText
        .Close True
    End With
    
    Application.DisplayAlerts = True
    
End Sub

Function remove_operator(str As String)
    On Error Resume Next
    Dim i As Long
    Dim char As Variant
    Dim new_str As String
    Dim char_set As String
    If str = vbNullString Then remove_operator = ""
    new_str = str
    char_set = ">,<,=, "
    
    For Each char In Split(char_set, ",")
        new_str = Replace(new_str, char, "")
    Next
    
    remove_operator = new_str
    
End Function

