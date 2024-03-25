Attribute VB_Name = "Consistency_check_module"
Option Explicit
Global SKIP_QUESTION As Boolean

Sub consistency_check()
    On Error GoTo errhandler
    Application.DisplayAlerts = False
    Dim t As Date
    Dim last_question As Long
    Dim i As Long
    Dim uuid_col As Long

    wait_form.main_label = "Please wait ..."
    wait_form.Show vbModeless
    wait_form.labelLine.Visible = True
    wait_form.Repaint
    
    Application.ScreenUpdating = False
    
    uuid_col = gen_column_number("_uuid", find_main_data)

    If uuid_col = 0 Then
         MsgBox "There is no _uuid column in the main dataset.", vbInformation
         Unload wait_form
         Exit Sub
    End If
    
    Dim tmp_ws As Worksheet
    If ThisWorkbook.sheets("xsurvey").Range("A1") = vbNullString Then
        MsgBox "Please import the tool.  ", vbInformation
        Unload wait_form
        Exit Sub
    End If
    
    Call remove_empty_col
    Call setup_check
    
    Set tmp_ws = sheets("temp_sheet")
    
    last_question = tmp_ws.Cells(Rows.count, 1).End(xlUp).Row
    
    If last_question < 11 Then
        MsgBox "No categorical question detected.  ", vbInformation
        Unload wait_form
        Exit Sub
    End If
    
    For i = 11 To last_question
        DoEvents
        wait_form.note = "Proccesing " & tmp_ws.Cells(i, 1)
'        Application.StatusBar = tmp_ws.Cells(i, 1)
        SKIP_QUESTION = False
        If no_value(tmp_ws.Cells(i, 1)) Then GoTo resume_loop
        
        Call data_injection(tmp_ws.Cells(i, 1))
  
        If SKIP_QUESTION Then GoTo resume_loop
        Call log_value_inconsistency
           
resume_loop:
    Next i
    
    If worksheet_exists("temp_sheet") Then
        Application.DisplayAlerts = False
        sheets("temp_sheet").Visible = xlSheetHidden
        sheets("temp_sheet").Delete
    End If
        
    Unload wait_form
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Exit Sub
    
errhandler:

    If worksheet_exists("temp_sheet") Then
        sheets("temp_sheet").Visible = xlSheetHidden
        sheets("temp_sheet").Delete
    End If
    
    Unload wait_form
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Application.DisplayAlerts = True
    MsgBox "Checking failed!       ", vbInformation

End Sub

Sub setup_check()
    Dim last_row_dt As Long
    Dim tool_ws As Worksheet
    Dim temp_ws As Worksheet
    Dim dt_ws As Worksheet
    
    Set tool_ws = ThisWorkbook.sheets("xsurvey_choices")
    Set dt_ws = sheets(find_main_data)
    
    If worksheet_exists("temp_sheet") <> True Then
        Call create_sheet(dt_ws.Name, "temp_sheet")
    End If
    
    Set temp_ws = sheets("temp_sheet")
    
    temp_ws.Cells.Clear
    
    temp_ws.Range("A1") = "type"
    temp_ws.Range("A2") = "select_one"
    
    ' maybe later
    ' temp_ws.Range("A3") = "select_multiple"
    
    temp_ws.Range("A5") = "question"
     
    temp_ws.Range("A10") = "question"

    temp_ws.Range("C1") = "choice"

    last_row_dt = dt_ws.UsedRange.Rows(dt_ws.UsedRange.Rows.count).Row
    
    tool_ws.Range("A1").CurrentRegion.AdvancedFilter Action:=xlFilterCopy, _
             CriteriaRange:=temp_ws.Range("A1").CurrentRegion, CopyToRange:=temp_ws.Range("A10"), Unique:=True

End Sub

Sub data_injection(question As String)

    Dim temp_ws As Worksheet
    Dim dt_ws As Worksheet
    Dim colle As New Collection
    Dim last_row As Long
    Dim q As String
    Dim uuid_col As Long
    Dim c1 As Long
    Dim c2 As Long
    Dim rng As Range
    Dim last_choice As Long
    Dim i As Long
    
    Set temp_ws = sheets("temp_sheet")
    temp_ws.Range("B:P").Delete
    Set dt_ws = sheets(find_main_data)
    
    Call tool_value_choice(question)
    
    uuid_col = gen_column_number("_uuid", dt_ws.Name)
    c1 = gen_column_number(question, dt_ws.Name)
    c2 = gen_column_number(question & "_label", dt_ws.Name)
       
    If c1 = 0 Then
        SKIP_QUESTION = True
        Exit Sub
    End If
    
    temp_ws.Columns("G:G").value = dt_ws.Columns(uuid_col).value
    temp_ws.Columns("H:H").value = dt_ws.Columns(c1).value
    temp_ws.Columns("E:E").value = dt_ws.Columns(c1).value
    
    temp_ws.Range("E1").value = temp_ws.Range("E1").value & "_unique"
    
    If c2 > 0 Then

        temp_ws.Columns("I:I").value = dt_ws.Columns(c2).value
        
        temp_ws.Range("I1") = question & "_labelX"
               
        temp_ws.Activate
        Call add_question_label(question)
        
    End If
    
    last_row = temp_ws.Cells(Rows.count, 7).End(xlUp).Row
    
    temp_ws.Range("E1:E" & last_row).RemoveDuplicates Columns:=1, Header:=xlYes

    last_choice = temp_ws.Cells(Rows.count, 5).End(xlUp).Row
    
    With temp_ws.Range("E1:E" & last_choice)
        If WorksheetFunction.CountA(.Cells) > 0 Then
            For i = last_choice To 2 Step -1
               If LenB(temp_ws.Cells(i, 5)) = 0 Then temp_ws.Cells(i, 5).Delete Shift:=xlShiftUp
            Next
        End If
    End With
         
End Sub

Sub log_value_inconsistency()
    Dim tool_rng As Range
    Dim value_rng As Range
    Dim dt_rng As Range
    Dim dt_ws As Worksheet
    Dim temp_ws As Worksheet
    Dim log_ws As Worksheet
    Dim inconsistant_values() As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim new_log As Long
    Dim last_row_tool As Long
    Dim last_row_value As Long
    Dim last_row_dt As Long
    
    Set temp_ws = sheets("temp_sheet")
    
    If worksheet_exists("log_book") <> True Then
        Call create_log_sheet(find_main_data)
    End If
    
    Set log_ws = sheets("log_book")

    last_row_tool = temp_ws.Cells(Rows.count, 3).End(xlUp).Row ' col C, choice
    last_row_value = temp_ws.Cells(Rows.count, 5).End(xlUp).Row ' col E, unique values in the dataset
    last_row_dt = temp_ws.Cells(Rows.count, 7).End(xlUp).Row ' col G, _uuid
    
    Set tool_rng = temp_ws.Range("C2:C" & last_row_tool)
    Set value_rng = temp_ws.Range("E2:E" & last_row_value)

    inconsistant_values = get_inconsistency(tool_rng, value_rng)

    If (Not inconsistant_values) = -1 Then GoTo label_check
    
    For i = 0 To UBound(inconsistant_values)
        For j = 2 To last_row_dt
            If temp_ws.Cells(j, 8) = inconsistant_values(i) Then
                new_log = log_ws.Cells(Rows.count, 1).End(xlUp).Row + 1
                log_ws.Cells(new_log, "A").value = temp_ws.Cells(j, 7)
                log_ws.Cells(new_log, "B").value = temp_ws.Cells(1, 8)
                log_ws.Cells(new_log, "C").value = "invalid option"
                log_ws.Cells(new_log, "D").value = temp_ws.Cells(j, 8)
            End If
        Next j
    Next i
    
label_check:

    If Right(temp_ws.Range("J1"), 6) = "labelX" Then
        For k = 2 To last_row_dt
            If temp_ws.Cells(k, "I") <> temp_ws.Cells(k, "J") Then
                new_log = log_ws.Cells(Rows.count, 1).End(xlUp).Row + 1
                log_ws.Cells(new_log, "A").value = temp_ws.Cells(k, "G")
                log_ws.Cells(new_log, "B").value = temp_ws.Range("I1")
                log_ws.Cells(new_log, "C").value = "check the label"
                log_ws.Cells(new_log, "D").value = temp_ws.Cells(k, "J")
            End If
        Next
    End If
    
End Sub


Function get_inconsistency(ByRef tool_rng As Range, ByRef dt_rng As Range) As String()
    Dim cell As Range
    Dim found As Range

    Dim uniques() As String
    Dim i As Long
    i = 0
    For Each cell In dt_rng
        If (Application.WorksheetFunction.CountIf(tool_rng, cell)) < 1 Then
            ReDim Preserve uniques(i)
            uniques(i) = cell.value
            i = i + 1
        End If
    Next cell
    
    get_inconsistency = uniques
End Function

Private Sub tool_value_choice(q_name As String)

    Dim temp_ws As Worksheet
    Dim tool_ws As Worksheet
    Dim last_choice As Long
    Dim coll As New Collection
    
    Set tool_ws = ThisWorkbook.sheets("xsurvey_choices")
    Set temp_ws = sheets("temp_sheet")
    
    temp_ws.Range("A6") = q_name
    temp_ws.Range("C1") = "choice"
    
    tool_ws.Range("A1").CurrentRegion.AdvancedFilter Action:=xlFilterCopy, _
            CriteriaRange:=temp_ws.Range("A5").CurrentRegion, CopyToRange:=temp_ws.Range("C1")
            
    last_choice = temp_ws.Cells(Rows.count, 3).End(xlUp).Row
        
End Sub




