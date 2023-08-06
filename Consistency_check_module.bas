Attribute VB_Name = "Consistency_check_module"

Sub consistency_check()
    Application.ScreenUpdating = False
    
    If WorksheetExists("survey_choices") <> True Then
        MsgBox "Please import the tool.  ", vbInformation
        Exit Sub
    End If
    
    Call setup_check
    Call data_injection("province")
    Call log_value_inconsistency

    Application.ScreenUpdating = True
End Sub

Sub setup_check()

Dim tool_ws As Worksheet
Dim temp_ws As Worksheet
Dim dt_ws As Worksheet

Set tool_ws = sheets("survey_choices")
Set dt_ws = sheets(find_main_data)

If WorksheetExists("temp_sheet") <> True Then
    Call create_sheet(dt_ws.Name, "temp_sheet")
End If

Set temp_ws = sheets("temp_sheet")

temp_ws.Cells.Clear

temp_ws.Range("A1") = "type"
temp_ws.Range("A2") = "select_one"
temp_ws.Range("A3") = "select_multiple"
temp_ws.Range("A5") = "question"

temp_ws.Range("C1") = "question"
temp_ws.Range("E1") = "choice"
temp_ws.Range("F1") = "choice_label"

last_row_dt = dt_ws.UsedRange.rows(dt_ws.UsedRange.rows.count).row

tool_ws.Range("A1").CurrentRegion.AdvancedFilter Action:=xlFilterCopy, _
         CriteriaRange:=temp_ws.Range("A1").CurrentRegion, CopyToRange:=temp_ws.Range("C1"), Unique:=True

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
    
    Set temp_ws = sheets("temp_sheet")
    Set dt_ws = sheets(find_main_data)
    
    Set colle = tool_value_choice(question, True)
    
    uuid_col = gen_column_number("_uuid", dt_ws.Name)
    c1 = gen_column_number(question, dt_ws.Name)
    c2 = gen_column_number(question & "_label", dt_ws.Name)
    
    If c2 = 0 Then
        temp_ws.columns("H:H").value = dt_ws.columns(uuid_col).value
        temp_ws.columns("I:I").value = dt_ws.columns(c1).value
        temp_ws.columns("L:L").value = dt_ws.columns(c1).value
        
        temp_ws.Range("M1:M" & last_row).RemoveDuplicates columns:=1, Header:=xlYes
        
        With Range("M1:M" & last_choice)
            If WorksheetFunction.CountA(.Cells) > 0 Then
                For i = last_choice To 2 Step -1
                   If LenB(Cells(i, 13)) = 0 Then Cells(i, 13).EntireRow.Delete Shift:=xlShiftUp
                Next
            End If
        End With
        
    Else
    
        temp_ws.columns("H:H").value = dt_ws.columns(uuid_col).value
        temp_ws.columns("I:I").value = dt_ws.columns(c1).value
        temp_ws.columns("J:J").value = dt_ws.columns(c2).value
        temp_ws.columns("M:M").value = dt_ws.columns(c1).value
        temp_ws.columns("N:N").value = dt_ws.columns(c2).value
        
        temp_ws.Range("J1") = question & "_labelX"
        temp_ws.Range("N1") = question & "_labelX"
        
        last_row = temp_ws.Cells(rows.count, 8).End(xlUp).row
               
        temp_ws.Range("M1:M" & last_row).RemoveDuplicates columns:=1, Header:=xlYes
        temp_ws.Range("N1:N" & last_row).RemoveDuplicates columns:=1, Header:=xlYes
        
        last_choice = temp_ws.Cells(rows.count, 13).End(xlUp).row
        last_label = temp_ws.Cells(rows.count, 14).End(xlUp).row
        
        Debug.Print last_choice, last_label
        
        With Range("M1:M" & last_choice)
            If WorksheetFunction.CountA(.Cells) > 0 Then
                For i = last_choice To 2 Step -1
                   If LenB(Cells(i, 13)) = 0 Then Cells(i, 13).Delete Shift:=xlShiftUp
                Next
            End If
        End With
        
        With Range("N1:N" & last_label)
            If WorksheetFunction.CountA(.Cells) > 0 Then
                For i = last_choice To 2 Step -1
                   If LenB(Cells(i, 14)) = 0 Then Cells(i, 14).Delete Shift:=xlShiftUp
                Next
            End If
        End With
        
        Call add_question_label(question)
        
    End If
    
End Sub

Sub log_value_inconsistency()
    Dim tool_rng As Range
    Dim value_rng As Range
    Dim dt_rng As Range
    Dim dt_ws As Worksheet
    Dim log_ws As Worksheet
    
    Dim inconsistant_values() As String
    Dim i As Long
    Dim new_log As Long
    
    Set temp_ws = sheets("temp_sheet")
    Set log_ws = sheets("log_book")

    last_row_tool = temp_ws.Cells(rows.count, 5).End(xlUp).row
    last_row_value = temp_ws.Cells(rows.count, 14).End(xlUp).row
    last_row_dt = temp_ws.Cells(rows.count, 8).End(xlUp).row
    
    Set tool_rng = Range("E2:E" & last_row_tool)
    Set value_rng = Range("N2:N" & last_row_value)

    inconsistant_values = get_inconsistency(tool_rng, value_rng)

    For i = 0 To UBound(inconsistant_values)
        Debug.Print inconsistant_values(i)
        
        For j = 2 To last_row_dt
            If temp_ws.Cells(j, 9) = inconsistant_values(i) Then
                new_log = log_ws.Cells(rows.count, 1).End(xlUp).row + 1
    
                log_ws.Cells(new_log, "A").value = temp_ws.Cells(j, 8)
                log_ws.Cells(new_log, "E").value = temp_ws.Cells(j, 9)
                log_ws.Cells(new_log, "B").value = temp_ws.Cells(1, 9)
                log_ws.Cells(new_log, "C").value = "invalid option"
            End If
        Next j
    Next i
    
End Sub

Function get_inconsistency(ByRef r1 As Range, ByRef r2 As Range) As String()
    Dim cell As Range
    Dim found As Range

    Dim uniques() As String
    Dim i As Long

    For Each cell In r2
        On Error Resume Next
        Set found = r1.Find(cell.value)
        On Error GoTo 0

        If (found Is Nothing) Then
            ReDim Preserve uniques(i)
            uniques(i) = cell.value
            i = i + 1
        End If
    Next cell

    get_inconsistency = uniques
End Function

Private Function tool_value_choice(q_name As String, with_choice As Boolean) As Collection

    Dim temp_ws As Worksheet
    Dim tool_ws As Worksheet
    Dim last_choice As Long
    Dim coll As New Collection
    
    Set tool_ws = sheets("survey_choices")
    Set temp_ws = sheets("temp_sheet")
    
    temp_ws.Range("A6") = q_name
    
    temp_ws.Range("E:F").ClearContents
    temp_ws.Range("E1") = "choice"
    temp_ws.Range("F1") = "choice_label"
    
    If with_choice Then
        tool_ws.Range("A1").CurrentRegion.AdvancedFilter Action:=xlFilterCopy, _
                CriteriaRange:=temp_ws.Range("A5").CurrentRegion, CopyToRange:=temp_ws.Range("E1:F1")
                
        last_choice = temp_ws.Cells(rows.count, 5).End(xlUp).row
        
        For Each c In temp_ws.Range("E2:E" & last_choice)
            coll.Add c.value
        Next
    Else
        tool_ws.Range("A1").CurrentRegion.AdvancedFilter Action:=xlFilterCopy, _
                CriteriaRange:=temp_ws.Range("A5").CurrentRegion, CopyToRange:=temp_ws.Range("E1:F1")
                
        last_label = temp_ws.Cells(rows.count, 6).End(xlUp).row
        
        For Each c In temp_ws.Range("F2:F" & last_label)
            coll.Add c.value
        Next
    End If
    
    Set tool_value_choice = coll
    
End Function




