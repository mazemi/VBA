Attribute VB_Name = "Datamerge_module"
Option Explicit

Sub generate_datamerge()

    Dim last_row_result  As Long
    Dim str_info As String
    Dim txt As String
    Dim ws As Worksheet
    Dim rng As Range
    
    Application.ScreenUpdating = False
       
    last_row_result = sheets("result").Cells(Rows.count, 1).End(xlUp).Row
    
    If last_row_result < 2 Then
        End
    End If

    If worksheet_exists("datamerge") Then
        Application.DisplayAlerts = False
        sheets("datamerge").Delete
        Application.DisplayAlerts = True
    End If
    
    Call create_sheet("result", "datamerge")
    
    Set ws = sheets("result")
    
    Set rng = ws.Range("A1").CurrentRegion
    rng.Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlYes

    Call make_dis_level
    Call make_header
    Call lookup
    DoEvents
    str_info = vbLf & analysis_form.TextInfo.value
    txt = "Formatting Datamerge... " & str_info
    analysis_form.TextInfo.value = txt
    analysis_form.Repaint
    
    Call add_label_to_datamerge
    Call populate_indicators
    Call make_dm_backend
    
    sheets("datamerge").Activate
    sheets("datamerge").Range("D5").Select
    ActiveWindow.FreezePanes = True
    
    str_info = vbLf & analysis_form.TextInfo.value
    txt = "Finalising... " & str_info
    analysis_form.TextInfo.value = txt
    sheets("datamerge").Range("A1").Select
    
End Sub

Sub make_dis_level()

    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Dim dt_ws As Worksheet
    Dim dis_ws As Worksheet
    Dim last_row_dt As Long
    Dim tmp_collection As New Collection
    Dim dis_collection As New Collection
    Dim uuid_col As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim q_col As String
    Dim c As Long
    Dim dt As String
    Dim rng As Range
    Dim last_row As Long
    Dim arr As Variant
    Dim dis_rng As Range
    Dim v As Variant
    Dim cel As Range
    
    Set res_ws = sheets("result")
    Set dt_ws = sheets(find_main_data)
    Set dm_ws = sheets("datamerge")
    Set dis_ws = sheets("disaggregation_setting")
    Set dis_rng = dis_ws.Range("A2:A" & dis_ws.Cells(dis_ws.Rows.count, "A").End(xlUp).Row)
    dm_ws.Cells.RowHeight = 14.2
    res_ws.Columns("B:D").Copy Destination:=dm_ws.Columns("A:C")
    Application.CutCopyMode = False
    
    dm_ws.Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
    last_row = dm_ws.Cells(Rows.count, "A").End(xlUp).Row
    
    Set tmp_collection = unique_values(dm_ws.Range("A2:A" & last_row))
    
    For Each cel In dis_rng
        For Each v In tmp_collection
            If cel = v Then
                dis_collection.Add v
            End If
        Next v
    Next cel
    
    
    uuid_col = gen_column_number("_uuid", find_main_data)
    
    last_row_dt = dt_ws.Cells(Rows.count, uuid_col).End(xlUp).Row
    dt = find_main_data

    If dm_ws.Range("A3") = vbNullString Then
        Exit Sub
    End If
        
    arr = dm_ws.Range("A2:B" & last_row)
 
    For i = 1 To UBound(arr, 1)
    
        ' real count
        If arr(i, 1) = "ALL" And arr(i, 2) = "ALL" Then
            dm_ws.Cells(i + 1, "D") = last_row_dt - 1
        Else
            q_col = gen_column_letter(CStr(arr(i, 1)), dt)
            c = Application.WorksheetFunction.CountIf(dt_ws.Range(q_col & "2:" & q_col & last_row_dt), arr(i, 2))
            dm_ws.Cells(i + 1, "D") = c
        End If
        
        ' levels order
        For j = 1 To dis_collection.count
            If dis_collection.item(j) = arr(i, 1) Then
                dm_ws.Cells(i + 1, "E") = j
            End If
        Next j
        
        'choice order
        dm_ws.Cells(i + 1, "F") = choice_order(CStr(arr(i, 1)), CStr(arr(i, 2)))
    Next i
    
    Set rng = dm_ws.Range("A1").CurrentRegion
    rng.Sort Key1:=dm_ws.Range("F2:F" & last_row), Order1:=xlAscending, Header:=xlYes
    rng.Sort Key1:=dm_ws.Range("E2:E" & last_row), Order1:=xlAscending, Header:=xlYes
    
    dm_ws.Columns("E:F").Delete
    dm_ws.Range("D1") = "count"
    dm_ws.Rows("1:3").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
End Sub

Sub make_header()
    Dim rng As Range
    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Dim ws As Worksheet
    Dim last_header As Long
    Dim last_indicator As Long
    Dim last_row As Long
    Dim analys_ws As Worksheet
    Dim c As Range
    Dim rng2 As Range
    Dim j As Long
    
    Set analys_ws = sheets("analysis_list")
    last_indicator = analys_ws.Cells(analys_ws.Rows.count, 1).End(xlUp).Row
    
    Set res_ws = sheets("result")
    Set dm_ws = sheets("datamerge")
    
    If Not worksheet_exists("temp_sheet") Then
        Call create_sheet("result", "temp_sheet")
    Else
        sheets("temp_sheet").Cells.Clear
    End If
    
    Set ws = sheets("temp_sheet")
    
    Set rng = res_ws.Range("A1").CurrentRegion
    rng.Sort Key1:=res_ws.Range("A1"), Order1:=xlAscending, Header:=xlYes
    
    ws.Columns("A:B").NumberFormat = "@"
    res_ws.Range("E:E,H:H,K:K").Copy Destination:=ws.Cells(1, "A")

    ws.Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2, 3), Header:=xlYes
    
    last_row = ws.Cells(Rows.count, 1).End(xlUp).Row
    ws.Range("D2:D" & last_row).Formula = "=IF(C2="""",A2& ""-value-"" &A2,A2& ""-value-"" &C2)"
    
    ws.Columns("D:D").Copy
    ws.Columns("D:D").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    ws.Columns("C:C").Delete
    
    last_header = ws.Cells(Rows.count, "A").End(xlUp).Row
    
    ws.Range("B2:C" & last_header).Copy
    dm_ws.Range("E3").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Application.CutCopyMode = False
End Sub

Sub lookup()
    On Error GoTo 0
    Application.ScreenUpdating = False

    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Dim temp_ws As Worksheet
    Dim rng As Range
    Dim last_col As Long
    Dim current_col As String
    Dim arr As Variant
    Dim dis_arr As Variant
    Dim header_arr As Variant
    Dim last_row_res As Long
    Dim last_row_dm As Long
    Dim last_row_temp As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Set res_ws = sheets("result")
    Set dm_ws = sheets("datamerge")
    Set temp_ws = sheets("temp_sheet")
    Set rng = res_ws.Range("A1").CurrentRegion
    temp_ws.Cells.Clear
    temp_ws.Visible = xlSheetVisible
    last_row_res = res_ws.Cells(Rows.count, 1).End(xlUp).Row
    last_row_dm = dm_ws.Cells(Rows.count, 1).End(xlUp).Row
    last_col = dm_ws.Cells(3, dm_ws.Columns.count).End(xlToLeft).Column
    
    dis_arr = dm_ws.Range("A5:B" & last_row_dm)
    
    For i = 1 To UBound(dis_arr, 1)
        With temp_ws
            .Cells.Clear
            .Range("A1") = "measurement value"
            .Range("B1") = "hkey order"
            .Range("D1") = "disaggregation"
            .Range("E1") = "disaggregation value"
            .Range("D2") = "'=" & dis_arr(i, 1)
            .Range("E2") = "'=" & dis_arr(i, 2)
        End With
        
        rng.AdvancedFilter xlFilterCopy, temp_ws.Range("D1").CurrentRegion, temp_ws.Range("A1:B1")
        last_row_temp = temp_ws.Cells(Rows.count, "A").End(xlUp).Row
        
        arr = temp_ws.Range("A2:B" & last_row_temp)
        For j = 1 To UBound(arr, 1)
            dm_ws.Cells(i + 4, arr(j, 2) + 4) = arr(j, 1)
        Next j
    Next i
    
    dm_ws.Columns("B:B").Delete
    res_ws.Columns("N:O").Delete
    
End Sub

Sub add_label_to_datamerge()
    Dim dm_ws As Worksheet
    Dim sc_ws As Worksheet
    Dim survey_ws As Worksheet
    Dim last_col As Long
    Dim last_row_survey_choice As Long
    Dim value_position As Long
    Dim question As String
    Dim choice As String
    Dim key As String
    Dim i As Long
    Dim j As Long
    Dim question_label As String
    Dim last_question As String
    Dim s As Long, e As Long
    Dim last_row_survey As Long
    Dim sc_arr As Variant
    Dim k As Long
    
    Set dm_ws = sheets("datamerge")
    Set sc_ws = ThisWorkbook.sheets("xsurvey_choices")
    Set survey_ws = ThisWorkbook.sheets("xsurvey")
    
    last_row_survey_choice = sc_ws.Cells(sc_ws.Rows.count, 1).End(xlUp).Row
    last_row_survey = survey_ws.Cells(survey_ws.Rows.count, 1).End(xlUp).Row
      
    dm_ws.Range("A1:A4").Merge
    dm_ws.Range("B1:B4").Merge
    dm_ws.Range("C1:C4").Merge
    
    last_col = dm_ws.Cells(3, Columns.count).End(xlToLeft).Column
    
    sc_arr = sc_ws.Range("A1").CurrentRegion
    
    'choices label loop
    For i = 4 To last_col
        key = Replace(dm_ws.Cells(4, i), "-value-", "")
        
        ' new approach
        For k = 1 To UBound(sc_arr, 1)
            If key = sc_arr(k, 6) Then
                dm_ws.Cells(2, i).value = sc_arr(k, 5)
                GoTo resume_loop
            End If
        Next k
        
resume_loop:

        value_position = Application.WorksheetFunction.Find("-value-", dm_ws.Cells(4, i))
        question = left(dm_ws.Cells(4, i), value_position - 1)
        question_label = find_question_label(question)
        dm_ws.Cells(1, i) = question_label

        If dm_ws.Cells(2, i) = vbNullString Then
            choice = Replace(dm_ws.Cells(4, i), question & "-value-", "")
            If question <> choice Then dm_ws.Cells(2, i) = choice
            If dm_ws.Cells(1, i) = vbNullString Then dm_ws.Cells(1, i) = question
        End If

    Next i
    
    Call merge_first_row
    Call styler
    
End Sub

Public Sub populate_indicators()
    Dim ws As Worksheet
    Dim indi_ws As Worksheet
    Dim last_col As Long
    Set ws = sheets("result")
    
    If Not worksheet_exists("indi_list") Then
        Call create_sheet(sheets(1).Name, "indi_list")
        sheets("indi_list").Visible = xlVeryHidden
    End If
    
    Set indi_ws = sheets("indi_list")
    With indi_ws
        .Cells.Clear
        .Columns("A:B").value = ws.Columns("E:F").value
        .Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
        .Columns("G").value = ws.Columns("B").value
        .Range("G1").CurrentRegion.RemoveDuplicates Columns:=1, Header:=xlYes
        .Rows("1:1").Delete Shift:=xlUp
    End With
End Sub

Function find_question_label(question As String) As String
    Dim s_ws As Worksheet
    Dim last_row_survey As Long
    Dim i As Long
    
    Set s_ws = ThisWorkbook.sheets("xsurvey")
    
    last_row_survey = s_ws.Cells(s_ws.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To last_row_survey
        If s_ws.Cells(i, "B").value = question Then
            find_question_label = s_ws.Cells(i, "C").value
            Exit Function
        End If
    Next i
    
    find_question_label = vbNullString
    
End Function

Function choice_order(question As String, choice As String) As Long
    Dim ws As Worksheet
    Dim keen_ws As Worksheet
    Dim rng As Range
    Dim choice_rng As Range
    Dim arr As Variant
    Dim last_row As Long
    Dim i As Long
    
    If Not worksheet_exists("keen") Then
        Call create_sheet(sheets(1).Name, "keen")
        sheets("keen").Visible = xlVeryHidden
    End If
    
    Set ws = ThisWorkbook.sheets("xsurvey_choices")
    Set rng = ws.Range("A1").CurrentRegion
    Set keen_ws = sheets("keen")
    With keen_ws
        .Cells.Clear
        .Range("A1") = "choice"
        .Range("C1") = "question"
        .Range("C2") = "'=" & question
    End With

    rng.AdvancedFilter xlFilterCopy, keen_ws.Range("C1:C2"), keen_ws.Range("A1"), True
    
    If keen_ws.Range("A2") = vbNullString Then
        choice_order = 0
        Exit Function
    End If
    
    Set choice_rng = keen_ws.Range("A1").CurrentRegion

    last_row = keen_ws.Cells(Rows.count, "A").End(xlUp).Row
    arr = keen_ws.Range("A2:A" & last_row)
    
    For i = 1 To UBound(arr, 1)
        If arr(i, 1) = choice Then
            choice_order = i
            Exit Function
        End If
    Next i
    
End Function

Sub merge_first_row()

    Dim ws As Worksheet
    Dim LastCol As Long
    Dim rng As Range
    Dim CurrentCol As Long, NextCol As Long
    Dim i As Long
    
    Set ws = sheets("datamerge")
    LastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Application.DisplayAlerts = False
    
    'loop through the columns in the first row
    For CurrentCol = 1 To LastCol
        Set rng = ws.Cells(1, CurrentCol)
        NextCol = CurrentCol + 1
        i = 1
        
        'loop while the next column has the same value as the current cell
        Do While NextCol <= LastCol And rng(rng.count).value = ws.Cells(1, NextCol).value
            'expand the range to include the next column
            Set rng = rng.Resize(1, i + 1)
            NextCol = NextCol + 1
            i = i + 1
        Loop
        
        If rng.Columns.count > 1 Then
            rng.Merge
            rng.HorizontalAlignment = xlCenter
        End If
        'set the current column to the last checked column
        CurrentCol = NextCol - 1
    Next CurrentCol
    
    Application.DisplayAlerts = True
End Sub

Sub styler()
    Dim ws As Worksheet
    Dim rng As Range
    Dim rng2 As Range
    Dim rng3 As Range
    Dim last_col As Long
    Dim last_col_letter As String
    
    Dim header_rng As Range
    Set ws = sheets("datamerge")
    Set rng = ws.Range("A3").CurrentRegion
    
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Set header_rng = ws.Rows("1:2")

    With header_rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlContext
    End With
  
    Set rng2 = ws.Range("A1:C4")

    rng2.HorizontalAlignment = xlCenter
    rng2.VerticalAlignment = xlCenter
    rng2.ReadingOrder = xlContext

    Set rng3 = ws.Rows("3:3")
    rng3.HorizontalAlignment = xlCenter
    rng3.ReadingOrder = xlContext
    rng3.VerticalAlignment = xlCenter

    ws.Rows(1).RowHeight = 43
    ws.Rows(2).RowHeight = 32
    last_col = ws.Cells(3, Columns.count).End(xlToLeft).Column
    last_col_letter = number_to_letter(last_col, ws)
        With ws
        .Columns("A:A").ColumnWidth = 12
        .Columns("B:B").ColumnWidth = 20
        .Columns("C:C").ColumnWidth = 6
        .Columns("D:" & last_col_letter).ColumnWidth = 14
        .Range("A1:" & last_col_letter & "4").Interior.Color = RGB(230, 230, 230)
        .Range("A1").CurrentRegion.Font.Size = 9
    End With

End Sub

Sub make_dm_backend()
    
    If worksheet_exists("dm_backend") Then
        sheets("dm_backend").Visible = xlSheetHidden
        Application.DisplayAlerts = False
        sheets("dm_backend").Delete
        Application.DisplayAlerts = True
    End If
    
    sheets("datamerge").Select
    sheets("datamerge").Copy Before:=sheets(1)
    ActiveSheet.Name = "dm_backend"
    sheets("dm_backend").Visible = xlVeryHidden
   
End Sub
