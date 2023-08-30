Attribute VB_Name = "Datamerge_module"
Option Explicit

Sub generate_datamerge()
    On Error Resume Next
    Application.ScreenUpdating = False
    DoEvents
    Dim last_row_result  As Long
    Dim str_info As String
    Dim txt As String
    
    last_row_result = sheets("result").Cells(rows.count, 1).End(xlUp).row
    
    If last_row_result < 2 Then
        End
    End If
          
    Call make_dis_level
    Call make_header
    Call make_key
    Call lookup
    Call clean_up
    Call add_label_to_datamerge
    Call populate_indicators
    
    sheets("datamerge").Activate
    sheets("datamerge").Range("D4").Select
    ActiveWindow.FreezePanes = True
    
    Application.ScreenUpdating = True
    
    str_info = vbLf & analysis_form.TextInfo.value

    txt = "Analysis finished. " & str_info

    analysis_form.TextInfo.value = txt
    sheets("datamerge").Range("A1").Select
    Application.Wait (Now + 0.00001)
    
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
    
    Set dm_ws = sheets("datamerge")
    Set sc_ws = ThisWorkbook.sheets("xsurvey_choices")
    Set survey_ws = ThisWorkbook.sheets("xsurvey")
    
    last_row_survey_choice = sc_ws.Cells(sc_ws.rows.count, 1).End(xlUp).row
    last_row_survey = survey_ws.Cells(survey_ws.rows.count, 1).End(xlUp).row
    
    
'    Debug.Print last_col
    
    dm_ws.rows("1:2").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    dm_ws.Range("A1:A3").Merge
    dm_ws.Range("B1:B3").Merge
    dm_ws.Range("C1:C3").Merge
    
    last_col = dm_ws.Cells(3, columns.count).End(xlToLeft).column
    
    'choices label loop
    For i = 4 To last_col
        key = Replace(dm_ws.Cells(3, i), "-value-", "")
        For j = 2 To last_row_survey_choice
            If key = sc_ws.Cells(j, "F").value Then
                dm_ws.Cells(2, i).value = sc_ws.Cells(j, "E").value
                GoTo resume_loop
            End If
        Next j
resume_loop:
    Next i
            
    'question label loop
    For i = 4 To last_col
        value_position = Application.WorksheetFunction.Find("-value-", dm_ws.Cells(3, i))
        question = left(dm_ws.Cells(3, i), value_position - 1)
        question_label = find_question_label(question)
        dm_ws.Cells(1, i) = question_label
    Next i

    ' dealing with calculation or custome indicators
    
    For i = 4 To last_col
        If dm_ws.Cells(2, i) = vbNullString Then
            value_position = Application.WorksheetFunction.Find("-value-", dm_ws.Cells(3, i))
            question = left(dm_ws.Cells(3, i), value_position - 1)
            choice = Replace(dm_ws.Cells(3, i), question & "-value-", "")
            If question <> choice Then dm_ws.Cells(2, i) = choice
            If dm_ws.Cells(1, i) = vbNullString Then dm_ws.Cells(1, i) = question
        End If
     
    Next i

    Call merge_first_row
    Call styler
End Sub

Private Sub populate_indicators()

'    Dim header_arr() As Variant
    Dim ws As Worksheet
    Dim indi_ws As Worksheet
    Dim last_col As Long
    
    Set ws = sheets("datamerge")
    
    If Not worksheet_exists("indi_list") Then
        Call create_sheet(sheets(1).Name, "indi_list")
        sheets("indi_list").visible = xlVeryHidden
    End If
    
    Set indi_ws = sheets("indi_list")
    indi_ws.Cells.Clear
    
    last_col = ws.Cells(1, columns.count).End(xlToLeft).column

    ws.Range(ws.Cells(1, 4), ws.Cells(1, last_col)).Copy
    indi_ws.Range("A1").PasteSpecial _
        Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=True, transpose:=True
        
    Call delete_blanks
    
End Sub

Sub delete_blanks()
    
    Dim rng As Range
    Dim str_delete As String
    Dim last_row As Long
    Dim last_indi As Long
    Dim last_measurement As Long
    Dim rows As Long, i As Long
    
    last_row = sheets("indi_list").Range("A" & sheets("indi_list").rows.count).End(xlUp).row
    Set rng = sheets("indi_list").Range("A1:A" & last_row)

    Set rng = sheets("indi_list").Range("A1:A" & last_row)
    rows = rng.rows.count
    For i = rows To 1 Step (-1)
        If WorksheetFunction.CountA(rng.rows(i)) = 0 Then rng.rows(i).Delete
    Next
End Sub


Function find_question_label(question As String) As String

    Dim s_ws As Worksheet
    Dim last_row_survey As Long
    Dim i As Long
    
    Set s_ws = ThisWorkbook.sheets("xsurvey")
    
    last_row_survey = s_ws.Cells(s_ws.rows.count, 1).End(xlUp).row
    
    For i = 2 To last_row_survey
        If s_ws.Cells(i, "B").value = question Then
            find_question_label = s_ws.Cells(i, "C").value
            Exit Function
        End If
    Next i
    
    find_question_label = vbNullString
    
End Function

Sub make_dis_level()
    On Error Resume Next
    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Dim dt_ws As Worksheet
    Dim res_rng As Range
    Dim dm_rng As Range
    Dim var_rng As Range
    Dim dis_label As Range
    Dim tmp As Range
    Dim last_row_dm As Long
    Dim last_row_dt As Long
    Dim dis_collection As New Collection
    Dim last_row_result As Long
    Dim dis As Variant
    Dim uuid_col As Long
    Dim i As Long
    Dim q_col As String
    Dim c As Long
    Dim dt As String
    
    Set res_ws = sheets("result")
    Set dt_ws = sheets(find_main_data)
    
    If Not worksheet_exists("datamerge") Then
        Call create_sheet("result", "datamerge")
    End If

    Set dm_ws = sheets("datamerge")
    
    dm_ws.Cells.Clear
    dm_ws.Cells.RowHeight = 14.4
    
    Set res_rng = res_ws.Range("A1").CurrentRegion
    
    res_rng.Sort key1:=res_rng.Range("E1"), Order1:=xlAscending, Header:=xlYes
    res_rng.Sort key1:=res_rng.Range("D1"), Order1:=xlAscending, Header:=xlYes
    res_rng.Sort key1:=res_rng.Range("B1"), Order1:=xlAscending, Header:=xlYes
    
    last_row_result = res_ws.Cells(res_ws.rows.count, 1).End(xlUp).row
    
    Set dis_collection = unique_values(res_ws.Range("B2:B" & last_row_result))
    
    For Each dis In dis_collection
        Call row_helper(CStr(dis))
    Next
    
    uuid_col = gen_column_number("_uuid", find_main_data)
    
    last_row_dt = dt_ws.Cells(rows.count, uuid_col).End(xlUp).row
    
    last_row_dm = dm_ws.Cells(dm_ws.rows.count, 1).End(xlUp).row
    
    dt = find_main_data
    For i = 2 To last_row_dm
        
        If dm_ws.Cells(i, "E") = "ALL" Then
            dm_ws.Cells(i, "F") = last_row_dt - 1
        Else
            q_col = gen_column_letter(dm_ws.Cells(i, "B").value, dt)
            c = Application.WorksheetFunction.CountIfs(dt_ws.Range(q_col & "2:" & q_col & last_row_dt), dm_ws.Cells(i, "E"))
            dm_ws.Cells(i, "F") = c
        End If
         
    Next
    
    dm_ws.columns("D:E").Delete
    
End Sub


Sub make_header()
    On Error Resume Next
    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Dim ws As Worksheet
    Dim last_header As Long
    Dim last_indicator As Long
    Dim last_row As Long
    Dim analys_ws As Worksheet
    Dim c As Range
    Dim rng As Range
    Dim j As Long
    
    Set analys_ws = sheets("analysis_list")
    
    last_indicator = analys_ws.Cells(analys_ws.rows.count, 1).End(xlUp).row
    
    Set res_ws = sheets("result")
    Set dm_ws = sheets("datamerge")
    
    If Not worksheet_exists("temp_sheet") Then
        Call create_sheet("datamerge", "temp_sheet")
    End If
    
    Set ws = sheets("temp_sheet")
    
    ws.Cells.Clear
    
    ws.Range("A:A").Value2 = res_ws.Range("E:E").Value2
    ws.Range("B:B").Value2 = res_ws.Range("K:K").Value2
    
    last_row = ws.Cells(rows.count, 1).End(xlUp).row
    
    ws.Range("C2:C" & last_row).Formula = "=IF(B2="""",A2& ""-value-"" &A2,A2& ""-value-"" &B2)"
    
    ws.columns("C:C").Copy

    ws.columns("E:E").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                                   SkipBlanks:=False, transpose:=False
        
    Application.CutCopyMode = False
    
    ws.Range("$E$1:$E$" & last_row).RemoveDuplicates columns:=1, Header:=xlNo
    
    ' add order number
    last_header = ws.Cells(rows.count, "E").End(xlUp).row
    
    ws.Range("F2:F" & last_header).Formula = "=find(""-value-"", E2)"
    
    ws.Range("G2:G" & last_header).Formula = "=left(E2, F2 -1)"
    
    Set rng = analys_ws.Range("A2:A" & last_indicator)
    
    For Each c In rng
        
        For j = 2 To last_header
            If c = ws.Cells(j, "G") Then
                ws.Cells(j, "H").value = c.row
            End If
        Next j
        
    Next
    
    ws.Range("E1") = "E"
    ws.Range("F1") = "F"
    ws.Range("G1") = "G"
    ws.Range("H1") = "H"
    
    ws.Range("E1").CurrentRegion.Sort key1:=Range("H1:H" & last_header), Order1:=xlAscending, Header:=xlYes
        
    ws.Range("E1") = vbNullString
    ws.Range("F:F").Cells.ClearContents
    
    ' copy to datamerge sheet
    ws.Range("E2").CurrentRegion.Copy

    dm_ws.Range("E1").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, transpose:=True

    dm_ws.Range("A1") = "key"
    dm_ws.Range("B1") = "Disaggregation"
    dm_ws.Range("C1") = "Disaggregation Label"
    dm_ws.Range("D1") = "Count"
     
End Sub

Sub make_key()
'    On Error Resume Next
    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Dim last_row As Long
    
    Set res_ws = sheets("result")
    Set dm_ws = sheets("datamerge")
    
    last_row = res_ws.Cells(rows.count, 1).End(xlUp).row
    
    res_ws.columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Application.CutCopyMode = False
    
    res_ws.Range("A1") = "key"
    
    res_ws.Range("A2:A" & last_row).Formula = "=C2 & E2 & IF(L2="""",F2& ""-value-"" &F2, F2& ""-value-"" &L2)"
    
End Sub

Sub lookup()
'    On Error Resume Next
    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Dim last_row As Long
    Dim last_col As Long
    Dim current_col As String
    
    Set res_ws = sheets("result")
    Set dm_ws = sheets("datamerge")
    
    last_row = dm_ws.Cells(rows.count, 1).End(xlUp).row
    last_col = dm_ws.Cells(1, columns.count).End(xlToLeft).column
    
    Dim i As Long
    Dim formula_str As String
    
    For i = 5 To last_col
        current_col = Public_module.number_to_letter(i, dm_ws)
        formula_str = "=VLOOKUP(A2" & "&" & current_col & "$1" & ",result!A:J,10,)"
        dm_ws.Range(current_col & "2:" & current_col & last_row).Formula = formula_str
    Next

End Sub

Sub clean_up()
'    On Error Resume Next
    sheets("datamerge").Range("A1").CurrentRegion.Copy

    sheets("datamerge").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                 :=False, transpose:=False
    Application.CutCopyMode = False
    
    sheets("datamerge").Cells.Replace What:="#N/A", replacement:="", LookAt:=xlWhole, _
                                      SearchOrder:=xlByColumns, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    
    Application.DisplayAlerts = False
    sheets("temp_sheet").Delete
    Application.DisplayAlerts = True
    
    sheets("datamerge").columns("A:A").Delete Shift:=xlToLeft
   
    sheets("datamerge").rows("1:1").RowHeight = 32
    
    With sheets("datamerge").rows("1:1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlContext
    End With
    
    sheets("datamerge").Cells.EntireColumn.AutoFit
    
    sheets("result").columns("A:A").Delete Shift:=xlToLeft

End Sub

Public Sub row_helper(val As String)
'    On Error Resume Next
    Dim data_ws As Worksheet
    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Dim data_rng As Range
    Dim arrTemp As Variant, key As Variant, k As Variant
    Dim dict As Object
    Dim dict_code As Object
    Dim i As Long
    Dim last_row As Long
    Dim col As String
    Dim survey_count As Long
    Dim new_row As Long
    Dim new_row2 As Long
    
    Set data_ws = sheets(find_main_data)
    Set data_rng = data_ws.Range("A1").CurrentRegion
    
    Set res_ws = sheets("result")
    Set dm_ws = sheets("datamerge")
    
    arrTemp = res_ws.[a1].CurrentRegion.value
    
    last_row = data_rng.rows(data_rng.rows.count).row
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set dict_code = CreateObject("Scripting.Dictionary")

    For i = LBound(arrTemp) To UBound(arrTemp)
        If arrTemp(i, 2) = val Then
            dict(arrTemp(i, 4)) = 1
            dict_code(arrTemp(i, 3)) = 1
        End If
    Next i
    
    new_row = dm_ws.Cells(rows.count, 1).End(xlUp).row + 1
    new_row2 = dm_ws.Cells(rows.count, 1).End(xlUp).row + 1

    For Each key In dict.Keys
        dm_ws.Cells(new_row, 1) = val & key
        dm_ws.Cells(new_row, 2) = val
        dm_ws.Cells(new_row, 3) = key
        new_row = new_row + 1
    Next key

    For Each k In dict_code.Keys
    
        If k <> "ALL" Then
            col = column_letter(val)
            survey_count = Application.WorksheetFunction.CountIf(data_ws.Range(col & "1:" & col & last_row), k)
            dm_ws.Cells(new_row2, 4) = survey_count
            dm_ws.Cells(new_row2, 5) = k
            new_row2 = new_row2 + 1
        Else
            dm_ws.Cells(new_row2, 4) = last_row - 1
            dm_ws.Cells(new_row2, 5) = k
            new_row2 = new_row2 + 1
        End If
        
    Next k
    
End Sub

Sub merge_first_row()
    Dim ws As Worksheet
    Dim LastCol As Long
    Dim rng As Range
    Dim CurrentCol As Long, NextCol As Long
    Dim i As Long
    
    Set ws = sheets("datamerge")
    LastCol = ws.Cells(1, ws.columns.count).End(xlToLeft).column
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
        
        If rng.columns.count > 1 Then
            rng.Merge
            rng.HorizontalAlignment = xlCenter
        End If
        'set the current column to the last checked column
        CurrentCol = NextCol - 1
        
    Next CurrentCol
    
    Application.DisplayAlerts = True
End Sub

Sub styler()

    Dim rng As Range
    Dim last_col As Long
    Dim last_col_letter As String
    
    Dim header_rng As Range
    Set rng = sheets("datamerge").Range("A3").CurrentRegion
    
    rng.Borders(xlDiagonalDown).LineStyle = xlNone
    rng.Borders(xlDiagonalUp).LineStyle = xlNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .weight = xlThin
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .weight = xlThin
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .weight = xlThin
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .weight = xlThin
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .weight = xlThin
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .weight = xlThin
    End With
    
    Set header_rng = sheets("datamerge").rows("1:2")

    With header_rng
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlContext
    End With
  
    sheets("datamerge").rows(1).RowHeight = 60
    sheets("datamerge").rows(2).RowHeight = 44
    last_col = sheets("datamerge").Cells(3, columns.count).End(xlToLeft).column
    last_col_letter = number_to_letter(last_col, sheets("datamerge"))
    sheets("datamerge").columns("A:A").ColumnWidth = 14
    sheets("datamerge").columns("B:B").ColumnWidth = 30
    sheets("datamerge").columns("C:C").ColumnWidth = 10
    sheets("datamerge").columns("D:" & last_col_letter).ColumnWidth = 18
    sheets("datamerge").Range("A1:" & last_col_letter & "3").Interior.Color = RGB(230, 230, 230)
    sheets("datamerge").Range("A1:" & last_col_letter & "3").Font.Size = 10
    
End Sub


