Attribute VB_Name = "Datamerge_module"
Sub generate_datamerge()
    On Error Resume Next
    Application.ScreenUpdating = False
    DoEvents
    last_row_result = sheets("result").Cells(rows.count, 1).End(xlUp).row
    
    If last_row_result < 2 Then
        End
    End If
         
'    Dim wb As Workbook
'    Set wb = ActiveWorkbook
    
    Call make_dis_level
    Call make_header
    Call make_key
    Call lookup
    Call clean_up

    sheets("datamerge").Activate
    
    Application.ScreenUpdating = True
    
    str_info = vbLf & analysis_form.TextInfo.value

    txt = "Analysis finished. " & str_info

    analysis_form.TextInfo.value = txt
    
    Application.Wait (Now + 0.00001)
    
End Sub

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
    
    Set res_ws = sheets("result")
    Set dt_ws = sheets(find_main_data)
    
    If Not WorksheetExists("datamerge") Then
        Call create_sheet("result", "datamerge")
    End If

    Set dm_ws = sheets("datamerge")
    
    dm_ws.Cells.Clear
    
    Set res_rng = res_ws.Range("A1").CurrentRegion
    
    res_rng.Sort Key1:=res_rng.Range("E1"), Order1:=xlAscending, Header:=xlYes
    res_rng.Sort Key1:=res_rng.Range("D1"), Order1:=xlAscending, Header:=xlYes
    res_rng.Sort Key1:=res_rng.Range("B1"), Order1:=xlAscending, Header:=xlYes
    
    last_row_result = res_ws.Cells(res_ws.rows.count, 1).End(xlUp).row
    
    Set dis_collection = unique_values(res_ws.Range("B2:B" & last_row_result))
    
    For Each dis In dis_collection
        Call row_helper(CStr(dis))
    Next
    
    uuid_col = gen_column_number("_uuid", find_main_data)
    
    last_row_dt = dt_ws.Cells(rows.count, uuid_col).End(xlUp).row
    
    last_row_dm = dm_ws.Cells(dm_ws.rows.count, 1).End(xlUp).row
    
    For i = 2 To last_row_dm
        If dm_ws.Cells(i, "E") = "ALL" Then
            dm_ws.Cells(i, "F") = last_row_dt - 1
        Else
            q_col = gen_column_letter(dm_ws.Cells(i, "B"), find_main_data)
            c = Application.WorksheetFunction.CountIfs(dt_ws.Range(q_col & "2:" & q_col & last_row_dt), dm_ws.Cells(i, "E"))
            dm_ws.Cells(i, "F") = c
        End If
         
    Next
    
    dm_ws.columns("D:E").Delete
    
End Sub

Sub order_col()
    Dim analys_ws As Worksheet
    Dim tmp_ws As Worksheet
    Dim rng As Range
    
    Set analys_ws = sheets("analysis_list")
    Set tmp_ws = sheets("temp_sheet")
    last_row = analys_ws.Cells(analys_ws.rows.count, 1).End(xlUp).row
    
    last_tmp = tmp_ws.Cells(tmp_ws.rows.count, 5).End(xlUp).row
    
    Set rng = analys_ws.Range("A2:A" & last_row)
    
    For Each c In rng
        Debug.Print c, c.row
        
        For j = 2 To last_tmp
            If c = tmp_ws.Cells(j, "G") Then
                tmp_ws.Cells(j, "I").value = c.row
            End If
        Next j
        
    Next
    
    
    
        
End Sub

Sub make_header()
    On Error Resume Next
    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Dim ws As Worksheet
    Dim last_header As Long
    Dim last_indicator As Long
    
    Dim analys_ws As Worksheet

    Dim rng As Range
    
    Set analys_ws = sheets("analysis_list")
    
    last_indicator = analys_ws.Cells(analys_ws.rows.count, 1).End(xlUp).row
    
    Set res_ws = sheets("result")
    Set dm_ws = sheets("datamerge")
    
    If Not WorksheetExists("temp_sheet") Then
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
        ' Debug.Print c, c.row
        
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
    
    ws.Range("E1").CurrentRegion.AutoFilter.Sort.SortFields.Clear
    
    ws.AutoFilter.Sort.SortFields.Add2 key:=Range("H1:H" & last_header), _
         SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
    With ws.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
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
    On Error Resume Next
    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Set res_ws = sheets("result")
    Set dm_ws = sheets("datamerge")
    
    last_row = res_ws.Cells(rows.count, 1).End(xlUp).row
    
    res_ws.columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Application.CutCopyMode = False
    
    res_ws.Range("A1") = "key"
    
    res_ws.Range("A2:A" & last_row).Formula = "=C2 & E2 & IF(L2="""",F2& ""-value-"" &F2, F2& ""-value-"" &L2)"
    
End Sub

Sub lookup()
    On Error Resume Next
    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
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
    On Error Resume Next
    sheets("datamerge").Range("A1").CurrentRegion.Copy

    sheets("datamerge").Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                 :=False, transpose:=False
    Application.CutCopyMode = False
    
    sheets("datamerge").Cells.Replace What:="#N/A", replacement:="", LookAt:=xlWhole, _
                                      SearchOrder:=xlByColumns, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
'                                      ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
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
    On Error Resume Next
    Dim data_ws As Worksheet
    Dim res_ws As Worksheet
    Dim dm_ws As Worksheet
    Dim data_rng As Range
    Dim arrTemp As Variant, key As Variant, K As Variant
    Dim dict As Object
    Dim dict_code As Object
    Dim i As Long
    Dim last_row As Long
    Dim col As String
    Dim survey_count As Long
    
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
    
'    new_row = 2
'    new_row2 = 2

    For Each key In dict.Keys
        dm_ws.Cells(new_row, 1) = val & key
        dm_ws.Cells(new_row, 2) = val
        dm_ws.Cells(new_row, 3) = key
        new_row = new_row + 1
    Next key

    For Each K In dict_code.Keys
    
        If K <> "ALL" Then
            col = column_letter(val)
            survey_count = Application.WorksheetFunction.CountIf(data_ws.Range(col & "1:" & col & last_row), K)
            dm_ws.Cells(new_row2, 4) = survey_count
            dm_ws.Cells(new_row2, 5) = K
            new_row2 = new_row2 + 1
        Else
            dm_ws.Cells(new_row2, 4) = last_row - 1
            dm_ws.Cells(new_row2, 5) = K
            new_row2 = new_row2 + 1
        End If
        
    Next K
    
End Sub



