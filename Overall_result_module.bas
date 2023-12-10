Attribute VB_Name = "Overall_result_module"
Option Explicit

Global NO_NUMERIC As Boolean
Global NO_CATEGORICAL As Boolean
Global LAST_CATEGORICAL As Long

Sub all_result_data()
    
    On Error GoTo errHandler
    Dim res_sheet As Worksheet
    Dim temp_ws As Worksheet
    Dim xx_sheet As Worksheet
    Dim indi_rng As Range
    Dim c As Range
    Dim i As Long
    Dim last_row As Long
    Dim last_row_result  As Long
       
    Application.ScreenUpdating = False
    LAST_CATEGORICAL = 0
    NO_NUMERIC = False
    NO_CATEGORICAL = False
  
    CHART_COUNT = 0

    If Not worksheet_exists("result") Then
        Unload wait_form
        MsgBox "There is no analysis results, please run the analysis first.", vbInformation
        End
    End If
    
    wait_form.main_label = "Please wait ..."
    wait_form.Show vbModeless
    wait_form.Repaint
   
    Set res_sheet = sheets("result")
    
    If worksheet_exists("overall") Then
        Application.DisplayAlerts = False
        sheets("overall").Delete
        Application.DisplayAlerts = True
    End If
    
    If Not worksheet_exists("overall") Then
        Call create_sheet("result", "overall")
    End If
    
    If Not worksheet_exists("temp_sheet") Then
        Call create_sheet("result", "temp_sheet")
    End If
    
    Set temp_ws = sheets("temp_sheet")
    temp_ws.Cells.Clear
    Set xx_sheet = sheets("overall")

    last_row_result = sheets("result").Cells(rows.count, 1).End(xlUp).Row
    
    If (Worksheets("result").AutoFilterMode And Worksheets("result").FilterMode) Or Worksheets("result").FilterMode Then
        Worksheets("result").ShowAllData
    End If
    
    res_sheet.Range("$A$1:$M$" & last_row_result).AutoFilter Field:=2, Criteria1:="ALL"
    res_sheet.Range("$A$1:$M$" & last_row_result).AutoFilter Field:=8, Criteria1:="percentage"
    res_sheet.columns("E:L").Copy
    
'    temp_ws.Select
    temp_ws.Activate
    temp_ws.Range("A1").Select
    ActiveSheet.Paste

    last_row = temp_ws.Cells(rows.count, 1).End(xlUp).Row
    
    If last_row = 1 Then
        NO_CATEGORICAL = True
        GoTo extract_avereges
    End If
    
    Application.CutCopyMode = False
    
    temp_ws.Range("C:D,F:G").Delete Shift:=xlToLeft
    
    ' sort data
    temp_ws.Cells(1, 1).Select
    Selection.AutoFilter
    temp_ws.AutoFilter.Sort.SortFields.Clear
    temp_ws.AutoFilter.Sort.SortFields.Add key:=Range("C1:C" & last_row), SortOn:=xlSortOnValues, _
                                            Order:=xlDescending, DataOption:=xlSortNormal
                                                  
                                            
    With temp_ws.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    temp_ws.AutoFilter.Sort.SortFields.Clear
    temp_ws.AutoFilter.Sort.SortFields.Add key:=Range("A1:A" & last_row), SortOn:=xlSortOnValues, _
                                            Order:=xlAscending, DataOption:=xlSortNormal

    With temp_ws.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    temp_ws.columns("C:C").Cut
    
    temp_ws.columns("E:E").Select
    
    ActiveSheet.Paste
    
    temp_ws.columns("C:C").Delete Shift:=xlToLeft
    
    ' sorting
    If Not worksheet_exists("indi_list") Then
        Call populate_indicators
    End If

    Set indi_rng = sheets("indi_list").UsedRange
    temp_ws.Cells(1, 5) = "sorting"
    For Each c In indi_rng
        For i = 1 To last_row
            If c.Value = temp_ws.Cells(i, 2) Then
                temp_ws.Cells(i, 5) = c.Row
            End If
        Next
    Next
    
    temp_ws.Activate
    
    Call Range("A1").CurrentRegion.Sort(Key1:=Range("E2"), Order1:=xlAscending, Header:=xlYes)
              
    temp_ws.rows("1:1").Delete Shift:=xlUp

    Call make_seperate_data
    
extract_avereges:
    Call add_numeric_table("average")
    Call add_numeric_table("median")

    If NO_CATEGORICAL And NO_NUMERIC Then
        Unload wait_form
        MsgBox "Figures are only for overall disaggregation level." & vbCrLf & _
                "Please add 'ALL' in the disaggrigations for generating figures.", vbInformation

        Application.DisplayAlerts = False

        If worksheet_exists("temp_sheet") Then
            sheets("temp_sheet").Visible = xlSheetHidden
            sheets("temp_sheet").Delete
        End If

        If worksheet_exists("overall") Then
            sheets("overall").Delete
        End If

        Application.DisplayAlerts = True
        sheets("result").Activate
        Call clear_active_filter
        End
    End If

    If (res_sheet.AutoFilterMode And res_sheet.FilterMode) Or res_sheet.FilterMode Then
        res_sheet.ShowAllData
    End If
    
    Application.DisplayAlerts = False

    Application.DisplayAlerts = True
    sheets("result").Activate
    Call clear_active_filter
    xx_sheet.Activate
    
    Unload wait_form
    Application.ScreenUpdating = True
    
Exit Sub

errHandler:

If worksheet_exists("temp_sheet") Then
    sheets("temp_sheet").Visible = xlSheetHidden
    sheets("temp_sheet").Delete
End If

MsgBox "Oops!, Something went wrong!                       ", vbCritical
End
End Sub

Sub add_numeric_table(measurement As String)
    Application.DisplayAlerts = False
    Dim res_sheet As Worksheet
    Dim t_sheet As Worksheet
    Dim rng As Range
    Dim cr_rng As Range
    Dim all_sheet As Worksheet
    Dim last_row_overall As Long
    Dim last_average As Long
    Dim new_row_overall As Long
    
    Set all_sheet = sheets("overall")
    Set res_sheet = sheets("result")
    Set rng = res_sheet.Range("A1").CurrentRegion
    
    If Not worksheet_exists("temp_sheet") Then
        Call create_sheet("result", "temp_sheet")
    End If

    Set t_sheet = sheets("temp_sheet")
    t_sheet.Cells.Clear
    
    t_sheet.Range("A1") = "disaggregation"
    t_sheet.Range("B1") = "measurement type"
    t_sheet.Range("A2") = "ALL"
    t_sheet.Range("B2") = measurement
    t_sheet.Range("D1") = "variable label"
    t_sheet.Range("E1") = "measurement value"
    
    Set cr_rng = t_sheet.Range("A1").CurrentRegion
    
    rng.AdvancedFilter xlFilterCopy, cr_rng, t_sheet.Range("D1:E1")
    
    last_average = t_sheet.Cells(rows.count, 4).End(xlUp).Row
    
    If measurement = "average" Then
        new_row_overall = all_sheet.Cells(rows.count, 1).End(xlUp).Row + (16 - LAST_CATEGORICAL)
    Else
       new_row_overall = all_sheet.Cells(rows.count, 1).End(xlUp).Row + 3
    End If
    
    If NO_CATEGORICAL Then
        new_row_overall = 1
        all_sheet.columns("A:A").ColumnWidth = 65
        all_sheet.columns("B:B").ColumnWidth = 10
    End If
    
    If last_average > 1 Then
        
        t_sheet.Range("D1").CurrentRegion.Copy
        all_sheet.Activate
        all_sheet.Cells(new_row_overall, 1).Select
        all_sheet.Paste
        
        all_sheet.Cells(new_row_overall, 1) = "Indicators"
        all_sheet.Cells(new_row_overall, 2) = Application.WorksheetFunction.Proper(measurement)
        
        Call add_border(all_sheet.Cells(new_row_overall, 1).CurrentRegion)
        
        With all_sheet.Range(all_sheet.Cells(new_row_overall, 1), all_sheet.Cells(new_row_overall, 2)).Interior
            .Pattern = xlSolid
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.8
        End With
        
        all_sheet.Range(all_sheet.Cells(new_row_overall, 1), all_sheet.Cells(new_row_overall, 2)).Font.Bold = True
    Else
        NO_NUMERIC = True
    End If
    
    If worksheet_exists("temp_sheet") Then
        sheets("temp_sheet").Delete
    End If
    
    Application.DisplayAlerts = True
End Sub

Sub make_seperate_data()

'    On Error Resume Next
    Dim xx_sheet As Worksheet
    Dim t_sheet As Worksheet
    Dim last_row_overall As Long
    Dim chart_width As Long
    Dim unique_choices As New Collection
    Dim choices_numbers As New Collection
    Dim tbl_rng As Range
    Dim chart_type As Boolean
    Dim last_row_temp As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    Dim m As Long
    Dim tool_type As String
    
    Set xx_sheet = sheets("overall")
    Set t_sheet = sheets("temp_sheet")
    
    last_row_overall = xx_sheet.Cells(rows.count, 1).End(xlUp).Row
    last_row_temp = t_sheet.Cells(rows.count, 1).End(xlUp).Row
    
    ' need to check if unique values more than 255
    Set unique_choices = unique_values(t_sheet.Range("A1:A" & last_row_temp))
    
    xx_sheet.columns("A:A").ColumnWidth = 65
    xx_sheet.columns("B:B").ColumnWidth = 10
    xx_sheet.columns("C:C").ColumnWidth = 6
    
    For i = 1 To unique_choices.count
        k = 0
        For j = 1 To last_row_temp + 1
            If t_sheet.Cells(j, 1) = unique_choices(i) Then
                k = k + 1
            Else
                If k > 0 Then
                    choices_numbers.Add k
                    k = 0
                End If
            End If
        Next
    Next
    
    n = last_row_overall
    
    For m = 1 To choices_numbers.count
        
        chart_width = 300
        chart_type = False
        
        xx_sheet.Cells(n, 1) = t_sheet.Cells(1, 2)
        xx_sheet.Cells(n, 2) = "Percentage"
        
        With xx_sheet.Range(xx_sheet.Cells(n, 1), xx_sheet.Cells(n, 2)).Interior
            .Pattern = xlSolid
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.8
        End With
        
        With xx_sheet.rows(n)
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        xx_sheet.Range(xx_sheet.Cells(n, 1), xx_sheet.Cells(n, 2)).Font.Bold = True
        
        LAST_CATEGORICAL = choices_numbers(m)
        
        t_sheet.Range("C1:D" & choices_numbers(m)).Copy
        xx_sheet.Activate
        xx_sheet.Cells(n + 1, 1).Select
        xx_sheet.Paste
        
        Set tbl_rng = xx_sheet.Range(xx_sheet.Cells(n, 1), xx_sheet.Cells(n + choices_numbers(m), 2))
        
        Call add_border(tbl_rng)
        
        tool_type = question_type(t_sheet.Cells(1, 1))
        
        If tool_type = "select_multiple" Or tool_type = "select_multiple_external" Then
            chart_type = True
        ElseIf choices_numbers(m) > 7 Then
            chart_type = False
        Else
            chart_type = True
        End If
        
        If choices_numbers(m) > 4 Then
            chart_width = Application.WorksheetFunction.Round(choices_numbers(m) * 280 / 4, 0)
        End If
        
        Dim ch_title As String, ch_title2 As String
        
        ch_title2 = t_sheet.Cells(1, 2).Value
        
        If Len(ch_title2) > 150 Then
            ch_title2 = left(ch_title2, 150)
        End If
        
        If choices_numbers(m) < 15 Then
            Call add_barchart(tbl_rng, ch_title2, chart_type, xx_sheet.Cells(n, 1).top, xx_sheet.Cells(n, 4).left, chart_width)
        End If
        
        t_sheet.Activate
        
        rows("1:" & choices_numbers(m)).Select
        Selection.Delete Shift:=xlUp
        
        If choices_numbers(m) < 15 Then
            n = xx_sheet.Cells(rows.count, 1).End(xlUp).Row + 15 - choices_numbers(m) + 2
        Else
            n = xx_sheet.Cells(rows.count, 1).End(xlUp).Row + 2
        End If

    Next
    
End Sub

Sub add_barchart(input_rng As Range, title As String, bar As Boolean, top As String, left As String, Optional bar_width As Long)
'    On Error Resume Next
    Dim ws As Worksheet
    Dim rng As Range
    Dim my_chart As Object
    Dim chtSeries

    Set ws = Worksheets("overall")
    Set rng = input_rng
    Set my_chart = ws.Shapes.AddChart2
    
    CHART_COUNT = CHART_COUNT + 1
    
    With my_chart.Chart
        .SetSourceData rng
        If bar Then
            .ChartType = xlColumnClustered
            .SetElement (msoElementDataLabelOutSideEnd)
            .Parent.Width = bar_width
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
            .SeriesCollection(1).Interior.Color = RGB(4, 49, 76)
        Else
            .ChartType = xlPie
            .Parent.Width = 500
            .SetElement (msoElementLegendRight)
            .Legend.left = 270
            .Legend.Width = 270
            .PlotArea.Width = 160
            .PlotArea.Height = 160
            .PlotArea.left = 20
            .PlotArea.top = 40
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
            .SetElement msoElementDataLabelInsideEnd
        End If
        .ChartTitle.text = title
        .Parent.top = top
        .Parent.left = left
        
    End With
    
End Sub

Sub add_border(rng As Range)
'    On Error Resume Next
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

End Sub



