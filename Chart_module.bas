Attribute VB_Name = "Chart_module"
Option Explicit
Global NO_NUMERIC As Boolean
Global NO_CATEGORICAL As Boolean
Global LAST_CATEGORICAL As Long
Global DISAGGREGATION_LEVEL As String
Global DISAGGREGATION_VALUE As String
Global DISAGGREGATION_LABEL As String
Global SHEET_NAME As String
Global IS_OVERALL As Boolean

Sub generate_multiple_data_chart(dis_level As String, val_collection As Collection, label_collection As Collection)
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim count As Integer
    wait_form.main_label = "Please wait ..."
    wait_form.labelLine.Visible = True
    wait_form.Show vbModeless
    wait_form.Repaint
    
    DISAGGREGATION_LEVEL = dis_level
    count = val_collection.count
    Debug.Print count
    
    For i = 1 To count
        Debug.Print val_collection.item(i)
        DISAGGREGATION_VALUE = val_collection.item(i)
        DISAGGREGATION_LABEL = label_collection.item(i)
        wait_form.note = "Proccesing " & DISAGGREGATION_LABEL
        wait_form.Repaint
        Call generate_data_chart
    Next i
    
    Unload wait_form
    
    Exit Sub

ErrorHandler:

If worksheet_exists("temp_sheet") Then
    sheets("temp_sheet").Visible = xlSheetHidden
    sheets("temp_sheet").Delete
End If

Application.ScreenUpdating = True
MsgBox "Oops!, Something went wrong!                       ", vbCritical
End

End Sub

Sub generate_data_chart()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Dim res_sheet As Worksheet
    Dim temp_ws As Worksheet
    Dim chart_sheet As Worksheet
    Dim indi_rng As Range
    Dim has_all_dis As Boolean
    Dim c As Range
    Dim i As Long
    Dim last_row As Long
    Dim last_row_result  As Long
    Dim last_row_indi_list As Long
    
    If DISAGGREGATION_LEVEL = "ALL" And DISAGGREGATION_VALUE = vbNullString Then
        SHEET_NAME = "overall"
        IS_OVERALL = True
    ElseIf DISAGGREGATION_LEVEL <> "ALL" And DISAGGREGATION_VALUE <> vbNullString Then
        SHEET_NAME = proper_sheet_name(DISAGGREGATION_LEVEL & "@" & DISAGGREGATION_LABEL)
        IS_OVERALL = False
    End If
    
    LAST_CATEGORICAL = 0
    NO_NUMERIC = False
    NO_CATEGORICAL = False
  
    CHART_COUNT = 0
    
    If IS_OVERALL Then
        wait_form.main_label = "Please wait ..."
        wait_form.labelLine.Visible = True
        wait_form.note = "Proccesing overall figures"
        wait_form.Show vbModeless
        wait_form.Repaint
    End If

    Set res_sheet = sheets("result")
    
    If worksheet_exists(SHEET_NAME) Then
        Application.DisplayAlerts = False
        sheets(SHEET_NAME).Delete
        Application.DisplayAlerts = True
    End If
    
    Set chart_sheet = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.count))
    chart_sheet.Name = SHEET_NAME

    If Not worksheet_exists("temp_sheet") Then
        Call create_sheet("result", "temp_sheet")
    End If
    
    Set temp_ws = sheets("temp_sheet")
    temp_ws.Cells.Clear
    Set chart_sheet = sheets(SHEET_NAME)

    last_row_result = sheets("result").Cells(Rows.count, 1).End(xlUp).Row
    
    If (Worksheets("result").AutoFilterMode And Worksheets("result").FilterMode) Or Worksheets("result").FilterMode Then
        Worksheets("result").ShowAllData
    End If
    
    If IS_OVERALL Then
        res_sheet.Range("$A$1:$M$" & last_row_result).AutoFilter Field:=2, Criteria1:=DISAGGREGATION_LEVEL
    Else
        res_sheet.Range("$A$1:$M$" & last_row_result).AutoFilter Field:=2, Criteria1:=DISAGGREGATION_LEVEL
        res_sheet.Range("$A$1:$M$" & last_row_result).AutoFilter Field:=3, Criteria1:=DISAGGREGATION_VALUE
    End If
    
    res_sheet.Range("$A$1:$M$" & last_row_result).AutoFilter Field:=8, Criteria1:="percentage"
    res_sheet.Columns("E:L").Copy
    
    temp_ws.Activate
    temp_ws.Range("A1").Select
    ActiveSheet.Paste

    last_row = temp_ws.Cells(Rows.count, 1).End(xlUp).Row
    
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
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    temp_ws.AutoFilter.Sort.SortFields.Clear
    temp_ws.AutoFilter.Sort.SortFields.Add key:=Range("A1:A" & last_row), SortOn:=xlSortOnValues, _
                                            Order:=xlAscending, DataOption:=xlSortNormal

    With temp_ws.AutoFilter.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    temp_ws.Columns("C:C").Cut
    temp_ws.Columns("E:E").Select
    ActiveSheet.Paste
    temp_ws.Columns("C:C").Delete Shift:=xlToLeft
    
    ' sorting
    If Not worksheet_exists("indi_list") Then
        Call populate_indicators
    End If

    last_row_indi_list = sheets("indi_list").Cells(Rows.count, 1).End(xlUp).Row
    Set indi_rng = sheets("indi_list").Range("B1:B" & last_row_indi_list)
    
    temp_ws.Cells(1, 5) = "sorting"
    For Each c In indi_rng
        For i = 1 To last_row
            If c.value = temp_ws.Cells(i, 2) Then
                temp_ws.Cells(i, 5) = c.Row
            End If
        Next
    Next
    
    temp_ws.Activate
    Call Range("A1").CurrentRegion.Sort(Key1:=Range("E2"), Order1:=xlAscending, header:=xlYes)
    temp_ws.Rows("1:1").Delete Shift:=xlUp
    Call make_seperate_data
    
extract_avereges:
    Call add_numeric_table("average")
    Call add_numeric_table("median")

    If NO_CATEGORICAL And NO_NUMERIC Then
        Unload wait_form
        MsgBox "Please first analyze the data with desired disaggrigations level." & vbCrLf & _
            "Then try to generate charts.", vbInformation

        Application.DisplayAlerts = False
        
        On Error Resume Next

        If worksheet_exists("temp_sheet") Then
            sheets("temp_sheet").Visible = xlSheetHidden
            sheets("temp_sheet").Delete
        End If

        If worksheet_exists(SHEET_NAME) Then
            sheets(SHEET_NAME).Delete
        End If

        Application.DisplayAlerts = True
        sheets("result").Activate
        Call clear_active_filter
        
        Application.DisplayAlerts = True
        End
    End If

    If (res_sheet.AutoFilterMode And res_sheet.FilterMode) Or res_sheet.FilterMode Then
        res_sheet.ShowAllData
    End If
    
    sheets("result").Activate
    Call clear_active_filter
    chart_sheet.Activate
    
    If IS_OVERALL Then
        Unload wait_form
    End If
    
    chart_sheet.Columns("A:B").Font.Size = 10
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
Exit Sub

ErrorHandler:

If worksheet_exists("temp_sheet") Then
    sheets("temp_sheet").Visible = xlSheetHidden
    sheets("temp_sheet").Delete
End If
Application.ScreenUpdating = True
MsgBox "Oops!, Something went wrong!                       ", vbCritical
End
End Sub

Sub add_numeric_table(measurement As String)
    Application.DisplayAlerts = False
    Dim res_sheet As Worksheet
    Dim t_sheet As Worksheet
    Dim rng As Range
    Dim cr_rng As Range
    Dim chart_sheet As Worksheet
    Dim last_row_overall As Long
    Dim last_average As Long
    Dim new_row As Long
    
    Set chart_sheet = sheets(SHEET_NAME)
    Set res_sheet = sheets("result")
    Set rng = res_sheet.Range("A1").CurrentRegion
    
    If Not worksheet_exists("temp_sheet") Then
        Call create_sheet("result", "temp_sheet")
    End If

    Set t_sheet = sheets("temp_sheet")
    
    With t_sheet
        .Cells.Clear
        .Range("A1:C1") = Array("disaggregation", "disaggregation value", "measurement type")
        
        If IS_OVERALL Then
            .Range("A2") = DISAGGREGATION_LEVEL
        Else
            .Range("A2") = "'=" & DISAGGREGATION_LEVEL
            .Range("B2") = "'=" & DISAGGREGATION_VALUE
        End If
        
        .Range("C2") = measurement
        .Range("E1:F1") = Array("variable label", "measurement value")
    
        Set cr_rng = .Range("A1").CurrentRegion
        rng.AdvancedFilter xlFilterCopy, cr_rng, .Range("E1:F1")
        last_average = .Cells(Rows.count, 5).End(xlUp).Row
    End With
    
    If measurement = "average" Then
        If LAST_CATEGORICAL > 12 Then
            new_row = chart_sheet.Cells(Rows.count, 1).End(xlUp).Row + 3
        Else
            new_row = chart_sheet.Cells(Rows.count, 1).End(xlUp).Row + (16 - LAST_CATEGORICAL)
        End If
    Else
       new_row = chart_sheet.Cells(Rows.count, 1).End(xlUp).Row + 3
    End If
    
    If NO_CATEGORICAL Then
        new_row = 1
        chart_sheet.Columns("A:A").ColumnWidth = 60
        chart_sheet.Columns("B:B").ColumnWidth = 10
    End If
    
    If last_average > 1 Then
        
        t_sheet.Range("E1").CurrentRegion.Copy
        With chart_sheet
            .Activate
            .Cells(new_row, 1).Select
            .Paste
            .Cells(new_row, 1) = "Indicators"
            .Cells(new_row, 2) = Application.WorksheetFunction.Proper(measurement)
        End With
        
        Call add_border(chart_sheet.Cells(new_row, 1).CurrentRegion)
        
        With chart_sheet.Range(chart_sheet.Cells(new_row, 1), chart_sheet.Cells(new_row, 2)).Interior
            .Pattern = xlSolid
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.8
        End With
        
        chart_sheet.Range(chart_sheet.Cells(new_row, 1), chart_sheet.Cells(new_row, 2)).Font.Bold = True
    Else
        NO_NUMERIC = True
    End If
    
    If worksheet_exists("temp_sheet") Then
        sheets("temp_sheet").Delete
    End If
    
    Application.DisplayAlerts = True
End Sub

Sub make_seperate_data()

    Dim chart_sheet As Worksheet
    Dim t_sheet As Worksheet
    Dim last_row_overall As Long
    Dim chart_width As Long
    Dim chart_height As Long
    Dim unique_choices As New Collection
    Dim choices_numbers As New Collection
    Dim tbl_rng As Range
    Dim chart_type As String
    Dim last_row_temp As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    Dim m As Long
    Dim tool_type As String
    
    Set chart_sheet = sheets(SHEET_NAME)
    Set t_sheet = sheets("temp_sheet")
    
    last_row_overall = chart_sheet.Cells(Rows.count, 1).End(xlUp).Row
    last_row_temp = t_sheet.Cells(Rows.count, 1).End(xlUp).Row
    
    ' need to check if unique values more than 255
    Set unique_choices = unique_values(t_sheet.Range("A1:A" & last_row_temp))
    
    chart_sheet.Columns("A:A").ColumnWidth = 60
    chart_sheet.Columns("B:B").ColumnWidth = 10
    chart_sheet.Columns("C:C").ColumnWidth = 6
    
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
        
        chart_sheet.Cells(n, 1) = t_sheet.Cells(1, 2)
        chart_sheet.Cells(n, 2) = "Percentage"
        
        With chart_sheet.Range(chart_sheet.Cells(n, 1), chart_sheet.Cells(n, 2)).Interior
            .Pattern = xlSolid
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.8
        End With
        
        With chart_sheet.Rows(n)
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        chart_sheet.Range(chart_sheet.Cells(n, 1), chart_sheet.Cells(n, 2)).Font.Bold = True
        
        LAST_CATEGORICAL = choices_numbers(m)
        
        t_sheet.Range("C1:D" & choices_numbers(m)).Copy
        chart_sheet.Activate
        chart_sheet.Cells(n + 1, 1).Select
        chart_sheet.Paste
        
        Set tbl_rng = chart_sheet.Range(chart_sheet.Cells(n, 1), chart_sheet.Cells(n + choices_numbers(m), 2))
        
        Call add_border(tbl_rng)
        
        tool_type = question_type(t_sheet.Cells(1, 1))
        
        On Error Resume Next
        
        If IS_OVERALL And choices_numbers(m) = 2 And tbl_rng.Value2(3, 2) > 3 And _
            (tool_type = "select_one" Or tool_type = "select_one_external") Then
            chart_type = "pie"
            
        ElseIf Not IS_OVERALL And choices_numbers(m) = 2 And tbl_rng.Value2(3, 2) > 3 And _
            (tool_type = "select_one" Or tool_type = "select_one_external") Then
            chart_type = "pie"
            
        ElseIf choices_numbers(m) < 16 Then
            chart_type = "col"
        Else
            chart_type = "bar"
        End If
        
        On Error GoTo 0
        
        If (m) > 4 And choices_numbers(m) < 16 Then
            chart_width = Application.WorksheetFunction.Round(choices_numbers(m) * 290 / 4, 0)
        ElseIf choices_numbers(m) >= 16 And choices_numbers(m) < 35 Then
            chart_height = Application.WorksheetFunction.Round(choices_numbers(m) * 57 / 4, 0)
        End If
        
        Dim ch_title As String, ch_title2 As String
        
        ch_title2 = t_sheet.Cells(1, 2).value
        
        If Len(ch_title2) > 150 Then
            ch_title2 = left(ch_title2, 150)
        End If
        
        If choices_numbers(m) < 16 Then
            Call add_barchart(tbl_rng, ch_title2, chart_type, chart_sheet.Cells(n, 1).top, chart_sheet.Cells(n, 4).left, chart_width)
        ElseIf choices_numbers(m) < 35 Then
            Call add_barchart(tbl_rng, ch_title2, chart_type, chart_sheet.Cells(n, 1).top, chart_sheet.Cells(n, 4).left, , chart_height)
        End If
        
        t_sheet.Activate
        
        Rows("1:" & choices_numbers(m)).Select
        Selection.Delete Shift:=xlUp
        
        If choices_numbers(m) < 15 Then
            n = chart_sheet.Cells(Rows.count, 1).End(xlUp).Row + 15 - choices_numbers(m) + 2
        Else
            n = chart_sheet.Cells(Rows.count, 1).End(xlUp).Row + 2
        End If

    Next
    
End Sub

Sub add_barchart(input_rng As Range, title As String, chart_type As String, top As String, left As String, Optional chart_width As Long, Optional chart_height As Long)
    On Error Resume Next
    Dim ws As Worksheet
    Dim rng As Range
    Dim my_chart As Object

    Set ws = Worksheets(SHEET_NAME)
    Set rng = input_rng
    Set my_chart = ws.Shapes.AddChart2
    
    CHART_COUNT = CHART_COUNT + 1
    
    With my_chart.Chart
        .SetSourceData rng
        If chart_type = "col" Then
            .ChartType = xlColumnClustered
            .SetElement (msoElementDataLabelOutSideEnd)
            .Parent.Width = chart_width
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
            .SeriesCollection(1).Interior.Color = RGB(4, 49, 76)
            
        ElseIf chart_type = "bar" Then
            .ChartType = xlBarClustered
            .SetElement (msoElementDataLabelOutSideEnd)
            .Parent.Width = 500
            .Parent.Height = chart_height
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
            .SeriesCollection(1).Interior.Color = RGB(4, 49, 76)
        
        ElseIf chart_type = "pie" Then
            .ChartType = xlPie
            .Parent.Width = 300
            .SetElement (msoElementLegendRight)
            .Legend.left = 200
            .Legend.Width = 100
            .PlotArea.Width = 160
            .PlotArea.Height = 160
            .PlotArea.left = 20
            .PlotArea.top = 40
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
            .SetElement msoElementDataLabelInsideEnd
        End If
        
        .ChartTitle.Text = title
        .Parent.top = top
        .Parent.left = left
        
    End With
    
End Sub

Sub add_border(rng As Range)

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

End Sub

Function proper_sheet_name(s_name As String) As String
    Dim illegal_chars As String
    Dim i As Integer
    
    ' illegal characters
    illegal_chars = "\/:*?""<>|"
    
    For i = 1 To Len(illegal_chars)
        s_name = Replace(s_name, Mid(illegal_chars, i, 1), "")
    Next i
    
    proper_sheet_name = left(Trim(s_name), 30)
End Function



