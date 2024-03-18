VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} single_chart_form 
   Caption         =   "Generate summary for one indicatore"
   ClientHeight    =   6648
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   10464
   OleObjectBlob   =   "single_chart_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "single_chart_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandBack_Click()
    Unload Me
    chart_form.Show
End Sub

Private Sub CommandRun_Click()
    On Error GoTo erHandler
    Application.ScreenUpdating = False
    Dim SelectedItemIndex As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Long
    Dim ws As Worksheet
    Dim dm_ws As Worksheet
    Dim sheetName As String
    Dim selected_var As String
    Dim var_arr() As Variant
    Dim last_dm_col As Long
    Dim last_dm_col_letter As String
    Dim dm_col_start As String
    Dim dm_col_end As String
    Dim chart_col_end As String
    Dim v As Variant
    Dim position As Integer
    Dim last_row As Double
    Dim dis_count As Long
    Dim option_count As Long
    
    If Me.ComboBoxDis.value = vbNullString Then
        Exit Sub
    End If
    
    SelectedItemIndex = -1
    
    For i = 0 To Me.ListBoxVars.ListCount - 1
        If ListBoxVars.Selected(i) Then
            SelectedItemIndex = i
        End If
    Next i
    
    If SelectedItemIndex = -1 Then
        Exit Sub
    End If
    
    selected_var = ListBoxVars.List(SelectedItemIndex)
    
    If selected_var = Me.ComboBoxDis.value Then
        MsgBox "The disaggregation level and the selected variable are the same. Please choose another variable.", vbInformation
        Exit Sub
    End If
 
    sheetName = "chart"
    i = 1

    Do While worksheet_exists(sheetName & "-" & i)
        i = i + 1
    Loop
    
    Set dm_ws = sheets("dm_backend")
    Set ws = sheets.Add(after:=sheets(sheets.count))
    ws.Name = sheetName & "-" & i
    
    last_row = dm_ws.Cells(Rows.count, 1).End(xlUp).Row
    ws.Range("A1:B" & last_row).value = dm_ws.Range("A1:B" & last_row).value
    
    last_dm_col = dm_ws.Cells(3, Columns.count).End(xlToLeft).Column
    last_dm_col_letter = number_to_letter(last_dm_col, dm_ws)
    var_arr = dm_ws.Range("D4:" & last_dm_col_letter & 4)
    var_arr = Application.Transpose(Application.Transpose(var_arr))
    k = 1
    Dim m As Long

    For j = 1 To UBound(var_arr)
        If selected_var = left(var_arr(j), InStr(var_arr(j), "-value-") - 1) Then
            k = k + 1
            m = j
        End If
    Next j
    
    chart_col_end = number_to_letter(k + 1, ws)
    dm_col_start = number_to_letter(m + 5 - k, dm_ws)
    dm_col_end = number_to_letter(m + 3, dm_ws)
    ws.Range("C1:" & chart_col_end & last_row).value = dm_ws.Range(dm_col_start & "1:" & dm_col_end & last_row).value
    
    Call arrange_table(ws)
    
    dis_count = ws.Cells(ws.Rows.count, 1).End(xlUp).Row - 2
    option_count = k - 3
    
    If option_count > 12 And dis_count > 35 Then
        Exit Sub
    End If
        
    Call make_chart(ws)
    
    Application.ScreenUpdating = True
    Exit Sub
    
erHandler:
    MsgBox " Oops!, Something went wrong!  ", vbInformation
    Unload single_chart_form

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Exit Sub
    
End Sub

Private Sub make_chart(ws As Worksheet)
    On Error Resume Next
    Dim rng As Range
    Dim ave_rng As Range
    Dim med_rng As Range
    Dim ave_chart As Object
    Dim med_chart As Object
    Dim cat_chart As Object
    Dim last_row As Long
    Dim last_col As Long
    
    last_row = ws.Cells(Rows.count, 1).End(xlUp).Row
    last_col = ws.Cells(2, Columns.count).End(xlToLeft).Column
    
    If NUMERIC_CHART Then
        
        Set ave_chart = ws.Shapes.AddChart2
        Set med_chart = ws.Shapes.AddChart2
        Set ave_rng = ws.Range("A2:B" & last_row)
        Set med_rng = Union(ws.Range("A2:A" & last_row), ws.Range("C2:C" & last_row))
        With ave_chart.Chart
            .SetSourceData ave_rng
            .ChartType = xlColumnClustered
            .SetElement (msoElementDataLabelOutSideEnd)
            .HasLegend = False
            .Parent.Width = 20 * last_row + 120
            .Parent.Height = 200
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
            .SeriesCollection(1).Interior.Color = RGB(4, 49, 76)
            .ChartTitle.Text = left(ws.Range("A1").value, 150) & " [Average]"
            .Parent.top = 0
            .Parent.left = 230
        End With
    
        With med_chart.Chart
            .SetSourceData med_rng
            .ChartType = xlColumnClustered
            .SetElement (msoElementDataLabelOutSideEnd)
            .HasLegend = False
            .Parent.Width = 20 * last_row + 120
            .Parent.Height = 200
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
            .SeriesCollection(1).Interior.Color = RGB(4, 49, 76)
            .ChartTitle.Text = left(ws.Range("A1").value, 150) & " [Median]"
            .Parent.top = 210
            .Parent.left = 230
        End With
        
    Else
        Set rng = ws.UsedRange
        Set rng = rng.Offset(1, 0).Resize(rng.Rows.count - 1)
        Set cat_chart = ws.Shapes.AddChart2

        With cat_chart.Chart
            .SetSourceData rng
            .ChartType = xlColumnClustered
            .PlotBy = xlColumns
            .SetElement (msoElementDataLabelOutSideEnd)
            .Parent.Width = 22 * last_row + 30 * last_col
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
            .SeriesCollection(1).Interior.Color = RGB(4, 49, 76)
            .ChartTitle.Text = left(ws.Range("A1").value, 150) & " [Percentage]"
            .Parent.Height = 300
            .ApplyLayout (3)
    
            If last_col > 4 Then
                .Parent.top = ws.Range("A" & last_row + 2).top
                .Parent.left = 0
            Else
                .Parent.top = 0
                .Parent.left = ws.Cells(2, last_col + 1).left + 10
            End If
        End With
        
    End If
End Sub

Private Sub arrange_table(ws As Worksheet)
    Application.ScreenUpdating = False
    Dim t As Double
    Dim dis_level As String
    Dim last_row As Long
    Dim last_col As Long
    Dim last_col_letter As String
    Dim i As Long
    Dim rng As Range
    
    dis_level = Me.ComboBoxDis.value
    last_row = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    last_col = ws.Cells(3, Columns.count).End(xlToLeft).Column

    For i = last_row To 5 Step -1
        If ws.Cells(i, 1).value <> dis_level Then
            ws.Rows(i).Delete
        End If
    Next i
    If ws.Range("C3") = "percentage" Then
        ws.Range("B1") = ws.Range("C3")
        ws.Range("B2") = ws.Range("A5")
        ws.Rows("3:4").Delete
        NUMERIC_CHART = False
    Else
        ws.Range("B3") = ws.Range("A5")
        ws.Rows("2").Delete
        ws.Rows("3").Delete
        NUMERIC_CHART = True
    End If
    
    ws.Columns(1).Delete
    ws.Range("A1") = ""
    last_col_letter = number_to_letter(last_col - 1, ws)
    ws.Range("A1:" & last_col_letter & 1).Merge
    If last_col >= 10 Then
        ws.Range("A:" & last_col_letter).ColumnWidth = 12
    Else
        ws.Range("A:" & last_col_letter).ColumnWidth = 14
    End If

    With ws.Range("A1:" & last_col_letter & 1)
        .Interior.Color = RGB(170, 170, 170)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    
    With ws.Range("A2")
        .Interior.Color = RGB(170, 170, 170)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
    
    With ws.Range("B2:" & last_col_letter & 2)
        .Interior.Color = RGB(220, 220, 220)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
       
    ws.Cells.Font.Size = 9
    ws.Rows("1:1").Font.Size = 10
    ws.Rows("1:1").RowHeight = 27
    
    If NUMERIC_CHART Then
        ws.Columns("A:C").ColumnWidth = 12
    Else
        ws.Rows("2:2").RowHeight = 33
    End If
    
    ws.Range("B2:" & last_col_letter & 2).WrapText = True
    ws.Range("A1").Font.Bold = True
    
    Set rng = ws.UsedRange
    
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
End Sub


Private Sub TextBoxSearch_Change()
    Dim filterText As String
    Dim originalItems As Variant
    Dim filteredRows() As Variant
    Dim filteredItems() As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim rowIdx As Long
    Dim v As Variant
    Dim cell As Range
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim filtered_rng As Range
    
    Set ws = sheets("indi_list")
    Set rng = ws.Range("A1").CurrentRegion
    
    filterText = LCase(TextBoxSearch.Text)
    originalItems = rng
    ListBoxVars.RowSource = ""
    ws.Columns("D:E").Clear
    ListBoxVars.Clear
    
    ' Filter the items based on the filter text
    j = 1
    For i = LBound(originalItems) To UBound(originalItems)
        If InStr(1, LCase(originalItems(i, 1)), filterText) > 0 Or _
            InStr(1, LCase(originalItems(i, 2)), filterText) > 0 Or filterText = "" Then
            ReDim Preserve filteredRows(1 To j)
            filteredRows(j) = i
            j = j + 1
        End If
    Next i
    
    If Not Not filteredRows Then
        k = 1
        For Each v In filteredRows
            ws.Cells(k, "D") = rng.Cells(v, 1)
            ws.Cells(k, "E") = rng.Cells(v, 2)
            k = k + 1
        Next v
        
        Set filtered_rng = ws.Range("D1").CurrentRegion
        
        For Each cell In filtered_rng.Rows
            Me.ListBoxVars.AddItem
            rowIdx = Me.ListBoxVars.ListCount - 1
            Me.ListBoxVars.List(rowIdx, 0) = cell.Cells(1, 1).value
            Me.ListBoxVars.List(rowIdx, 1) = cell.Cells(1, 2).value
        Next cell
        
    End If

End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim ws As Worksheet
    Dim rng As Range
    Dim dis_rng As Range
    Dim cell As Range
    Dim rowIdx As Integer
    
    If Not worksheet_exists("dm_backend") Or Not worksheet_exists("indi_list") Then
        MsgBox "Please first analyze the data, then try to generate charts. ", vbInformation
        Unload single_chart_form
        Exit Sub
    End If
    
    Set ws = sheets("indi_list")
    Set rng = ws.Range("A1").CurrentRegion
    Me.TextBoxSearch.BorderStyle = 1
    With Me.ListBoxVars
        .BorderStyle = 1
        .columnWidths = "100,170"
    End With
    
    For Each cell In rng.Rows
        Me.ListBoxVars.AddItem
        rowIdx = Me.ListBoxVars.ListCount - 1
        Me.ListBoxVars.List(rowIdx, 0) = cell.Cells(1, 1).value
        Me.ListBoxVars.List(rowIdx, 1) = cell.Cells(1, 2).value
    Next cell
    
    Set dis_rng = ws.Range("G1").CurrentRegion
    
    For Each cell In dis_rng
        Me.ComboBoxDis.AddItem cell.value
    Next cell
    
    For Each cell In dis_rng
        If cell.value <> "ALL" Then
            Me.ComboBoxDis.value = cell.value
            Exit For
        End If
        
    Next cell
    
End Sub









