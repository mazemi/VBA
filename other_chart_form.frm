VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} other_chart_form 
   Caption         =   "Generate Charts"
   ClientHeight    =   6096
   ClientLeft      =   96
   ClientTop       =   432
   ClientWidth     =   7554
   OleObjectBlob   =   "other_chart_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "other_chart_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub ComboBoxDis_Change()
    On Error Resume Next
    Dim ws_res As Worksheet
    Dim ws As Worksheet
    Dim rng As Range
    Dim option_rng As Range
    Dim dis_rng As Range
    Dim cell As Range
    Dim last_row As Long
    Dim rowIdx As Integer

    Set ws_res = sheets("result")
    Set ws = sheets("indi_list")
    
    ws.Range("K:L").Clear
    ws.Range("I1") = "disaggregation"
    ws.Range("K1") = "disaggregation value"
    ws.Range("L1") = "disaggregation label"

    Set rng = ws_res.Range("A1").CurrentRegion
    Set dis_rng = ws.Range("G1").CurrentRegion
    ws.Range("I2") = "'=" & Me.ComboBoxDis

    rng.AdvancedFilter xlFilterCopy, ws.Range("I1:I2"), ws.Range("K1").CurrentRegion, True
    last_row = ws.Cells(ws.Rows.count, "K").End(xlUp).Row
    Set option_rng = ws.Range("K2:L" & last_row)
    
    Call sort_disaggregation
    
    Me.ListBoxVars.Clear
    With Me.ListBoxVars
        .BorderStyle = 1
        .columnWidths = "120,100"
    End With
    
    For Each cell In option_rng.Rows
        Me.ListBoxVars.AddItem
        rowIdx = Me.ListBoxVars.ListCount - 1
        Me.ListBoxVars.List(rowIdx, 0) = cell.Cells(1, 1).value
        Me.ListBoxVars.List(rowIdx, 1) = cell.Cells(1, 2).value
    Next cell
End Sub

Private Sub CommandCancel_Click()
    Unload Me
    chart_form.Show
End Sub

Private Sub CommandNext_Click()
    Dim selected_indexes As Collection
    Dim selected_values As New Collection
    Dim selected_labels As New Collection
    Dim index As Variant

    If Me.ComboBoxDis.value <> "" And count_selected_items(Me.ListBoxVars) > 0 Then
        
        Set selected_indexes = get_selected_items
        
        For Each index In selected_indexes
            selected_values.Add Me.ListBoxVars.List(index, 0)
            selected_labels.Add Me.ListBoxVars.List(index, 1)
        Next index
    ElseIf count_selected_items(Me.ListBoxVars) = 0 Then
        MsgBox "Please select up to three disaggregation values from the list and then click 'Generate Chart'.", vbInformation
    End If
    Unload Me
    Call generate_multiple_data_chart(Me.ComboBoxDis.value, selected_values, selected_labels)
    
End Sub

Function count_selected_items(list_box As MSForms.listBox) As Integer
    Dim count As Integer
    Dim i As Integer
    
    count = 0
    
    For i = 0 To list_box.ListCount - 1
        If list_box.Selected(i) Then
            count = count + 1
        End If
    Next i
    
    count_selected_items = count
End Function

Function get_selected_items() As Collection
    Dim selected_indexes As New Collection
    Dim i As Integer
    Dim maximum_seleted As Integer
    Dim seleted_options As Integer
    Dim list_box As MSForms.listBox
    
    Set list_box = Me.ListBoxVars
    
    ' for the sake of performance "maximum_seleted" variable is limitted to 3
    maximum_seleted = 3
    
    For i = 0 To list_box.ListCount - 1
        If list_box.Selected(i) And seleted_options < maximum_seleted Then
            seleted_options = seleted_options + 1
            selected_indexes.Add i
        End If
    Next i
    
    Set get_selected_items = selected_indexes

End Function


Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim ws_res As Worksheet
    Dim ws As Worksheet
    Dim rng As Range
    Dim option_rng As Range
    Dim dis_rng As Range
    Dim cell As Range
    Dim last_row As Long
    Dim rowIdx As Integer
    
    If Not worksheet_exists("result") Or Not worksheet_exists("dm_backend") Or Not worksheet_exists("indi_list") Then
        MsgBox "Please first analyze the data, then try to generate charts. ", vbInformation
        Unload other_chart_form
        Exit Sub
    End If
    
    Set ws_res = sheets("result")
    Set ws = sheets("indi_list")
    
    ws.Range("K:L").Clear
    
    ws.Range("I1") = "disaggregation"
    ws.Range("K1") = "disaggregation value"
    ws.Range("L1") = "disaggregation label"

    Set rng = ws_res.Range("A1").CurrentRegion
    Set dis_rng = ws.Range("G1").CurrentRegion
    
    For Each cell In dis_rng
        If cell.value <> "ALL" Then
            Me.ComboBoxDis.AddItem cell.value
        End If
    Next cell
    
    For Each cell In dis_rng
        If cell.value <> "ALL" Then
            Me.ComboBoxDis.value = cell.value
            ws.Range("I2") = "'=" & cell.value
            Exit For
        End If
    Next cell
    
    rng.AdvancedFilter xlFilterCopy, ws.Range("I1:I2"), ws.Range("K1").CurrentRegion, True
    last_row = ws.Cells(ws.Rows.count, "K").End(xlUp).Row
    Set option_rng = ws.Range("K2:L" & last_row)
    
    Call sort_disaggregation
    
    Me.ListBoxVars.Clear
    With Me.ListBoxVars
        .BorderStyle = 1
        .columnWidths = "120,100"
    End With
    
    For Each cell In option_rng.Rows
        Me.ListBoxVars.AddItem
        rowIdx = Me.ListBoxVars.ListCount - 1
        Me.ListBoxVars.List(rowIdx, 0) = cell.Cells(1, 1).value
        Me.ListBoxVars.List(rowIdx, 1) = cell.Cells(1, 2).value
    Next cell
    
End Sub

Sub sort_disaggregation()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sortRange As Range

    Set ws = sheets("indi_list")
    lastRow = ws.Cells(ws.Rows.count, "K").End(xlUp).Row
    Set sortRange = ws.Range("K2:L" & lastRow)
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range("K2:K" & lastRow), Order:=xlAscending
        .SetRange sortRange
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

