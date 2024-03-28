VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} find_form 
   Caption         =   "Find The Indicator"
   ClientHeight    =   5988
   ClientLeft      =   -330
   ClientTop       =   -1560
   ClientWidth     =   10476
   OleObjectBlob   =   "find_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "find_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit




Private Sub ListBoxVars_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'    On Error Resume Next
    Dim i As Long
    
    With Me.ListBoxVars
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Call CommandGo_Click
                Exit For
            End If
        Next
    End With
End Sub

Private Sub CommandGo_Click()
    Dim item As String
    Dim item_index As Long

    For item_index = 0 To Me.ListBoxVars.ListCount - 1
        If ListBoxVars.Selected(item_index) = True Then
            item = ListBoxVars.List(item_index, 1)
        End If
    Next

    If item = vbNullString Then Exit Sub
    
    On Error Resume Next
    If Len(item) < 120 Then
        Cells.Find(What:=item, after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Activate
    
    Else
        item = left(item, 100)
        Cells.Find(What:=item, after:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    
    End If

End Sub

Private Sub TextBoxSearch_Change()
    Dim filterText As String
    Dim originalItems As Variant
    Dim filteredRows() As Variant
    Dim filteredItems() As String
    Dim i As Long
    Dim j As Long
    Dim v As Variant
    Dim cell As Range
    Dim k As Long
    Dim rowIdx As Long
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

    With Me
        .StartUpPosition = 0
        .left = Application.left + (0.5 * Application.Width) - (0.5 * .Width)
        .top = Application.top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
    If Not worksheet_exists("datamerge") And Not worksheet_exists("overall") Then
        MsgBox "The datamerge or overall sheets do not exist!   ", vbInformation
        Unload Me
        Exit Sub
    End If

    If Not worksheet_exists("indi_list") Then
        MsgBox "The indicators dose not exist!   ", vbInformation
        Unload Me
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim last_row As Long
    Dim arr() As Variant

    With Me.ListBoxVars
        .BorderStyle = 1
        .columnWidths = "100,170"
    End With
    
    Set ws = sheets("indi_list")
    
    Me.ListBoxVars.List = ws.Range("A1").CurrentRegion.value
    
End Sub
