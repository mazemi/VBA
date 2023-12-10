VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} find_form 
   Caption         =   "Find The Indicator"
   ClientHeight    =   5316
   ClientLeft      =   -60
   ClientTop       =   -264
   ClientWidth     =   7182
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



Private Sub ListBoxIndicator_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim i As Long
    
    With Me.ListBoxIndicator
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

    For item_index = 0 To ListBoxIndicator.ListCount - 1
        If ListBoxIndicator.Selected(item_index) = True Then
            item = ListBoxIndicator.List(item_index)
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

Private Sub UserForm_Initialize()
    On Error Resume Next
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

    Set ws = sheets("indi_list")
    
    Me.ListBoxIndicator.List = ws.Range("A1").CurrentRegion.Value
    
End Sub
