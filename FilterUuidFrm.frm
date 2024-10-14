VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FilterUuidFrm 
   Caption         =   "Filter uuid"
   ClientHeight    =   3876
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   4668
   OleObjectBlob   =   "FilterUuidFrm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FilterUuidFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandFilter_Click()
    Dim ws As Worksheet
    Dim uuidColumnIndex As Long
    Dim filterValues As Variant
    Dim textBoxContent As String
    Dim filterCriteria As String
    Dim lastRow As Long
    Dim lastCol As Long
    
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    On Error Resume Next
    uuidColumnIndex = Application.Match("_uuid", ws.Rows(1), 0)
    On Error GoTo 0
    
    If uuidColumnIndex = 0 Then
        MsgBox "The _uuid column not found.            ", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    textBoxContent = Me.Textuuid.Text
    
    If textBoxContent = "" Then
        Exit Sub
    End If
    
    filterValues = Split(textBoxContent, vbNewLine)
    
    filterCriteria = Join(filterValues, ",")

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
    
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter _
        Field:=uuidColumnIndex, Criteria1:=Split(filterCriteria, ","), Operator:=xlFilterValues
    
    Application.ScreenUpdating = True
End Sub

Private Sub CommandPaste_Click()
    Dim cell As Range
    Dim textBoxContent As String
    
    textBoxContent = ""

    For Each cell In Selection
        textBoxContent = textBoxContent & cell.Value & vbNewLine
    Next cell

    Me.Textuuid.Text = textBoxContent
End Sub
