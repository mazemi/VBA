VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} empty_col_form 
   Caption         =   "Empty Columns"
   ClientHeight    =   3930
   ClientLeft      =   -294
   ClientTop       =   -1338
   ClientWidth     =   6390
   OleObjectBlob   =   "empty_col_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "empty_col_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    On Error Resume Next
    
    With Me
        .StartUpPosition = 0
        .left = Application.left + (0.5 * Application.Width) - (0.5 * .Width)
        .top = Application.top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    Application.DisplayAlerts = False
    sheets("temp_sheet").Visible = xlSheetHidden
    sheets("temp_sheet").Delete
    Application.DisplayAlerts = True
End Sub

Private Sub ListBoxEmptyCols_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = sheets(find_main_data)
    Dim col_name As String
    Dim col_number As Long
    
    col_name = Me.ListBoxEmptyCols.value
    col_number = letter_to_number(col_name, ws)
    ws.Activate
    
    ActiveWindow.ScrollRow = 1
    If col_number > 2 Then
        ActiveWindow.ScrollColumn = col_number - 2
    ElseIf col_number = 2 Then
        ActiveWindow.ScrollColumn = col_number - 1
    ElseIf col_number = 1 Then
        ActiveWindow.ScrollColumn = col_number
    End If
    
    ws.Columns(col_number).Activate
    
End Sub

Private Sub remove_empty_col_command_Click()
    Dim lb As MSForms.ListBox
    Dim listBoxArray() As Variant
    Dim i As Integer
    
    Set lb = Me.ListBoxEmptyCols
    
    ReDim listBoxArray(1 To lb.ListCount)
    
    For i = 1 To lb.ListCount
        listBoxArray(i) = lb.List(i - 1, 0)
    Next i
    
    For i = UBound(listBoxArray) To LBound(listBoxArray) Step -1
        Call DeleteColumnByName(listBoxArray(i))
    Next i
    
    Unload Me
    
    MsgBox "Empty columns have been removed.   ", vbInformation
    
End Sub


Sub DeleteColumnByName(ByVal columnName As String)
    Dim ws As Worksheet
    Dim col As Range
    Set ws = sheets(find_main_data)
    ws.Columns(columnName).Delete
End Sub

