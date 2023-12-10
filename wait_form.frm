VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} wait_form 
   ClientHeight    =   930
   ClientLeft      =   -108
   ClientTop       =   -474
   ClientWidth     =   3828
   OleObjectBlob   =   "wait_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "wait_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
End Sub


