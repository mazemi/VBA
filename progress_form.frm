VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} progress_form 
   Caption         =   "Progress"
   ClientHeight    =   1662
   ClientLeft      =   -426
   ClientTop       =   -2004
   ClientWidth     =   6048
   OleObjectBlob   =   "progress_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "progress_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    If CloseMode = vbFormControlMenu Then
        End
    End If
End Sub


