VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} data_checking_form 
   Caption         =   "Data Checking"
   ClientHeight    =   3348
   ClientLeft      =   -72
   ClientTop       =   -396
   ClientWidth     =   6636
   OleObjectBlob   =   "data_checking_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "data_checking_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandRun_Click()
    If Me.OptionWrongValue = True Then
        issueText = "Wrong value"
    ElseIf Me.OptionOutlier = True Then
        issueText = "Outlier"
    ElseIf Me.OptionHarmonization = True Then
        issueText = "Harmonization"
    Else
        issueText = Me.TextOther.Value
        res = SaveRegistrySetting("ramSetting", "issueTextReg", issueText)
    End If
    
    patternCheckAction = True
    Unload Me
End Sub


Private Sub OptionBlank_Click()
    Me.TextOther.Enabled = True
End Sub

Private Sub OptionHarmonization_Click()
    Me.TextOther.Enabled = False
End Sub

Private Sub OptionOutlier_Click()
    Me.TextOther.Enabled = False
End Sub

Private Sub OptionWrongValue_Click()
    Me.TextOther.Enabled = False
End Sub

Private Sub UserForm_Initialize()
    patternCheckAction = False
    Me.TextOther.Enabled = False
    Me.OptionWrongValue = True
    issueText = "Wrong value"

    Me.TextOther.Value = GetRegistrySetting("ramSetting", "issueTextReg")
    
End Sub
