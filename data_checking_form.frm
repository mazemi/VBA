VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} data_checking_form 
   Caption         =   "Data Checking"
   ClientHeight    =   3408
   ClientLeft      =   -564
   ClientTop       =   -2550
   ClientWidth     =   6708
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
        Public_module.ISSUE_TEXT = "Wrong value"
    ElseIf Me.OptionOutlier = True Then
        Public_module.ISSUE_TEXT = "Outlier"
    ElseIf Me.OptionHarmonization = True Then
        Public_module.ISSUE_TEXT = "Translation and Harmonization"
    Else
        Public_module.ISSUE_TEXT = Me.TextOther.value
        res = SaveRegistrySetting("ramSetting", "issueTextReg", Public_module.ISSUE_TEXT)
    End If
    
    Public_module.PATTERN_CHECK_ACTION = True
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

    With Me
        .StartUpPosition = 0
        .left = Application.left + (0.5 * Application.Width) - (0.5 * .Width)
        .top = Application.top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
    Public_module.PATTERN_CHECK_ACTION = False
    Me.TextOther.Enabled = False
    Me.OptionWrongValue = True
    Public_module.ISSUE_TEXT = "Wrong value"

    Me.TextOther.value = GetRegistrySetting("ramSetting", "issueTextReg")
    
End Sub

