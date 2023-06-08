VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} koboForm 
   Caption         =   "KOBO Account"
   ClientHeight    =   2610
   ClientLeft      =   -60
   ClientTop       =   -258
   ClientWidth     =   6162
   OleObjectBlob   =   "koboForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "koboForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandSave_Click()
    user = SaveRegistrySetting("ramSetting", "koboUserReg", Me.TextUser.Value)
    Password = SaveRegistrySetting("ramSetting", "koboPasswordReg", Me.TextPassword.Value)
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.TextUser.Value = GetRegistrySetting("ramSetting", "koboUserReg")
    Me.TextPassword.Value = GetRegistrySetting("ramSetting", "koboPasswordReg")
End Sub

