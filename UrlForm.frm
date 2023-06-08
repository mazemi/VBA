VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UrlForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2580
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   4536
   OleObjectBlob   =   "UrlForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UrlForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandSave_Click()
    user = SaveRegistrySetting("ramSetting", "koboAuditReg", Me.TextAudit.Value)
    Password = SaveRegistrySetting("ramSetting", "koboPhotoReg", Me.TextPhoto.Value)
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.TextAudit.Value = GetRegistrySetting("ramSetting", "koboAuditReg")
    Me.TextPhoto.Value = GetRegistrySetting("ramSetting", "koboPhotoReg")
End Sub
