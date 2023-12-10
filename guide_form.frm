VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} guide_form 
   Caption         =   "Quick Guide"
   ClientHeight    =   4692
   ClientLeft      =   -102
   ClientTop       =   -474
   ClientWidth     =   6846
   OleObjectBlob   =   "guide_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "guide_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub get_version()
    Dim ie As Object
    Dim url As String
    Dim myPoints As String

    url = "https://mazemi.github.io/testok/"
    Set ie = CreateObject("InternetExplorer.Application")

    With ie
      .Visible = 0
      .Navigate url
       While .Busy Or .ReadyState <> 4
         DoEvents
       Wend
    End With
    Debug.Print 1
    Dim Doc As HTMLDocument
    Set Doc = ie.Document

    version_value = Trim(Doc.getElementsByName("ram-version")(0).Value)
    Debug.Print myPoints
    Me.LabelVersion.Caption = "Latest version: " & version_value
End Sub

Private Sub CommandButton1_Click()
    Call get_version
End Sub

Private Sub UserForm_Initialize()
    Me.LabelCurrentVersion.Caption = "Currenet Version: " & Info_module.VERSION
End Sub
