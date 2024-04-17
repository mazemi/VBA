VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} select_data_form 
   Caption         =   "Select Data"
   ClientHeight    =   1710
   ClientLeft      =   -300
   ClientTop       =   -1338
   ClientWidth     =   4836
   OleObjectBlob   =   "select_data_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "select_data_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub CommandAdd_Click()
    On Error Resume Next
    dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", Me.ComboSheets.value)
    
    If Me.ComboSheets.value = "" Then
        End
    End If
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next

    With Me
        .StartUpPosition = 0
        .left = Application.left + (0.5 * Application.Width) - (0.5 * .Width)
        .top = Application.top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
    Dim dt As String
    Dim sheet_li As Collection
    Dim sh As Variant
    
    dt = ""
    dt = GetRegistrySetting("ramSetting", "dataReg")

    If dt <> "" Then
        If worksheet_exists(dt) Then
            Me.ComboSheets.Text = dt
        End If
    End If
 
    Set sheet_li = sheet_list

    For Each sh In sheet_li
        If ActiveWorkbook.Worksheets(CStr(sh)).Visible And Not IsInArray(CStr(sh), InitializeExcludedSheets) And left(CStr(sh), 6) <> "chart-" Then
            Me.ComboSheets.AddItem sh
        End If
    Next sh
    
End Sub


