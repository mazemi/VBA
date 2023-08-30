VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} select_data_form 
   Caption         =   "Select Data"
   ClientHeight    =   1746
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   4782
   OleObjectBlob   =   "select_data_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "select_data_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandAdd_Click()
    dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", Me.ComboSheets.value)
    
    If Me.ComboSheets.value = "" Then
        End
    End If
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    Dim dt As String
    dt = ""
    dt = GetRegistrySetting("ramSetting", "dataReg")

    If dt <> "" Then
        If worksheet_exists(dt) Then
            Me.ComboSheets.Text = dt
        End If
    End If
    
    Dim sheet_li As Collection
    Set sheet_li = sheet_list                    'Get the collection of worksheet names
    Dim sh As Variant                            'name of a sheet
    For Each sh In sheet_li
        If ActiveWorkbook.Worksheets(CStr(sh)).visible Then
            If CStr(sh) <> "result" And CStr(sh) <> "log_book" And CStr(sh) <> "analysis_list" And _
               CStr(sh) <> "dissagregation_setting" And CStr(sh) <> "overall" And CStr(sh) <> "survey" And _
                CStr(sh) <> "keen" And CStr(sh) <> "indi_list" And CStr(sh) <> "temp_sheet" And _
               CStr(sh) <> "choices" And CStr(sh) <> "xsurvey_choices" And CStr(sh) <> "datamerge" Then
                Me.ComboSheets.AddItem sh
            End If
        End If
    Next sh
    
End Sub


