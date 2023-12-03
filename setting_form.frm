VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} setting_form 
   Caption         =   "Setting"
   ClientHeight    =   5058
   ClientLeft      =   -54
   ClientTop       =   -258
   ClientWidth     =   7914
   OleObjectBlob   =   "setting_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "setting_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandSave_Click()
    On Error Resume Next
    Dim current_dt_name As String
    Dim new_name As String
    
    If Me.ComboData.Value <> vbNullString Then
        current_dt_name = Me.ComboData.Value
        new_name = alpha_numeric_only(current_dt_name)
        
        If Len(new_name) > 15 Then
            new_name = left(new_name, 15)
        End If
        
        Public_module.DATA_SHEET = new_name
        dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", new_name)
        sheets(current_dt_name).name = new_name
        
    Else
        Public_module.DATA_SHEET = Me.ComboData.Value
        dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", Me.ComboData.Value)
    End If
    
    ' save to registry:
    user = SaveRegistrySetting("ramSetting", "koboUserReg", Me.TextUser.Value)
    Password = SaveRegistrySetting("ramSetting", "koboPasswordReg", Me.TextPassword.Value)
    audit = SaveRegistrySetting("ramSetting", "koboAuditReg", Me.ComboAudit.Value)
    
    Unload Me
    
End Sub

Private Sub LabelReset_Click()
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Dim answer As Integer
    answer = MsgBox("All setting values, tool and cleaning plan will be removed." & vbCrLf & _
                    "Do you want to Continue?", vbQuestion + vbYesNo)
    
    If answer = vbYes Then
        user = SaveRegistrySetting("ramSetting", "koboUserReg", vbNullString)
        Password = SaveRegistrySetting("ramSetting", "koboPasswordReg", vbNullString)
        dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", vbNullString)
        audit = SaveRegistrySetting("ramSetting", "koboAuditReg", vbNullString)
        kobo_tool = SaveRegistrySetting("ramSetting", "koboToolReg", vbNullString)
        
        dt = SaveRegistrySetting("ramSetting", "dataReg", vbNullString)
        smp = SaveRegistrySetting("ramSetting", "samplingReg", vbNullString)
        dt_strata = SaveRegistrySetting("ramSetting", "dataStrataReg", vbNullString)
        smp_strata = SaveRegistrySetting("ramSetting", "samplingStrataReg", vbNullString)
        smp_population = SaveRegistrySetting("ramSetting", "samplingPopulationReg", vbNullString)
        
        If worksheet_exists("keen") Then
            sheets("keen").Visible = xlSheetHidden
            sheets("keen").Delete
        End If
        
        If worksheet_exists("keen2") Then
            sheets("keen2").Delete
        End If
    
        If worksheet_exists("temp_sheet") Then
            sheets("temp_sheet").Visible = xlSheetHidden
            sheets("temp_sheet").Delete
        End If
        
        If worksheet_exists("redeem") Then
            sheets("redeem").Visible = xlSheetHidden
            sheets("redeem").Delete
        End If
        
        If worksheet_exists("dissagregation_setting") Then
            sheets("dissagregation_setting").Visible = xlSheetHidden
            sheets("dissagregation_setting").Delete
        End If
         
        If worksheet_exists("indi_list") Then
            sheets("indi_list").Visible = xlSheetHidden
            sheets("indi_list").Delete
        End If
        
        If worksheet_exists("analysis_list") Then
            sheets("analysis_list").Delete
        End If
    
        ThisWorkbook.sheets("xsurvey").Cells.Clear
        ThisWorkbook.sheets("xchoices").Cells.Clear
        ThisWorkbook.sheets("xsurvey_choices").Cells.Clear
        ThisWorkbook.sheets("xlogical_checks").Cells.Clear
        
        Me.tooLabel.Caption = ""
    End If
    
    Application.StatusBar = False
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
     
    Call UserForm_Initialize
    Me.ComboData = ""
End Sub

Private Sub UserForm_Initialize()
    Dim dt_sheet As String
    Me.Label_import.Visible = False
    Me.TextUser.Value = GetRegistrySetting("ramSetting", "koboUserReg")
    Me.TextPassword.Value = GetRegistrySetting("ramSetting", "koboPasswordReg")
    dt_sheet = GetRegistrySetting("ramSetting", "dataReg")
    Me.ComboAudit.Value = GetRegistrySetting("ramSetting", "koboAuditReg")
    
    If worksheet_exists(dt_sheet) Then
        Me.ComboData.Value = dt_sheet
    End If
    
    If ThisWorkbook.sheets("xsurvey").Range("A1") <> vbNullString Then
        Me.tooLabel = "Integrated Tool: " & vbCrLf & GetRegistrySetting("ramSetting", "koboToolReg")
    End If
    
    Call PopulateComboBox
    
    Me.LabelVersion.Caption = "version " & VERSION
End Sub

Private Sub PopulateComboBox()
    On Error Resume Next
    Dim header_arr() As Variant
    Dim filtered_arr() As String
    Dim ws As Worksheet
    
    Set ws = ActiveWorkbook.ActiveSheet

    header_arr = ws.Range(ws.Cells(1, 1), ws.Cells(1, 1).End(xlToRight)).Value2
    
    With Application
        header_arr = .Transpose(.Transpose(header_arr))
    End With

    filtered_arr = Filter(header_arr, "URL", True, vbTextCompare)
    
    Me.ComboAudit.List = filtered_arr
    
    Dim sheet_li As Collection
    
    'Get the collection of worksheet names
    Set sheet_li = sheet_list
    
    'name of a sheet
    Dim sh As Variant
    
    For Each sh In sheet_li
        If ActiveWorkbook.Worksheets(CStr(sh)).Visible Then
            If CStr(sh) <> "result" And CStr(sh) <> "log_book" And CStr(sh) <> "analysis_list" And _
               CStr(sh) <> "dissagregation_setting" And CStr(sh) <> "overall" And CStr(sh) <> "survey" And _
                CStr(sh) <> "keen" And CStr(sh) <> "indi_list" And CStr(sh) <> "temp_sheet" And _
                CStr(sh) <> "choices" And CStr(sh) <> "datamerge" Then
                Me.ComboData.AddItem sh
            End If
        End If
    Next sh
    
End Sub

Private Sub CommandTools_Click()
    
    On Error GoTo errHandler
    
    Application.ScreenUpdating = False
    
    Dim objFSO As New FileSystemObject
    Dim FileSelected As String

    Set myFile = Application.FileDialog(msoFileDialogOpen)
    With myFile
        .title = "Choose File"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Exit Sub
        End If
        FileSelected = .SelectedItems(1)
    End With

    Me.Label_import.Visible = True
    DoEvents
    
    Call import_survey(FileSelected)
    Call import_choices(FileSelected)
    Call make_survey_choice

    Me.Label_import.Caption = "imported"
    
    Application.ScreenUpdating = True
    
    tool_path = SaveRegistrySetting("ramSetting", "koboToolReg", FileSelected)
    
    Me.tooLabel = "Integrated Tool: " & vbCrLf & GetRegistrySetting("ramSetting", "koboToolReg")
    
    ActiveWorkbook.Save
        
    Exit Sub

errHandler:
        Me.Label_import.Visible = False
'        Unload wait_form
        MsgBox "There is an issue!   " & vbCrLf & _
        "Please select a valid KOBO tool.   ", vbCritical
        Application.ScreenUpdating = True
End Sub




