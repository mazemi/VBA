VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} setting_form 
   Caption         =   "Setting"
   ClientHeight    =   4962
   ClientLeft      =   -222
   ClientTop       =   -1032
   ClientWidth     =   7932
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
    
    If Me.ComboData.value <> vbNullString Then
        current_dt_name = Me.ComboData.value
        new_name = alpha_numeric_only(current_dt_name)
        
        If Len(new_name) > 15 Then
            new_name = left(new_name, 15)
        End If
        
        Public_module.DATA_SHEET = new_name
        dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", new_name)
        sheets(current_dt_name).Name = new_name
        
    Else
        Public_module.DATA_SHEET = Me.ComboData.value
        dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", Me.ComboData.value)
    End If
    
    ' save to registry:
    user = SaveRegistrySetting("ramSetting", "koboUserReg", Me.TextUser.value)
    Password = SaveRegistrySetting("ramSetting", "koboPasswordReg", Me.TextPassword.value)
    audit = SaveRegistrySetting("ramSetting", "koboAuditReg", Me.ComboAudit.value)
    
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
            sheets("keen2").Visible = xlSheetHidden
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
        
        If worksheet_exists("disaggregation_setting") Then
            sheets("disaggregation_setting").Visible = xlSheetHidden
            sheets("disaggregation_setting").Delete
        End If
         
        If worksheet_exists("indi_list") Then
            sheets("indi_list").Visible = xlSheetHidden
            sheets("indi_list").Delete
        End If
        
        If worksheet_exists("analysis_list") Then
            sheets("analysis_list").Delete
        End If
    
        If worksheet_exists("result") Then
            sheets("result").Delete
        End If
        
        If worksheet_exists("dm_backend") Then
            sheets("dm_backend").Delete
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
    
    Me.ComboData = ""
    Call UserForm_Initialize

End Sub

Private Sub UserForm_Initialize()
    Dim dt_sheet As String
    Me.Label_import.Visible = False
    Me.TextUser.value = GetRegistrySetting("ramSetting", "koboUserReg")
    Me.TextPassword.value = GetRegistrySetting("ramSetting", "koboPasswordReg")
    dt_sheet = GetRegistrySetting("ramSetting", "dataReg")
    Me.ComboAudit.value = GetRegistrySetting("ramSetting", "koboAuditReg")
    
    If worksheet_exists(dt_sheet) Then
        Me.ComboData.value = dt_sheet
    End If
    
    If ThisWorkbook.sheets("xsurvey").Range("A1") <> vbNullString Then
        Me.tooLabel = "Integrated Tool: " & vbCrLf & GetRegistrySetting("ramSetting", "koboToolReg")
    End If
    
    Call PopulateComboBox
    
    Me.LabelVersion.Caption = "version " & VERSION
End Sub


Private Sub PopulateComboBox()
    On Error Resume Next
    Dim ws As Worksheet
    Dim sheet_li As Collection
    Dim sh As Variant
    Dim header_arr() As Variant
    Dim filtered_arr() As String
    
    Set ws = ActiveWorkbook.ActiveSheet

    header_arr = ws.Range(ws.Cells(1, 1), ws.Cells(1, 1).End(xlToRight)).Value2
    
    With Application
        header_arr = .Transpose(.Transpose(header_arr))
    End With

    filtered_arr = Filter(header_arr, "URL", True, vbTextCompare)
    
    Me.ComboAudit.List = filtered_arr
    
    Set sheet_li = sheet_list
    
    For Each sh In sheet_li
        If ActiveWorkbook.Worksheets(CStr(sh)).Visible And Not IsInArray(CStr(sh), InitializeExcludedSheets) And left(CStr(sh), 6) <> "chart-" Then
            Me.ComboData.AddItem sh
        End If
    Next sh
    
End Sub


Private Sub CommandTools_Click()
    
    On Error GoTo errhandler
    Me.Label_import.Visible = False
    Application.ScreenUpdating = False
    
    Dim objFSO As New FileSystemObject
    Dim FileSelected As String

    Set myFile = Application.FileDialog(msoFileDialogOpen)
    With myFile
        .title = "Choose File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx"
        If .Show <> -1 Then
            Exit Sub
        End If
        FileSelected = .SelectedItems(1)
    End With
    
 Debug.Print FileSelected
    DoEvents
    Me.bar.Width = 10
    Me.bar.Visible = True
    
    If Not check_tools_file(FileSelected) Then
        MsgBox "Something went wrong!   " & vbCrLf & _
        "Please select a valid KOBO tool with survey and choices sheets.   ", vbCritical
        Me.bar.Width = 0
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Call import_survey(FileSelected)
    DoEvents
    Me.bar.Width = 30
    Call import_choices(FileSelected)
    DoEvents
    Me.bar.Width = 40
    Call make_survey_choice
    DoEvents
    Me.bar.Width = 63
    
    Application.Wait (Now + 0.00001)
    Me.bar.Visible = False
    Me.Label_import.Visible = True
    
    tool_path = SaveRegistrySetting("ramSetting", "koboToolReg", FileSelected)
    Me.tooLabel = "Integrated Tool: " & vbCrLf & GetRegistrySetting("ramSetting", "koboToolReg")
    
    ActiveWorkbook.Save
    Me.Label_import.Caption = "imported"
    Me.bar.Visible = False
    Application.ScreenUpdating = True
    Exit Sub

errhandler:
        Me.Label_import.Visible = False
        MsgBox "Something went wrong!   " & vbCrLf & _
        "Please select a valid KOBO tool.   ", vbCritical
        Me.bar.Width = 0
        Application.ScreenUpdating = True
End Sub

Private Function check_tools_file(filePath As String) As Boolean
    Dim excelApp As Object
    Dim closedWorkbook As Object
    Dim sheet As Object
    Dim has_survey As Boolean
    Dim has_choices As Boolean
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    
    Set closedWorkbook = excelApp.Workbooks.Open(filePath, ReadOnly:=True)
    For Each sheet In closedWorkbook.sheets
        If sheet.Name = "survey" Then
            has_survey = True
        End If
        
        If sheet.Name = "choices" Then
            has_choices = True
        End If
        
        Debug.Print sheet.Name
    Next sheet
    
    closedWorkbook.Close False
    
    excelApp.Quit
    
    Set sheet = Nothing
    Set closedWorkbook = Nothing
    Set excelApp = Nothing
    
    If has_choices And has_survey Then
        check_tools_file = True
        Debug.Print "fine"
    Else
        check_tools_file = False
        Debug.Print "not fine"
    End If
    
End Function



