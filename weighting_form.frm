VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} weighting_form 
   Caption         =   "Weighting Setting"
   ClientHeight    =   5034
   ClientLeft      =   -516
   ClientTop       =   -2358
   ClientWidth     =   7476
   OleObjectBlob   =   "weighting_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "weighting_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False













Private Sub CommandTestStrata_Click()
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    If Me.CombData.value = "" Or Me.ComboSampling.value = "" Or Me.ComboDataStrata.value = "" Or Me.ComboSamplingStrata.value = "" Or Me.ComboPopulation.value = "" Then
        MsgBox "Please set up all the parameteres.          ", vbExclamation
        Exit Sub
    End If

    Dim current_smp_name As String
    Dim new_smp_name As String
    
    Dim current_dt_name As String
    Dim new_dt_name As String
      
    current_smp_name = Me.ComboSampling.value
    new_smp_name = alpha_numeric_only(current_smp_name)
    
    current_dt_name = Me.CombData.value
    new_dt_name = alpha_numeric_only(current_dt_name)
    
    If current_dt_name = current_smp_name Then
        MsgBox "Please select a correct sampling or main data sheet. You have selected the same name for both!     ", vbExclamation
        Exit Sub
    End If
    
'    If current_smp_name <> new_smp_name Then
'        MsgBox "Please rename the sampling sheet with no space and no special characters!  ", vbExclamation
'        Exit Sub
'    End If
    
    If Len(new_smp_name) > 15 Then
        new_smp_name = left(new_smp_name, 15)
    End If
    
    If Len(new_dt_name) > 15 Then
        new_dt_name = left(new_dt_name, 15)
    End If
     
    dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", new_dt_name)
     
    If new_dt_name <> current_dt_name Then
        sheets(current_dt_name).Name = new_dt_name
    End If
    
    If new_smp_name <> current_smp_name Then
        sheets(current_smp_name).Name = new_smp_name
    End If

    dt = SaveRegistrySetting("ramSetting", "dataReg", new_dt_name)
    smp = SaveRegistrySetting("ramSetting", "samplingReg", new_smp_name)
    dt_strata = SaveRegistrySetting("ramSetting", "dataStrataReg", Me.ComboDataStrata.value)
    smp_strata = SaveRegistrySetting("ramSetting", "samplingStrataReg", Me.ComboSamplingStrata.value)
    smp_population = SaveRegistrySetting("ramSetting", "samplingPopulationReg", Me.ComboPopulation.value)

    Public_module.DATA_SHEET = new_dt_name
    Public_module.SAMPLE_SHEET = new_smp_name
    Public_module.DATA_STRATA = Me.ComboDataStrata.value
    Public_module.SAMPLE_STRATA = Me.ComboSamplingStrata.value
    Public_module.SAMPLE_POPULATION = Me.ComboPopulation.value
    
    sheets(Public_module.DATA_SHEET).Activate
    clear_active_filter
    
    sheets(Public_module.SAMPLE_SHEET).Activate
    clear_active_filter
    DoEvents
    
    Call generate_strata
    Call unmatched_strata
    Call UserForm_Initialize
    
    Application.DisplayAlerts = False
    
    If worksheet_exists("temp_sheet") Then
        sheets("temp_sheet").Visible = xlSheetHidden
        sheets("temp_sheet").Delete
    End If
     
    sheets(Public_module.SAMPLE_SHEET).Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Sub CommandWeight_Click()
    On Error GoTo ErrorHandler
    If Me.CombData.value = "" Or Me.ComboSampling.value = "" Or Me.ComboDataStrata.value = "" Or _
        Me.ComboSamplingStrata.value = "" Or Me.ComboPopulation.value = "" Then
        MsgBox "Please set up all the parameteres.          ", vbExclamation
        Exit Sub
    End If
    
    Dim current_smp_name As String
    Dim new_smp_name As String
    
    Dim current_dt_name As String
    Dim new_dt_name As String
    
    current_smp_name = Me.ComboSampling.value
    new_smp_name = alpha_numeric_only(current_smp_name)
    
    current_dt_name = Me.CombData.value
    new_dt_name = alpha_numeric_only(current_dt_name)
    
    If current_dt_name = current_smp_name Then
        MsgBox "Please select a correct sampling or main data sheet. You have selected the same name for both!     ", vbExclamation
        Exit Sub
    End If
    
    If current_smp_name <> new_smp_name Then
        MsgBox "Please rename the sampling sheet with no space and no special characters!  ", vbExclamation
        Exit Sub
    End If
    
    If Len(new_dt_name) > 15 Then
        new_dt_name = left(new_dt_name, 15)
    End If
     
    dt_sheet = SaveRegistrySetting("ramSetting", "dataReg", new_dt_name)
     
    If new_dt_name <> current_dt_name Then
        sheets(current_dt_name).Name = new_dt_name
    End If
    
    dt = SaveRegistrySetting("ramSetting", "dataReg", new_dt_name)
    smp = SaveRegistrySetting("ramSetting", "samplingReg", current_smp_name)
    dt_strata = SaveRegistrySetting("ramSetting", "dataStrataReg", Me.ComboDataStrata.value)
    smp_strata = SaveRegistrySetting("ramSetting", "samplingStrataReg", Me.ComboSamplingStrata.value)
    smp_population = SaveRegistrySetting("ramSetting", "samplingPopulationReg", Me.ComboPopulation.value)
    
    Public_module.DATA_SHEET = new_dt_name
    Public_module.SAMPLE_SHEET = current_smp_name
    Public_module.DATA_STRATA = Me.ComboDataStrata.value
    Public_module.SAMPLE_STRATA = Me.ComboSamplingStrata.value
    Public_module.SAMPLE_POPULATION = Me.ComboPopulation.value
    
    Call calculate_weight
    
    Application.DisplayAlerts = False
    
    If worksheet_exists("temp_sheet") Then
        sheets("temp_sheet").Visible = xlSheetHidden
        sheets("temp_sheet").Delete
    End If
    
    Application.DisplayAlerts = True
    
    Me.CommandWeight.Caption = "Prosseccing ..."
    DoEvents
    Unload Me
    Exit Sub
    
ErrorHandler:
    MsgBox "Weighting failed!, please check the sampling framework and your dataset and set the parameters properly.", vbInformation
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next

    With Me
        .StartUpPosition = 0
        .left = Application.left + (0.5 * Application.Width) - (0.5 * .Width)
        .top = Application.top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
    
    Dim sheet_li As Collection
    Set sheet_li = sheet_list
    Dim dt As String
    Dim smp As String
    Dim sh As Variant
    
    For Each sh In sheet_li
        If ActiveWorkbook.Worksheets(CStr(sh)).Visible And Not IsInArray(CStr(sh), InitializeExcludedSheets) And left(CStr(sh), 6) <> "chart-" Then
            Me.ComboSampling.AddItem sh
        End If
    Next sh
    
    dt = GetRegistrySetting("ramSetting", "dataReg")
    smp = GetRegistrySetting("ramSetting", "samplingReg")
    
    If worksheet_exists(smp) Then
        Me.ComboSampling.value = smp
    End If
    
    If worksheet_exists(dt) Then
        Me.CombData.value = dt
    End If
  
    Me.ComboDataStrata.value = GetRegistrySetting("ramSetting", "dataStrataReg")
    Me.ComboSamplingStrata.value = GetRegistrySetting("ramSetting", "samplingStrataReg")
    Me.ComboPopulation.value = GetRegistrySetting("ramSetting", "samplingPopulationReg")
       
End Sub

Private Sub PopulateComboBox(SHEET_NAME As String, con As String)
    On Error Resume Next
    Dim header_arr() As Variant
    Dim c As control
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.sheets(SHEET_NAME)
    
    ' array of header (list of questions)
    header_arr = ws.Range(ws.Cells(1, 1), ws.Cells(1, 1).End(xlToRight)).Value2
    
    For Each c In Me.Controls
        If c.Name = con Then
            c.Clear
            For Each i In header_arr

                c.AddItem i
            Next
        End If
    Next
    
End Sub

Private Sub ComboSampling_Change()
    'This subroutine updates the population and sampling strata combo boxes based on the selected worksheet name
    Dim val As String
    val = Me.ComboSampling.value
    
    Me.ComboPopulation.Enabled = True
    Me.ComboPopulation.Clear
    
    Me.ComboSamplingStrata.Enabled = True
    Me.ComboSamplingStrata.Clear
    
    Call PopulateComboBox(val, "ComboPopulation")
    Call PopulateComboBox(val, "ComboSamplingStrata")
End Sub

Private Sub CombData_Change()
    Dim val As String
    val = Me.CombData.value

    Me.ComboDataStrata.Enabled = True
    Me.ComboDataStrata.Clear
    
    Call PopulateComboBox(val, "ComboDataStrata")
End Sub


