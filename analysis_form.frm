VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} analysis_form 
   Caption         =   "Analysis"
   ClientHeight    =   5286
   ClientLeft      =   36
   ClientTop       =   174
   ClientWidth     =   7452
   OleObjectBlob   =   "analysis_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "analysis_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    DoEvents
    Public_module.CANCEL_PROCESS = True
    End
End Sub

Private Sub CommandRunAnalysis_Click()
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim uuid_col As Long
    
    On Error GoTo errHandler
    
    If Not worksheet_exists("dissagregation_setting") Then
        MsgBox "Please set the dissagregation levels. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If Not worksheet_exists("analysis_list") Then
        MsgBox "Please set the analysis indicators. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If sheets("dissagregation_setting").Cells(2, 1) = vbNullString Then
        MsgBox "Please set the dissagregation levels. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    Dim clean_data As String
    clean_data = find_main_data

    If clean_data = vbNullString Then
        MsgBox "Pleass set your clean data set.      ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    uuid_col = gen_column_number("_uuid", find_main_data)
    
    If uuid_col = 0 Then
        MsgBox "The '_uuid' column dose not exist. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    Call analyze
    
    Call generate_datamerge
    
    wb.Save
    
    Unload analysis_form

    Exit Sub
    
errHandler:
    MsgBox "Pleass set properly your main dataset, disaggregation levels and analysis variables.      ", vbInformation
    Exit Sub

End Sub

Private Sub UserForm_Initialize()

    If Not worksheet_exists("dissagregation_setting") Then
        MsgBox "Please set the dissagregation levels. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If Not worksheet_exists("analysis_list") Then
        MsgBox "Please set the analysis indicators. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If sheets("dissagregation_setting").Cells(2, 1) = vbNullString Then
        MsgBox "Please set the dissagregation levels. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    Public_module.DATA_SHEET = find_main_data
    Me.Frame1.BorderStyle = fmBorderStyleSingle
    Me.TextInfo.SpecialEffect = fmSpecialEffectFlat
    Me.CommandRunAnalysis.BackStyle = fmSpecialEffectFlat
End Sub

