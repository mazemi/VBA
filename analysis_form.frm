VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} analysis_form 
   Caption         =   "Analysis"
   ClientHeight    =   5172
   ClientLeft      =   -252
   ClientTop       =   -1122
   ClientWidth     =   7020
   OleObjectBlob   =   "analysis_form.frx":0000
   ShowModal       =   0   'False
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
Application.ScreenUpdating = False
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim uuid_col As Long
    Dim start_time As Double
    start_time = Timer
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
    
    DoEvents
    
    Me.dmLabel.Visible = False
    Debug.Print "call dm: " & Timer - start_time

    Call do_analize

    str_info = vbLf & analysis_form.TextInfo.Value
    txt = "Generating Datamerge... " & str_info
    Me.TextInfo.Value = txt
    Me.Repaint
    Call generate_datamerge

    Application.DisplayAlerts = False
    
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
        
    Application.DisplayAlerts = True
    
    wb.Save
    
    Unload analysis_form

    Exit Sub

errHandler:

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
    
    Application.DisplayAlerts = True
    
    MsgBox " Oops!, Something went wrong! Pleass check properly your main dataset, disaggregation levels and analysis variables.      ", vbInformation
    
    Unload analysis_form
    
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
    
    If worksheet_exists("datamerge") Then
        Me.dmLabel.Visible = True
    End If
    
    Public_module.DATA_SHEET = find_main_data
    Me.Frame1.BorderStyle = fmBorderStyleSingle
    Me.TextInfo.SpecialEffect = fmSpecialEffectFlat
    Me.CommandRunAnalysis.BackStyle = fmSpecialEffectFlat
End Sub

