VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} analysis_form 
   Caption         =   "Analysis"
   ClientHeight    =   4908
   ClientLeft      =   -576
   ClientTop       =   -2682
   ClientWidth     =   7902
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
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim t As Double
    t = Timer
    Set wb = ActiveWorkbook
    Dim uuid_col As Long
    Dim start_time As Double
    start_time = Timer
    On Error GoTo ErrorHandler
    
    If Not worksheet_exists("disaggregation_setting") Then
        MsgBox "Please set the disaggregation levels. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If Not worksheet_exists("analysis_list") Then
        MsgBox "Please set the analysis indicators. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If sheets("disaggregation_setting").Cells(2, 1) = vbNullString Then
        MsgBox "Please set the disaggregation levels. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If find_uuid_coln = 0 Then
        MsgBox "The '_uuid' column dose not exist in your data. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If check_exist_dis_levels() <> vbNullString Then
        MsgBox check_exist_dis_levels & " disagregation dose not exist in the clean dataset. Please set the disaggregation levels properly. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If check_null_dis_levels() <> vbNullString Then
        MsgBox check_null_dis_levels & " disagregation has empty valuse. Please set the disaggregation levels properly. ", vbInformation
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
    
    Call save_as_xlsx
    
    uuid_col = gen_column_number("_uuid", find_main_data)
    
    If uuid_col = 0 Then
        MsgBox "The '_uuid' column dose not exist. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    DoEvents
    
    Me.dmLabel.Visible = False
    Debug.Print "start analysis: ", Timer - t
    Call do_analize

    str_info = vbLf & analysis_form.TextInfo.value
    txt = "Generating Datamerge... " & str_info
    Me.TextInfo.value = txt
    Me.Repaint
    
    Debug.Print "start datamerge: ", Timer - t
    Call generate_datamerge
    Debug.Print "end datamerge: ", Timer - t
    Call remove_tmp
    
    wb.Save
    Debug.Print "end of process: ", Timer - t
    Unload analysis_form
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Exit Sub

ErrorHandler:

    Call remove_tmp
    MsgBox " Oops!, Something went wrong! Pleass check properly your main dataset, disaggregation levels and analysis variables.  ", vbInformation
    Unload analysis_form

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Exit Sub

End Sub

Private Sub UserForm_Initialize()
    
    With Me
        .StartUpPosition = 0
        .left = Application.left + (0.5 * Application.Width) - (0.5 * .Width)
        .top = Application.top + (0.5 * Application.Height) - (0.5 * .Height)
    End With
      
    If Not worksheet_exists("disaggregation_setting") Then
        MsgBox "Please set the disaggregation levels. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If Not worksheet_exists("analysis_list") Then
        MsgBox "Please set the analysis indicators. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If sheets("disaggregation_setting").Cells(2, 1) = vbNullString Then
        MsgBox "Please set the disaggregation levels. ", vbInformation
        Unload analysis_form
        Exit Sub
    End If
    
    If worksheet_exists("datamerge") Then
        Me.dmLabel.Visible = True
    End If
'    Me.CheckBoxNonSeletedOptions.value = True
    
    Public_module.DATA_SHEET = find_main_data
    Me.Frame1.BorderStyle = fmBorderStyleSingle
    Me.TextInfo.SpecialEffect = fmSpecialEffectFlat
    Me.CommandRunAnalysis.BackStyle = fmSpecialEffectFlat
End Sub

Sub save_as_xlsx()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim currentFilePath As String
    Dim newFilePath As String
    
    Set wb = ActiveWorkbook
    
    If LCase(Right(wb.FullName, 3)) = "csv" Then
        currentFilePath = wb.FullName
        newFilePath = left(currentFilePath, Len(currentFilePath) - 3) & "xlsx"
        wb.SaveAs FileName:=newFilePath, FileFormat:=xlOpenXMLWorkbook
'        wb.Close False
        Debug.Print "File saved as XLSX format"
    End If
End Sub

