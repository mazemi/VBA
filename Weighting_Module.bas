Attribute VB_Name = "Weighting_Module"
Option Explicit

Sub generate_strata()
    On Error Resume Next
    Public_module.DATA_SHEET = find_main_data
    Dim main_strata_col_number As Long
    Dim samp_strata_col_number As Long
    Dim main_ws As Worksheet
    Dim last_main_strata As Long
    Dim last_smp_strata As Long
    
    Set main_ws = sheets(Public_module.DATA_SHEET)
    
    If Not worksheet_exists("temp_sheet") Then
        Call create_sheet(main_ws.Name, "temp_sheet")
    End If
    
    sheets("temp_sheet").Cells.Clear

    main_strata_col_number = gen_column_number(Public_module.DATA_STRATA, Public_module.DATA_SHEET)
    sheets(Public_module.DATA_SHEET).Columns(main_strata_col_number).Copy Destination:=sheets("temp_sheet").Columns(1)
    sheets("temp_sheet").Columns(1).RemoveDuplicates Columns:=1, Header:=xlNo

    samp_strata_col_number = gen_column_number(Public_module.SAMPLE_STRATA, Public_module.SAMPLE_SHEET)
    sheets(Public_module.SAMPLE_SHEET).Columns(samp_strata_col_number).Copy Destination:=sheets("temp_sheet").Columns(2)
    sheets("temp_sheet").Columns(2).RemoveDuplicates Columns:=1, Header:=xlNo
    
    last_main_strata = sheets("temp_sheet").Cells(Rows.count, 1).End(xlUp).Row
    last_smp_strata = sheets("temp_sheet").Cells(Rows.count, 2).End(xlUp).Row
    
    sheets("temp_sheet").Range("C2:C" & last_main_strata).Formula = "=A2 & ""A"""
    sheets("temp_sheet").Range("D2:D" & last_smp_strata).Formula = "=B2 & ""A"""
    
    sheets("temp_sheet").Columns("C:D").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Application.CutCopyMode = False
    sheets("temp_sheet").Columns("A:B").Delete Shift:=xlToLeft
    
End Sub

Sub unmatched_strata()
    On Error Resume Next
    Dim m_strata() As Variant
    Dim S_strata() As Variant
    Dim col As New Collection
    Dim msg, msg_title As String
    Dim last_main_strata As Long
    Dim last_smp_strata As Long
    Dim main_strata As Variant
    Dim smp_strata As Variant
    Dim i As Variant
    
    msg_title = "The following strata dose not exist in the sampling frame." & vbCrLf & _
                "Please check the data and sampling framework for below codes:" & vbCrLf
    
    last_main_strata = sheets("temp_sheet").Cells(Rows.count, 1).End(xlUp).Row
    last_smp_strata = sheets("temp_sheet").Cells(Rows.count, 2).End(xlUp).Row

    main_strata = sheets("temp_sheet").Range("A2:A" & last_main_strata).Value2
    smp_strata = sheets("temp_sheet").Range("B2:B" & last_smp_strata).Value2
    
    Set col = unmatched_elements(main_strata, smp_strata, False)
    
    For Each i In col
         msg = left(i, Len(i) - 1) & ", " & msg
    Next
    
    If msg <> "" Then
        msg = left(msg, Len(msg) - 2)
        msg = msg_title & msg
        MsgBox msg, vbExclamation
    Else
        MsgBox "Samapling information is good.         ", vbInformation
    End If
    
    Application.DisplayAlerts = False
            
    If worksheet_exists("temp_sheet") Then
        sheets("temp_sheet").Delete
    End If
    
    Application.DisplayAlerts = True
   
End Sub

Sub calculate_weight()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim samp_ws As Worksheet
    Dim dt_ws As Worksheet
    Dim main_stra As String
    Dim sample_stra As String
    Dim sample_stra_number As Long
    Dim main_stra_number As Long
    Dim sample_population_number As Long
    Dim total_population As Long
    Dim total_survey As Long
    Dim w_col_number As Long
    Dim w_col As String
    Dim sw_col As String
    Dim fw_col As String
    Dim diff_col3 As Long
    Dim total_weight0 As Double
    Dim alpha As Double
    Dim new_col_number As Long
    Dim new_col As String
    Dim last_row As Long
    Dim diff_col1, diff_col2 As Long
    Dim smp_weight_col_number As Long
    Dim smp_weight_col As String
    Dim weight_col_number As Long
    Dim weight_col As String
    Dim diff_main As Long
    Dim diff_sampling_strata As Long
    Dim diff_sampling_weight As Long
    Dim diff_sampling_strata_weight As Long
    Dim data_last_row As Long
    Dim population_col As String
    Dim last_row_dt As Long
    Dim new_col_number_dt As Long
    Dim new_col_dt As String
    Dim last_col As String
    Dim new_strata_col_number As Long
    Dim new_strata_col As String
    Dim befor_weight_col As String
    
    Set dt_ws = sheets(Public_module.DATA_SHEET)
    Set samp_ws = sheets(Public_module.SAMPLE_SHEET)
    
    Call remove_empty_col
    
    Call clear_filter(dt_ws)
    Call clear_filter(samp_ws)
    
    ' column letters and coulmn numbers
    main_stra = gen_column_letter(Public_module.DATA_STRATA, Public_module.DATA_SHEET)
    main_stra_number = gen_column_number(Public_module.DATA_STRATA, Public_module.DATA_SHEET)
    sample_stra = gen_column_letter(Public_module.SAMPLE_STRATA, Public_module.SAMPLE_SHEET)
    sample_stra_number = gen_column_number(Public_module.SAMPLE_STRATA, Public_module.SAMPLE_SHEET)
    population_col = gen_column_letter(Public_module.SAMPLE_POPULATION, Public_module.SAMPLE_SHEET)
    sample_population_number = Public_module.letter_to_number(population_col, samp_ws)
       
    ' last column in sampling sheet
    new_col_number = samp_ws.Cells(1, Columns.count).End(xlToLeft).Column + 1
    
    ' last column in main sheet
    new_col_number_dt = dt_ws.Cells(1, Columns.count).End(xlToLeft).Column + 1
    new_col_dt = Split(dt_ws.Cells(, new_col_number_dt).Address, "$")(1)
    
    ' last rows
    last_row = samp_ws.Cells(Rows.count, 1).End(xlUp).Row
'    last_row = samp_ws.UsedRange.rows(samp_ws.UsedRange.rows.count).row
    last_row_dt = dt_ws.UsedRange.Rows(dt_ws.UsedRange.Rows.count).Row
    
    'new codes:
    ' ********** in the sampling sheet
    Dim row_number As Long
    
    ' new strata in the sampling frame
    For row_number = 1 To last_row
        samp_ws.Cells(row_number, new_col_number) = samp_ws.Cells(row_number, sample_stra_number) & "A"
    Next row_number
    
    ' new strata in the main dataset
    For row_number = 1 To last_row_dt
        dt_ws.Cells(row_number, new_col_number_dt) = dt_ws.Cells(row_number, main_stra_number) & "A"
    Next row_number
    
    ' add number of surveyed
    samp_ws.Cells(1, new_col_number + 1) = "surveyed"
    For row_number = 2 To last_row
        samp_ws.Cells(row_number, new_col_number + 1) = Application.WorksheetFunction.CountIf(dt_ws.Columns(new_col_number_dt), _
            samp_ws.Cells(row_number, new_col_number))
    Next row_number
    
    total_population = WorksheetFunction.sum(samp_ws.Columns(sample_population_number))
    total_survey = WorksheetFunction.sum(samp_ws.Columns(new_col_number + 1))
     
     ' add weight0 and sum_weight0
    Dim w0 As Double
    samp_ws.Cells(1, new_col_number + 2) = "weight0"
    samp_ws.Cells(1, new_col_number + 3) = "sum_weight0"
    For row_number = 2 To last_row
    
        If samp_ws.Cells(row_number, new_col_number + 1) > 0 Then
            w0 = (samp_ws.Cells(row_number, sample_population_number) / total_population) / ((samp_ws.Cells(row_number, new_col_number + 1) / total_survey))
        Else
             w0 = 0
        End If
        
        samp_ws.Cells(row_number, new_col_number + 2) = w0
        samp_ws.Cells(row_number, new_col_number + 3) = w0 * samp_ws.Cells(row_number, new_col_number + 1)
         
    Next row_number
    
    total_weight0 = WorksheetFunction.sum(samp_ws.Columns(new_col_number + 3))
    
    Dim correction_coefficient As Double
    correction_coefficient = total_survey / total_weight0
    
    ' add final weight
    samp_ws.Cells(1, new_col_number + 4) = "weight"
    For row_number = 2 To last_row
        samp_ws.Cells(row_number, new_col_number + 4) = correction_coefficient * samp_ws.Cells(row_number, new_col_number + 2)
    Next row_number
    
    ' ********** in the main dataset
    
    ' target column letteres in the sampling sheet
    new_col = Split(samp_ws.Cells(, new_col_number).Address, "$")(1)
    last_col = Split(samp_ws.Cells(, new_col_number + 4).Address, "$")(1)
    befor_weight_col = Split(samp_ws.Cells(, new_col_number + 3).Address, "$")(1)
    
    ' last column in the data sheet
    weight_col_number = dt_ws.Cells(1, Columns.count).End(xlToLeft).Column + 1
    weight_col = Split(dt_ws.Cells(, weight_col_number).Address, "$")(1)
    
    dt_ws.Cells(1, weight_col_number).value = "weight"
    Dim target_rng As Range
    
    ' columns("E:H")
    Set target_rng = samp_ws.Columns(new_col & ":" & last_col)
    Dim i As Variant
    
    For row_number = 2 To last_row_dt
        dt_ws.Cells(row_number, new_col_number_dt + 1) = Application.VLookup(dt_ws.Cells(row_number, new_col_number_dt), _
            target_rng, 5, False)
    Next
    
    ' delete temperory columns
    samp_ws.Columns(new_col & ":" & befor_weight_col).Delete Shift:=xlToLeft
    dt_ws.Columns(new_col_dt & ":" & new_col_dt).Delete Shift:=xlToLeft

    MsgBox "The weight has been added.           ", vbInformation

    Application.ScreenUpdating = True
    Application.CutCopyMode = False

End Sub


