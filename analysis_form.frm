VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} analysis_form 
   Caption         =   "Analysis"
   ClientHeight    =   5286
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   7542
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
    cancel_proc = True
    
'    MsgBox "Process canceled.               ", vbInformation
    End
End Sub

Private Sub CommandRunAnalysis_Click()
t = Timer

'    On Error GoTo errHandler
    Dim dt As String
    dt = sheets("dissagregation_setting").Cells(1, 4)
    
    If dt = "" Then
        MsgBox "Pleass set your disaggregation levels.      "
    End If
    
    Call analyze
       
 MsgBox Timer - t
 
'    Exit Sub

'errHandler:
'   MsgBox "Pleass set your disaggregation levels and analysis variables.      "
'   Exit Sub
End Sub

Private Sub UserForm_Initialize()
    Me.Frame1.BorderStyle = fmBorderStyleSingle
    Me.TextInfo.SpecialEffect = fmSpecialEffectFlat
    Me.CommandRunAnalysis.BackStyle = fmSpecialEffectFlat
End Sub

Sub analyze()
cancel_proc = False
Application.ScreenUpdating = False
'Application.EnableEvents = False
Dim ws As Worksheet
Dim keenSheet As Worksheet
Dim main_ws As Worksheet
Dim unique_diss_options As Range, Item As Variant
Dim unique_diss_options_labels As Range, rng As Range
Dim SourceRange As Range, cl As Range, small_dt As Range
Dim disaggregation_labels_collection As New Collection
Dim disaggregation_collection As New Collection
Dim new_row As Long
Dim last_row_survey_sheet As Long
Dim k As Long
Dim filtered_col As Long
Dim measurement As String, diss As String, combined_cols As String
Dim weighting As Boolean, numeric As Boolean, over_all As Boolean, calcuted_index As Boolean
Dim i As Variant
Dim diss_rng As Range, question_rng As Range
Dim str_info As String

Dim res As Variant

Dim result_df As New Array2D
result_df.data = Null
result_df.insertColumnsBlank 1, 10

With sheets("dissagregation_setting")
    last_diss = .Cells(rows.count, 1).End(xlUp).row
    Set diss_rng = .Range("A2:B" & last_diss)
End With

With sheets("analysis_setting")
    last_question = .Cells(rows.count, 1).End(xlUp).row
    Set question_rng = .Range("A2:B" & last_question)
End With

Set main_ws = sheets("new_RAM2_clean_data")

Dim diss_value As Range

' check if result sheet exist
Call check_result_sheet(main_ws.Name)

' check if keen sheet exist
If WorksheetExists("keen") <> True Then
    Call create_sheet(main_ws.Name, "keen")
End If

Set keenSheet = sheets("keen")
'sheets("keen").Visible = False
Set ws = sheets("result")

For Each diss_value In diss_rng.columns(1).Cells
    DoEvents
    is_weight = diss_rng.columns(2).rows(diss_value.row - 1)
    diss = diss_value
    If is_weight = "yes" Then
        weighting = True
    Else
        weighting = False
    End If
    
    ' start looping through question_rng
    For Each q_value In question_rng.columns(1).Cells
    
        If cancel_proc Then
            End
        End If

        If Len(str_info) > 2000 Then
            Me.TextInfo.value = Left(Me.TextInfo.value, 1000)
        End If
       
        str_info = vbLf & Me.TextInfo.value
        
        txt = "Disaggregation level : " & diss_value & " > " & q_value & CStr(ana) & str_info
        txt = Replace(txt, "0", "")
        Me.TextInfo.value = txt
        
        ' main doevents
        DoEvents
        
'        Application.StatusBar = "Disaggregation level : " & diss_value & " > " & q_value
        
        is_number = question_rng.columns(2).rows(q_value.row - 1)
        If is_number = "number" Then
            numeric = True
        Else
            numeric = False
        End If
        
        measurement = q_value
        calcuted_index = False
        
        ' start of the numeric process
        diss_col_letter = gen_column_letter(diss, main_ws.Name)
        measurement_col_letter = gen_column_letter(measurement, main_ws.Name)
        
        weight_col_letter = gen_column_letter("weight", main_ws.Name)
        
        last_row_choice = Worksheets("choices").Cells(rows.count, 1).End(xlUp).row
        last_row_survey_sheet = Worksheets("survey").Cells(rows.count, 1).End(xlUp).row
        
        Dim sum_w2 As Single
        Dim sum_weight As Single
        Dim mean_w_value As Single
        Dim simple_mean As Single
        
        Dim count_weight As Long
        Dim simple_count As Long
    
        If diss = "ALL" And numeric = True Then
            new_row = ws.UsedRange.rows(ws.UsedRange.rows.count).row + 1
            If weighting Then
            
                ' copy necesory data to sheet keen sheet
                keenSheet.columns("A") = main_ws.columns(measurement_col_letter).Value2
                keenSheet.columns("B") = main_ws.columns(weight_col_letter).Value2
                
                last_row_keen = keenSheet.Cells(rows.count, 1).End(xlUp).row
                
                Set small_dt = keenSheet.Cells(1, 1).CurrentRegion
                
                Call delete_blank_rows(1)
'                Call remove_rows(small_dt, 1)
                
                Call add_mulitipication("C")
                
                sum_w2 = Application.WorksheetFunction.Sum(keenSheet.Range("C2:C" & CStr(last_row_keen)))
                sum_weight = Application.WorksheetFunction.Sum(keenSheet.Range("B2:B" & CStr(last_row_keen)))
                count_weight = Application.WorksheetFunction.count(keenSheet.Range("B2:B" & CStr(last_row_keen)))
                If count_weight > 0 Then
                    
                    result_df.insertRowsBlank result_df.rowCount + 1
                    n = result_df.rowCount
                    
                    result_df.value(n, 1) = n - 1
                    result_df.value(n, 2) = diss
                    result_df.value(n, 3) = diss
                    result_df.value(n, 4) = diss
                    result_df.value(n, 5) = Worksheets("keen").Cells(1, 1)
                    result_df.value(n, 6) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), _
                        WorksheetFunction.Match(keenSheet.Cells(1, 1), sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))

                    result_df.value(n, 7) = "mean"
                    result_df.value(n, 8) = Application.WorksheetFunction.Round(sum_w2 / sum_weight, 2)
                    result_df.value(n, 9) = count_weight
                    
                Else
                        
                    result_df.insertRowsBlank result_df.rowCount + 1
                    
                    n = result_df.rowCount
                    
                    result_df.value(n, 1) = n - 1
                    result_df.value(n, 2) = diss
                    result_df.value(n, 3) = diss
                    result_df.value(n, 4) = diss
                    result_df.value(n, 5) = Worksheets("keen").Cells(1, 1)
                    result_df.value(n, 6) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), _
                        WorksheetFunction.Match(keenSheet.Cells(1, 1), sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))

                    result_df.value(n, 7) = "mean"
                    result_df.value(n, 9) = 0
                        
                End If
            Else ' all disaggregation without weighting
            
                ' copy necesory data to sheet keen sheet

                keenSheet.columns("A") = main_ws.columns(measurement_col_letter).Value2
                
                last_row_keen = keenSheet.Cells(rows.count, 1).End(xlUp).row
                
                Set small_dt = keenSheet.Cells(1, 1).CurrentRegion
                
                Call delete_blank_rows(1)
'                Call remove_rows(small_dt, 1)
            
                simple_count = Application.WorksheetFunction.count(keenSheet.Range("A2:A" & CStr(last_row_keen)))
                If simple_count > 0 Then
                    simple_mean = Application.WorksheetFunction.Average(keenSheet.Range("A2:A" & CStr(last_row_keen)))
                    
                    result_df.insertRowsBlank result_df.rowCount + 1
                    
                    n = result_df.rowCount

                    result_df.value(n, 1) = n - 1
                    result_df.value(n, 2) = diss
                    result_df.value(n, 3) = diss
                    result_df.value(n, 4) = diss
                    result_df.value(n, 5) = Worksheets("keen").Cells(1, 1)
                    result_df.value(n, 6) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), _
                         WorksheetFunction.Match(keenSheet.Cells(1, 1), sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))
                    result_df.value(n, 7) = "mean"
                    result_df.value(n, 8) = Application.WorksheetFunction.Round(simple_mean, 2)
                    result_df.value(n, 9) = simple_count
                    
                Else
                    result_df.insertRowsBlank result_df.rowCount + 1
                    
                    n = result_df.rowCount

                    result_df.value(n, 1) = n - 1
                    result_df.value(n, 2) = diss
                    result_df.value(n, 3) = diss
                    result_df.value(n, 4) = diss
                    result_df.value(n, 5) = Worksheets("keen").Cells(1, 1)
                    result_df.value(n, 6) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), _
                         WorksheetFunction.Match(keenSheet.Cells(1, 1), sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))
                    result_df.value(n, 7) = "mean"

                    result_df.value(n, 9) = 0
                    
                End If
            
            End If
        ElseIf diss <> "ALL" And numeric = True Then ' disaggregation is not all and the variable is numeric
        
            '* check if diss var is a select on or not?
            diss_type = match_type(diss)
            
            keenSheet.columns("A") = main_ws.columns(diss_col_letter).Value2
            
            keenSheet.columns("B") = main_ws.columns(measurement_col_letter).Value2
            
            If weighting Then
                keenSheet.columns("C") = main_ws.columns(weight_col_letter).Value2
            End If
            
            If Left(diss_type, 10) = "select_one" Then
            
                '* find diss labels and all choices for select one only:
                key_name_arr = Split(diss_type, " ")
                key_name = key_name_arr(UBound(key_name_arr))
            
                Set SourceRange = Worksheets("choices").Range("A1:C" & last_row_choice)
                SourceRange.AutoFilter
                SourceRange.AutoFilter Field:=1, Criteria1:=key_name
                SourceRange.SpecialCells (xlCellTypeVisible)
                
                Set unique_diss_options = SourceRange.columns(SourceRange.columns.count - 1).Cells
                Set disaggregation_collection = unique_index_values(unique_diss_options.SpecialCells(xlCellTypeVisible))
                
                Set unique_diss_options_labels = SourceRange.columns(SourceRange.columns.count).SpecialCells(xlCellTypeVisible).Cells
                
                ' feeding the collection of disaggrigation options
                For Each i In unique_diss_options_labels
                    disaggregation_labels_collection.Add i
                Next
        
            Else ' the disaggregation level is calculated field
                calcuted_index = True
                Dim calculated_col As Range
                last_row_main_data = main_ws.Cells(rows.count, 1).End(xlUp).row
                Set calculated_col = main_ws.Range(diss_col_letter & "2:" & diss_col_letter & last_row_main_data)
            
                Set disaggregation_collection = unique_index_values(calculated_col)
                Set disaggregation_labels_collection = disaggregation_collection
                
                If disaggregation_collection.count > 10 Then
                    Debug.Print "There is a problem with the disaggregation variable!"
                End If
                
                  
            End If
            
            ' new row number in the analysis sheet
            new_row = ws.UsedRange.rows(ws.UsedRange.rows.count).row + 1
            
            last_row_keen = keenSheet.Cells(rows.count, 1).End(xlUp).row
            
            Set small_dt = keenSheet.Cells(1, 1).CurrentRegion
            
            Call delete_blank_rows(2)
'            Call remove_rows(small_dt, 2)
            
            last_row_keen = keenSheet.Cells(rows.count, 1).End(xlUp).row
            
            ' k is counter for looping through the disaggregation_labels_collection
            k = 0
            
            '* if weighting is required
            If weighting Then
                
                Call add_mulitipication("D")
                
                For Each Item In disaggregation_collection
'                DoEvents
                    k = k + 1
                    If Item <> "name" Then
                        sum_w2 = Application.WorksheetFunction.SumIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), _
                            Item, keenSheet.Range("D2:D" & CStr(last_row_keen)))
                        sum_weight = Application.WorksheetFunction.SumIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), _
                            Item, keenSheet.Range("C2:C" & CStr(last_row_keen)))
                        count_weight = Application.WorksheetFunction.CountIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), Item)
                        If sum_w2 > 0 Then
                        
                            result_df.insertRowsBlank result_df.rowCount + 1
                            
                            n = result_df.rowCount
        
                            result_df.value(n, 1) = n - 1
                            result_df.value(n, 2) = diss
                            result_df.value(n, 3) = Item
                            result_df.value(n, 4) = disaggregation_labels_collection(k)
                            result_df.value(n, 5) = Worksheets("keen").Cells(1, 2)
                            result_df.value(n, 6) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), _
                                 WorksheetFunction.Match(keenSheet.Cells(1, 2), sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))
                            result_df.value(n, 7) = "mean"
                            result_df.value(n, 8) = Application.WorksheetFunction.Round(sum_w2 / sum_weight, 2)
                            result_df.value(n, 9) = count_weight

                            
                        Else
                            result_df.insertRowsBlank result_df.rowCount + 1
                            
                            n = result_df.rowCount
        
                            result_df.value(n, 1) = n - 1
                            result_df.value(n, 2) = diss
                            result_df.value(n, 3) = Item
                            result_df.value(n, 4) = disaggregation_labels_collection(k)
                            result_df.value(n, 5) = Worksheets("keen").Cells(1, 2)
                            
                            If calcuted_index Then
                                result_df.value(n, 6) = ""
                            Else
                                result_df.value(n, 6) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), _
                                    WorksheetFunction.Match(keenSheet.Cells(1, 2), sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))
                            End If
                            result_df.value(n, 7) = "mean"

                            result_df.value(n, 9) = 0

                        End If
                        
                        new_row = new_row + 1
                    End If
                Next
            ElseIf Not weighting Then
                For Each Item In disaggregation_collection
'                    DoEvents
                    k = k + 1
                    If Item <> "name" Then
                        simple_count = WorksheetFunction.CountIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), Item)
                        If simple_count > 0 Then
                            simple_mean = WorksheetFunction.AverageIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), _
                                Item, keenSheet.Range("B2:B" & CStr(last_row_keen)))
                                
                            result_df.insertRowsBlank result_df.rowCount + 1
                            
                            n = result_df.rowCount
                            
                            result_df.value(n, 1) = new_row - 1
                            result_df.value(n, 2) = diss
                            result_df.value(n, 3) = Item
                            result_df.value(n, 4) = disaggregation_labels_collection(k)
                            result_df.value(n, 5) = Worksheets("keen").Cells(1, 2)
                            result_df.value(n, 6) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), _
                                WorksheetFunction.Match(keenSheet.Cells(1, 2), sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))

                            result_df.value(n, 7) = "mean"
                            result_df.value(n, 8) = Application.WorksheetFunction.Round(simple_mean, 2)
                            result_df.value(n, 9) = simple_count
                        Else
                            result_df.insertRowsBlank result_df.rowCount + 1
                            
                            n = result_df.rowCount

                            result_df.value(n, 1) = new_row - 1
                            result_df.value(n, 2) = diss
                            result_df.value(n, 3) = Item
                            result_df.value(n, 4) = disaggregation_labels_collection(k)
                            result_df.value(n, 5) = Worksheets("keen").Cells(1, 2)
                            If calcuted_index Then
                                result_df.value(n, 6) = ""
                            Else
                                result_df.value(n, 6) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), _
                                    WorksheetFunction.Match(keenSheet.Cells(1, 2), sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))
                            End If
                            result_df.value(n, 7) = "mean"
                            result_df.value(n, 9) = 0
                        End If
                        
                        new_row = new_row + 1
                    End If
                Next
            End If
        
        
        ElseIf diss = "ALL" And Not numeric Then  ' disaggregation is all and the variable is catagory
        
            ' if select one
            
        
            
            ' if select multiple
        
        
        ElseIf diss <> "ALL" And Not numeric Then  ' disaggregation is all and the variable is catagory
        
            ' if select one
        
            
            ' if select multiple
        
            Set disaggregation_collection = Nothing
            Set disaggregation_labels_collection = Nothing
        
        End If
        
        sheets("keen").Cells.Clear
    

    Next ' loop for question_rng
    
    ' write the result of disaggregation to the result sheet
    last_result = sheets("result").Range("A" & rows.count).End(xlUp).row
    result_df.writeDataToRange sheets("result").Range("A" & last_result + 1)
    result_df.data = Null
    result_df.insertColumnsBlank 1, 10
    
Next ' loop for disaggrigation

'Application.StatusBar = False
Application.ScreenUpdating = True
End Sub

' remove null rows from keen sheet
Private Sub remove_rows(rng As Range, col As Long)
    Dim Lrange As Range
    Dim n As Long
    Set Lrange = rng
    For n = Lrange.rows.count To 1 Step -1
        If Lrange.Cells(n, col).value = "" Then
            Lrange.rows(n).Delete
        End If
    Next
End Sub

' if we need to apply weighting a column will be generated by the name w2 = numeric_value * weight
Sub add_mulitipication(target_col As String)
    Dim last_row As Long
    last_row = Worksheets("keen").Cells(rows.count, 1).End(xlUp).row
    Worksheets("keen").Range(target_col & "1").FormulaR1C1 = "w2"
    Application.CutCopyMode = False
    Worksheets("keen").Range(target_col & "2").FormulaR1C1 = "=RC[-2]*RC[-1]"
    Worksheets("keen").Range(target_col & "2").AutoFill Destination:=Worksheets("keen").Range(target_col & "2:" & target_col & CStr(last_row))
End Sub

Sub check_result_sheet(sheet_name As String)
    ' check if keen sheet exist
    If WorksheetExists("result") <> True Then
        Call create_sheet(sheet_name, "result")
        sheets("result").Cells(1, 1) = "row"
        sheets("result").Cells(1, 2) = "disaggregation"
        sheets("result").Cells(1, 3) = "disaggregation value"
        sheets("result").Cells(1, 4) = "disaggregation label"
        sheets("result").Cells(1, 5) = "variable"
        sheets("result").Cells(1, 6) = "variable label"
        sheets("result").Cells(1, 7) = "measurement type"
        sheets("result").Cells(1, 8) = "measurement value"
        sheets("result").Cells(1, 9) = "measurement numbers"
    End If
End Sub

Function unique_index_values(rng As Range) As Collection
    Dim d As Object, c As Range, h, tmp As String
    Dim unique_collection As New Collection
    
    Set d = CreateObject("scripting.dictionary")
    For Each c In rng
        tmp = Trim(c.value)
        If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
    Next c

    For Each h In d.Keys
'        Debug.Print h
         unique_collection.Add CStr(h)
    Next h
    Set unique_index_values = unique_collection
End Function

' remove null rows from keen sheet
Sub delete_blank(col As Long)
    On Error GoTo NoBlanks
    
'    sheets("keen").Range("C1:C8210").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    sheets("keen").colunms(col).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Exit Sub
NoBlanks:
   
    Exit Sub
End Sub


Sub delete_blank_rows(col_number As Long)

    Dim rng As Range
    Dim str_delete As String
    
    Set rng = sheets("keen").Range("A1").CurrentRegion
    
    last_keen = rng.rows.count
    
    rng.Sort rng.columns(col_number), , , , , , , Header:=xlYes
    
    last_measurement = sheets("keen").Cells(rows.count, col_number).End(xlUp).row
    
    str_delete = CStr(last_measurement + 1) & ":" & last_keen
    
    If last_measurement < last_keen Then
        sheets("keen").rows(str_delete).Delete Shift:=xlUp
'        MsgBox str_delete
    End If

End Sub
