Attribute VB_Name = "analysis_module"
Option Explicit

Sub analyze()
    
    Public_module.CANCEL_PROCESS = False
    Application.ScreenUpdating = False
    Dim result_sheet As Worksheet
    Dim keen_sheet As Worksheet
    Dim main_ws As Worksheet
    Dim disaggregation_collection As New Collection
    Dim unique_choices As New Collection
    Dim last_row_survey_sheet As Long
    Dim last_row_choice As Long
    Dim last_row_survey_choices As Long
    Dim measurement As Variant
    Dim measurement_str As String, dis_str As String
    Dim i As Variant ' for looping through select choices
    Dim dis_rng As Range, question_rng As Range
    Dim str_info As String
    Dim overall As Boolean, wgt As Boolean, m_type As String
    Dim dis_value As Range
    Dim sc_sheet As Worksheet
    Dim sum_w2 As Double
    Dim sum_weight As Double
    Dim mean_w_value As Double
    Dim simple_mean As Double
    Dim count_weight As Long
    Dim simple_count As Long
    Dim last_dis As Long
    Dim last_question As Long
    Dim last_row_main_data As Long
    Dim measurement_rng As Range
    Dim weight_col_letter As String
    Dim is_weight As String
    Dim dis_col_letter As String
    Dim txt As String
    Dim measurement_col_letter As String
    Dim measurement_type As String
    Dim last_row_keen As Long
    Dim n As Long ' new row number for result sheet
    Dim disaggregation As Variant
    Dim choice_percentage As Double
    Dim choice_count As Long
    Dim last_unified_row As Long
    Dim dis_count As Long
    
    Set sc_sheet = sheets("survey_choices")

    With sheets("dissagregation_setting")
        last_dis = .Cells(rows.count, 1).End(xlUp).row
        Set dis_rng = .Range("A2:B" & last_dis)
    End With

    With sheets("analysis_list")
        last_question = .Cells(rows.count, 1).End(xlUp).row
        Set question_rng = .Range("A2:B" & last_question)
    End With

    Set main_ws = sheets(find_main_data)
    last_row_main_data = main_ws.Cells(rows.count, 1).End(xlUp).row

    ' check if result sheet exist
    Call check_result_sheet(main_ws.Name)

    ' check if keen sheet exist
    If Not WorksheetExists("keen") Then
        Call create_sheet(main_ws.Name, "keen")
'        sheets("keen").Visible = xlVeryHidden
    End If

    Set keen_sheet = sheets("keen")
    sheets("keen").Cells.Clear

    Set result_sheet = sheets("result")

    last_row_choice = Worksheets("choices").Cells(rows.count, 1).End(xlUp).row
    last_row_survey_sheet = Worksheets("survey").Cells(rows.count, 1).End(xlUp).row
    last_row_survey_choices = Worksheets("survey_choices").Cells(rows.count, 1).End(xlUp).row
        
    ' if wieght column exist in the main data sheet, then extract its column name
    If has_weight Then
        weight_col_letter = gen_column_letter("weight", main_ws.Name)
    End If

    ' important variables:
    ' overall = true or false (true means ALL level disaggrigation)
    ' dis_value = disaggrigation level
    ' wgt = true or false (weight)
    ' m_type = number or select_one or select_multiple (measurement type)
    
    analysis_form.TextInfo.value = "Starting ... "
    
    Application.Wait (Now + 0.00001)
    
    Call remove_NA
    
    str_info = vbLf & analysis_form.TextInfo.value
   
    analysis_form.TextInfo.value = "Removed NAs" & str_info
            
    For Each dis_value In dis_rng.columns(1).Cells
        dis_str = CStr(dis_value)
        is_weight = dis_rng.columns(2).rows(dis_value.row - 1)
    
        If LCase(dis_str) = "all" Then
            overall = True
        Else
            overall = False
        End If
    
        If is_weight = "yes" Then
            If Not has_weight Then
                MsgBox "You have set weight for " & dis_value & " disaggregation level, " & vbCrLf & _
                       "but weight column dose not exist in the data!     " & vbCrLf & _
                       "Please add the weight column in the main data first.     " & vbCrLf & _
                       "The analysis proccesing will be stoped now.     ", vbCritical
                End
            End If
            wgt = True
        Else
            wgt = False
        End If
    
        dis_col_letter = gen_column_letter(dis_str, main_ws.Name)

        ' start looping through question_rng
        For Each measurement In question_rng.columns(1).Cells
        
            DoEvents
            measurement_str = CStr(measurement)
        
            If Public_module.CANCEL_PROCESS Then
                End
            End If
            
            measurement_col_letter = gen_column_letter(measurement_str, main_ws.Name)
            
            ' skip empty rows in the analysis list
            If measurement_col_letter = "" Then
                GoTo NextIteration
            End If
            
            ' skipp empty columns
            If WorksheetFunction.CountA(main_ws.columns(measurement_col_letter & ":" & measurement_col_letter)) = 1 Then
                str_info = vbLf & analysis_form.TextInfo.value
                txt = "Disaggregation level : " & dis_value & " > " & measurement & " skipped " & str_info
                txt = Replace(txt, "0", "")
                analysis_form.TextInfo.value = txt
                GoTo NextIteration
            End If
        
            ' show progress on the analysis user form
            If Len(str_info) > 2000 Then
                analysis_form.TextInfo.value = left(analysis_form.TextInfo.value, 1000)
            End If
       
            str_info = vbLf & analysis_form.TextInfo.value
        
            txt = "Disaggregation level : " & dis_value & " > " & measurement & str_info
            txt = Replace(txt, "0", "")
            analysis_form.TextInfo.value = txt
        
            
            measurement_type = question_rng.columns(2).rows(measurement.row - 1)
        
            If measurement_type = "integer" Or measurement_type = "decimal" Or measurement_type = "number" Then
                m_type = "number"
            ElseIf measurement_type = "select_multiple" Or measurement_type = "multiple" Then
                m_type = "select_multiple"
            ElseIf measurement_type = "select_one" Or measurement_type = "one" Then
                m_type = "select_one"
            Else
                ' skip the indicator because it has other data types or empty
                GoTo NextIteration
            End If
        
            ' skip the indicator because it is the same with disaggregation level
            If dis_str = measurement_str Then
                GoTo NextIteration
            End If
            
            ' skip invalid indicator indicator
            If dis_str <> "ALL" And measurement_col_letter = "" Then
                GoTo NextIteration
            End If
            
            On Error GoTo NextIteration
         
            ' start of the select case
            Select Case True

                ' numeric calculation
                ' case 1:
            Case overall And wgt And m_type = "number"
                
                Call inject_data(measurement_str, dis_str, wgt)
        
                Call add_mulitipication("C")
    
                last_row_keen = keen_sheet.Cells(rows.count, 1).End(xlUp).row
    
                sum_w2 = Application.WorksheetFunction.Sum(keen_sheet.Range("C2:C" & CStr(last_row_keen)))
                sum_weight = Application.WorksheetFunction.Sum(keen_sheet.Range("B2:B" & CStr(last_row_keen)))
                count_weight = Application.WorksheetFunction.count(keen_sheet.Range("B2:B" & CStr(last_row_keen)))
        
                n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1
    
                result_sheet.Cells(n, 1) = n - 1
                result_sheet.Cells(n, 2) = UCase(dis_str)
                result_sheet.Cells(n, 3) = UCase(dis_str)
                result_sheet.Cells(n, 4) = UCase(dis_str)
                result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 1)
                result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 1))
                result_sheet.Cells(n, 7) = count_weight
                result_sheet.Cells(n, 8) = "average"
                result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(sum_w2 / sum_weight, 1)
                result_sheet.Cells(n, 13) = "w"
        
                sheets("keen").Cells.Clear
       
                ' case 2:
            Case overall And Not wgt And m_type = "number"
                Call inject_data(measurement_str, dis_str, wgt)
        
                last_row_keen = keen_sheet.Cells(rows.count, 1).End(xlUp).row
    
                simple_count = Application.WorksheetFunction.count(keen_sheet.Range("A2:A" & CStr(last_row_keen)))
                simple_mean = Application.WorksheetFunction.Average(keen_sheet.Range("A2:A" & CStr(last_row_keen)))
    
                n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1

                result_sheet.Cells(n, 1) = n - 1
                result_sheet.Cells(n, 2) = UCase(dis_str)
                result_sheet.Cells(n, 3) = UCase(dis_str)
                result_sheet.Cells(n, 4) = UCase(dis_str)
                result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 1)
                result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 1))
                result_sheet.Cells(n, 7) = simple_count
                result_sheet.Cells(n, 8) = "average"
                result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(simple_mean, 1)
    
                sheets("keen").Cells.Clear
   
                ' case 3:
            Case Not overall And wgt And m_type = "number"
                Call inject_data(measurement_str, dis_str, wgt)
    
                ' convert to numeber
                keen_sheet.columns(2).TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
                                                   TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                                                   Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                                                   :=Array(1, 1), TrailingMinusNumbers:=True
    
                Call add_mulitipication("D")
                last_row_keen = keen_sheet.Cells(rows.count, 1).End(xlUp).row
    
                Set measurement_rng = keen_sheet.Range("A2:A" & last_row_keen)
                Set disaggregation_collection = unique_values(measurement_rng)
    
                For Each disaggregation In disaggregation_collection
    
                    sum_w2 = Application.WorksheetFunction.SumIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), _
                                                                 disaggregation, keen_sheet.Range("D2:D" & CStr(last_row_keen)))
                    sum_weight = Application.WorksheetFunction.SumIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), _
                                                                     disaggregation, keen_sheet.Range("C2:C" & CStr(last_row_keen)))
                    count_weight = Application.WorksheetFunction.CountIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), disaggregation)
                    n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1

                    result_sheet.Cells(n, 1) = n - 1
                    result_sheet.Cells(n, 2) = dis_str
                    result_sheet.Cells(n, 3) = disaggregation
                    result_sheet.Cells(n, 4) = choice_label(dis_str, CStr(disaggregation))
                    result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 2)
                    result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 2))
                    result_sheet.Cells(n, 7) = count_weight
                    result_sheet.Cells(n, 8) = "average"
                    result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(sum_w2 / sum_weight, 1)
                    result_sheet.Cells(n, 13) = "w"
        
                Next

                sheets("keen").Cells.Clear
   
                ' case 4:
            Case Not overall And Not wgt And m_type = "number"
                Call inject_data(measurement_str, dis_str, wgt)
    
                ' convert to numeber
                keen_sheet.columns(2).TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
                                                   TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                                                   Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
                                                   :=Array(1, 1), TrailingMinusNumbers:=True
    
                last_row_keen = keen_sheet.Cells(rows.count, 1).End(xlUp).row
                Set measurement_rng = keen_sheet.Range("A2:A" & last_row_keen)
                Set disaggregation_collection = unique_values(measurement_rng)
    
                For Each disaggregation In disaggregation_collection
        
                    simple_count = WorksheetFunction.CountIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), disaggregation)
                    simple_mean = WorksheetFunction.AverageIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), _
                                                              disaggregation, keen_sheet.Range("B2:B" & CStr(last_row_keen)))
            
                    n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1

                    result_sheet.Cells(n, 1) = n - 1
                    result_sheet.Cells(n, 2) = dis_str
                    result_sheet.Cells(n, 3) = disaggregation
                    result_sheet.Cells(n, 4) = choice_label(dis_str, CStr(disaggregation))
                    result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 2)
                    result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 2))
                    result_sheet.Cells(n, 7) = simple_count
                    result_sheet.Cells(n, 8) = "average"
                    result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(simple_mean, 1)
        
                Next

                sheets("keen").Cells.Clear

                ' select_one calculation
                ' case 5:
            Case overall And wgt And m_type = "select_one"
                Call inject_data(measurement_str, dis_str, wgt)
    
                last_row_keen = keen_sheet.Cells(keen_sheet.rows.count, 1).End(xlUp).row
    
                Set unique_choices = unique_values(keen_sheet.Range("A2:A" & last_row_keen))
    
                simple_count = Application.WorksheetFunction.CountA(keen_sheet.Range("A2:A" & CStr(last_row_keen)))
    
                If unique_choices.count > 0 Then
        
                    For i = 1 To unique_choices.count

                        sum_weight = Application.WorksheetFunction.Sum(keen_sheet.Range("B2:B" & CStr(last_row_keen)))
                        sum_w2 = Application.WorksheetFunction.SumIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), _
                                                                     unique_choices(i), keen_sheet.Range("B2:B" & CStr(last_row_keen)))
                
                        choice_percentage = Application.WorksheetFunction.Round(sum_w2 / sum_weight * 100, 1)
                        choice_count = Application.WorksheetFunction.Round(((last_row_keen - 1) * choice_percentage) / 100, 0)
            
                        n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1
            
                        result_sheet.Cells(n, 1) = n - 1
                        result_sheet.Cells(n, 2) = UCase(dis_str)
                        result_sheet.Cells(n, 3) = UCase(dis_str)
                        result_sheet.Cells(n, 4) = UCase(dis_str)
                        result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 1)
                        result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 1))
                        result_sheet.Cells(n, 7) = simple_count
                        result_sheet.Cells(n, 8) = "percentage"
                        result_sheet.Cells(n, 9) = choice_percentage
                        result_sheet.Cells(n, 10) = choice_count
                        result_sheet.Cells(n, 11) = unique_choices(i)
                        result_sheet.Cells(n, 12) = choice_label(Worksheets("keen").Cells(1, 1), CStr(unique_choices(i)))
                        result_sheet.Cells(n, 13) = "w"
                             
                    Next i
        
                End If
    
                sheets("keen").Cells.Clear

                ' case 6:
            Case overall And Not wgt And m_type = "select_one"
                Call inject_data(measurement_str, dis_str, wgt)
                last_row_keen = keen_sheet.Cells(rows.count, 1).End(xlUp).row
                Set measurement_rng = keen_sheet.Range("A2:A" & last_row_keen)
                Set unique_choices = unique_values(measurement_rng)
                simple_count = Application.WorksheetFunction.CountA(keen_sheet.Range("A2:A" & CStr(last_row_keen)))
    
                If unique_choices.count > 0 Then
        
                    For i = 1 To unique_choices.count
        
                        choice_count = Application.CountIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), unique_choices(i))
                
                        choice_percentage = Round(choice_count / simple_count * 100, 1)
            
                        n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1
    
                        result_sheet.Cells(n, 1) = n - 1
                        result_sheet.Cells(n, 2) = UCase(dis_str)
                        result_sheet.Cells(n, 3) = UCase(dis_str)
                        result_sheet.Cells(n, 4) = UCase(dis_str)
                        result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 1)
                        result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 1))
                        result_sheet.Cells(n, 7) = simple_count
                        result_sheet.Cells(n, 8) = "percentage"
                        result_sheet.Cells(n, 9) = choice_percentage
                        result_sheet.Cells(n, 10) = choice_count
                        result_sheet.Cells(n, 11) = unique_choices(i)
                        result_sheet.Cells(n, 12) = choice_label(Worksheets("keen").Cells(1, 1), CStr(unique_choices(i)))
        
                    Next
        
                End If

                sheets("keen").Cells.Clear

                ' case 7:
            Case Not overall And wgt And m_type = "select_one"
                Call inject_data(measurement_str, dis_str, wgt)
                
                ' Call unifier(True, True)
                
    
                last_row_keen = keen_sheet.Cells(rows.count, 1).End(xlUp).row
                
                ' last_unified_row = keen_sheet.Cells(rows.count, 6).End(xlUp).row
    
                Set measurement_rng = keen_sheet.Range("A2:A" & last_row_keen)
                Set disaggregation_collection = unique_values(measurement_rng)
  
'                Set unique_choices = unique_values(keen_sheet.Range("F2:F" & last_unified_row))
                Set unique_choices = unique_values(keen_sheet.Range("B2:B" & last_row_keen))
    
                ' loop through disaggregation options:
                For Each disaggregation In disaggregation_collection
    
                    If unique_choices.count > 0 Then
        
                        simple_count = Application.WorksheetFunction.CountIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), CStr(disaggregation))
                 
                        For i = 1 To unique_choices.count
               
                            sum_weight = Application.WorksheetFunction.SumIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), _
                                                                             CStr(disaggregation), keen_sheet.Range("C2:C" & CStr(last_row_keen)))
               
                            sum_w2 = Application.WorksheetFunction.SumIfs(keen_sheet.Range("C2:C" & CStr(last_row_keen)), _
                                                                          keen_sheet.Range("A2:A" & CStr(last_row_keen)), CStr(disaggregation), _
                                                                          keen_sheet.Range("B2:B" & CStr(last_row_keen)), unique_choices(i))
                              
                            choice_percentage = Application.WorksheetFunction.Round(sum_w2 / sum_weight * 100, 1)
                            choice_count = Application.WorksheetFunction.Round((simple_count * choice_percentage) / 100, 0)
               
                            n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1
               
                            result_sheet.Cells(n, 1) = n - 1
                            result_sheet.Cells(n, 2) = dis_str
                            result_sheet.Cells(n, 3) = disaggregation
                            result_sheet.Cells(n, 4) = choice_label(dis_str, CStr(disaggregation))
                            result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 2)
                            result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 2))
                            result_sheet.Cells(n, 7) = simple_count
                            result_sheet.Cells(n, 8) = "percentage"
                            result_sheet.Cells(n, 9) = choice_percentage
                            result_sheet.Cells(n, 10) = choice_count
                            result_sheet.Cells(n, 11) = unique_choices(i)
                            result_sheet.Cells(n, 12) = choice_label(Worksheets("keen").Cells(1, 2), CStr(unique_choices(i)))
                            result_sheet.Cells(n, 13) = "w"
                                          
                        Next i
            
                    End If
        
                Next

                sheets("keen").Cells.Clear

                ' case 8:
            Case Not overall And Not wgt And m_type = "select_one"
                Call inject_data(measurement_str, dis_str, wgt)
                
                ' Call unifier(True, False)
    
                last_row_keen = keen_sheet.Cells(rows.count, 1).End(xlUp).row
                ' last_unified_row = keen_sheet.Cells(rows.count, 6).End(xlUp).row
    
                Set measurement_rng = keen_sheet.Range("A2:A" & last_row_keen)
                Set disaggregation_collection = unique_values(measurement_rng)
  
                ' Set unique_choices = unique_values(keen_sheet.Range("F2:F" & last_unified_row))
                Set unique_choices = unique_values(keen_sheet.Range("B2:B" & last_row_keen))
                
                ' loop through disaggregation options:
                For Each disaggregation In disaggregation_collection
    
                    If unique_choices.count > 0 Then
            
                        simple_count = Application.WorksheetFunction.CountIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), CStr(disaggregation))
            
                        For i = 1 To unique_choices.count
               
                            dis_count = Application.WorksheetFunction.CountIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), _
                                                                              CStr(disaggregation))
                
                            choice_count = Application.WorksheetFunction.CountIfs(keen_sheet.Range("A2:A" & CStr(last_row_keen)), _
                                            CStr(disaggregation), keen_sheet.Range("B2:B" & CStr(last_row_keen)), unique_choices(i))
                                  
                            choice_percentage = Application.WorksheetFunction.Round(choice_count / dis_count * 100, 1)
                
                            n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1
                
                            result_sheet.Cells(n, 1) = n - 1
                            result_sheet.Cells(n, 2) = dis_str
                            result_sheet.Cells(n, 3) = disaggregation
                            result_sheet.Cells(n, 4) = choice_label(dis_str, CStr(disaggregation))
                            result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 2)
                            result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 2))
                            result_sheet.Cells(n, 7) = simple_count
                            result_sheet.Cells(n, 8) = "percentage"
                            result_sheet.Cells(n, 9) = choice_percentage
                            result_sheet.Cells(n, 10) = choice_count
                            result_sheet.Cells(n, 11) = unique_choices(i)
                            result_sheet.Cells(n, 12) = choice_label(Worksheets("keen").Cells(1, 2), CStr(unique_choices(i)))

                        Next i
            
                    End If
    
                Next

                sheets("keen").Cells.Clear

                ' select_multiple calculation
                ' case 9:
            Case overall And wgt And m_type = "select_multiple"
                Call inject_data(measurement_str, dis_str, wgt)
                Call unifier(False, True)
    
                last_unified_row = keen_sheet.Cells(rows.count, 6).End(xlUp).row
                Set unique_choices = unique_values(keen_sheet.Range("F2:F" & last_unified_row))
    
                last_row_keen = keen_sheet.Cells(keen_sheet.rows.count, 1).End(xlUp).row
                simple_count = Application.WorksheetFunction.CountA(keen_sheet.Range("A2:A" & CStr(last_row_keen)))
    
                If unique_choices.count > 0 Then
        
                    For i = 1 To unique_choices.count

                        sum_weight = Application.WorksheetFunction.Sum(keen_sheet.Range("B2:B" & CStr(last_row_keen)))
                        sum_w2 = Application.WorksheetFunction.SumIf(keen_sheet.Range("F2:F" & CStr(last_unified_row)), _
                                                                     unique_choices(i), keen_sheet.Range("G2:G" & CStr(last_unified_row)))
                
                        choice_percentage = Application.WorksheetFunction.Round(sum_w2 / sum_weight * 100, 1)
                        choice_count = Application.WorksheetFunction.Round(((last_row_keen - 1) * choice_percentage) / 100, 0)
            
                        n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1
            
                        result_sheet.Cells(n, 1) = n - 1
                        result_sheet.Cells(n, 2) = UCase(dis_str)
                        result_sheet.Cells(n, 3) = UCase(dis_str)
                        result_sheet.Cells(n, 4) = UCase(dis_str)
                        result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 1)
                        result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 1))
                        result_sheet.Cells(n, 7) = simple_count
                        result_sheet.Cells(n, 8) = "percentage"
                        result_sheet.Cells(n, 9) = choice_percentage
                        result_sheet.Cells(n, 10) = choice_count
                        result_sheet.Cells(n, 11) = unique_choices(i)
                        result_sheet.Cells(n, 12) = choice_label(Worksheets("keen").Cells(1, 1), CStr(unique_choices(i)))
                        result_sheet.Cells(n, 13) = "w"
                             
                    Next i
        
                End If
    
                sheets("keen").Cells.Clear
    
                ' case 10:
            Case overall And Not wgt And m_type = "select_multiple"
                Call inject_data(measurement_str, dis_str, wgt)
                Call unifier(False, False)
    
                last_unified_row = keen_sheet.Cells(rows.count, 6).End(xlUp).row
                Set unique_choices = unique_values(keen_sheet.Range("F2:F" & last_unified_row))
                last_row_keen = keen_sheet.Cells(keen_sheet.rows.count, 1).End(xlUp).row
                simple_count = Application.WorksheetFunction.CountA(keen_sheet.Range("A2:A" & CStr(last_row_keen)))
            
                If unique_choices.count > 0 Then
        
                    For i = 1 To unique_choices.count
              
                        choice_count = Application.WorksheetFunction.CountIf(keen_sheet.Range("F2:F" & CStr(last_unified_row)), unique_choices(i))
                
                        choice_percentage = Application.WorksheetFunction.Round(choice_count / (last_row_keen - 1) * 100, 1)
                       
                        n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1
            
                        result_sheet.Cells(n, 1) = n - 1
                        result_sheet.Cells(n, 2) = UCase(dis_str)
                        result_sheet.Cells(n, 3) = UCase(dis_str)
                        result_sheet.Cells(n, 4) = UCase(dis_str)
                        result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 1)
                        result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 1))
                        result_sheet.Cells(n, 7) = simple_count
                        result_sheet.Cells(n, 8) = "percentage"
                        result_sheet.Cells(n, 9) = choice_percentage
                        result_sheet.Cells(n, 10) = choice_count
                        result_sheet.Cells(n, 11) = unique_choices(i)
                        result_sheet.Cells(n, 12) = choice_label(Worksheets("keen").Cells(1, 1), CStr(unique_choices(i)))
                             
                    Next i
        
                End If

                ' case 11:
            Case Not overall And wgt And m_type = "select_multiple"
                Call inject_data(measurement_str, dis_str, wgt)
                Call unifier(True, True)
    
                last_row_keen = keen_sheet.Cells(rows.count, 1).End(xlUp).row
                last_unified_row = keen_sheet.Cells(rows.count, 6).End(xlUp).row
    
                Set measurement_rng = keen_sheet.Range("A2:A" & last_row_keen)
                Set disaggregation_collection = unique_values(measurement_rng)
  
                Set unique_choices = unique_values(keen_sheet.Range("F2:F" & last_unified_row))

                ' loop through disaggregation options:
                For Each disaggregation In disaggregation_collection
    
                    If unique_choices.count > 0 Then
        
                        simple_count = Application.WorksheetFunction.CountIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), CStr(disaggregation))
                 
                        For i = 1 To unique_choices.count
               
                            sum_weight = Application.WorksheetFunction.SumIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), _
                                                                             CStr(disaggregation), keen_sheet.Range("C2:C" & CStr(last_row_keen)))
               
                            sum_w2 = Application.WorksheetFunction.SumIfs(keen_sheet.Range("G2:G" & CStr(last_unified_row)), _
                                                                          keen_sheet.Range("E2:E" & CStr(last_unified_row)), CStr(disaggregation), _
                                                                          keen_sheet.Range("F2:F" & CStr(last_unified_row)), unique_choices(i))
                              
                            choice_percentage = Application.WorksheetFunction.Round(sum_w2 / sum_weight * 100, 1)
                            choice_count = Application.WorksheetFunction.Round((simple_count * choice_percentage) / 100, 0)
               
                            n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1
               
                            result_sheet.Cells(n, 1) = n - 1
                            result_sheet.Cells(n, 2) = dis_str
                            result_sheet.Cells(n, 3) = disaggregation
                            result_sheet.Cells(n, 4) = choice_label(dis_str, CStr(disaggregation))
                            result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 2)
                            result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 2))
                            result_sheet.Cells(n, 7) = simple_count
                            result_sheet.Cells(n, 8) = "percentage"
                            result_sheet.Cells(n, 9) = choice_percentage
                            result_sheet.Cells(n, 10) = choice_count
                            result_sheet.Cells(n, 11) = unique_choices(i)
                            result_sheet.Cells(n, 12) = choice_label(Worksheets("keen").Cells(1, 2), CStr(unique_choices(i)))
                            result_sheet.Cells(n, 13) = "w"
                                          
                        Next i
            
                    End If
        
                Next

                sheets("keen").Cells.Clear

                ' case 12:
            Case Not overall And Not wgt And m_type = "select_multiple"
                Call inject_data(measurement_str, dis_str, wgt)
                Call unifier(True, False)
    
                last_row_keen = keen_sheet.Cells(rows.count, 1).End(xlUp).row
                last_unified_row = keen_sheet.Cells(rows.count, 6).End(xlUp).row
    
                Set measurement_rng = keen_sheet.Range("A2:A" & last_row_keen)
                Set disaggregation_collection = unique_values(measurement_rng)
  
                Set unique_choices = unique_values(keen_sheet.Range("F2:F" & last_unified_row))
    
                ' loop through disaggregation options:
                For Each disaggregation In disaggregation_collection
    
                    If unique_choices.count > 0 Then
            
                        simple_count = Application.WorksheetFunction.CountIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), CStr(disaggregation))
            
                        For i = 1 To unique_choices.count
               
                            dis_count = Application.WorksheetFunction.CountIf(keen_sheet.Range("A2:A" & CStr(last_row_keen)), _
                                                                              CStr(disaggregation))
                
                            choice_count = Application.WorksheetFunction.CountIfs(keen_sheet.Range("E2:E" & CStr(last_unified_row)), _
                                                                                  CStr(disaggregation), keen_sheet.Range("F2:F" & CStr(last_unified_row)), unique_choices(i))
                                  
                            choice_percentage = Application.WorksheetFunction.Round(choice_count / dis_count * 100, 1)
                
                            n = result_sheet.Cells(rows.count, 1).End(xlUp).row + 1
                
                            result_sheet.Cells(n, 1) = n - 1
                            result_sheet.Cells(n, 2) = dis_str
                            result_sheet.Cells(n, 3) = disaggregation
                            result_sheet.Cells(n, 4) = choice_label(dis_str, CStr(disaggregation))
                            result_sheet.Cells(n, 5) = Worksheets("keen").Cells(1, 2)
                            result_sheet.Cells(n, 6) = var_label(Worksheets("keen").Cells(1, 2))
                            result_sheet.Cells(n, 7) = simple_count
                            result_sheet.Cells(n, 8) = "percentage"
                            result_sheet.Cells(n, 9) = choice_percentage
                            result_sheet.Cells(n, 10) = choice_count
                            result_sheet.Cells(n, 11) = unique_choices(i)
                            result_sheet.Cells(n, 12) = choice_label(Worksheets("keen").Cells(1, 2), CStr(unique_choices(i)))
                
                        Next i
            
                    End If
    
                Next

                sheets("keen").Cells.Clear

                ' end of the select case
            End Select

NextIteration:

            sheets("keen").Cells.Clear
            
        Next  ' loop for question_rng
   
    Next  ' loop for disaggrigation

    Application.DisplayAlerts = False
    
    If WorksheetExists("keen") Then
'        sheets("keen").Delete
    End If
    
    Application.DisplayAlerts = True

    Application.ScreenUpdating = True
      
    str_info = vbLf & analysis_form.TextInfo.value

    txt = "Generating Datamerge... " & str_info

    analysis_form.TextInfo.value = txt
    
End Sub

' if we need to apply weighting a column will be generated by the name w2 = numeric_value * weight
Sub add_mulitipication(target_col As String)
    Dim last_row As Long
    last_row = Worksheets("keen").Cells(rows.count, 1).End(xlUp).row
    Worksheets("keen").Range(target_col & "1").FormulaR1C1 = "w2"
    Application.CutCopyMode = False
    Worksheets("keen").Range(target_col & "2").FormulaR1C1 = "=RC[-2]*RC[-1]"
    Worksheets("keen").Range(target_col & "2").AutoFill _
        Destination:=Worksheets("keen").Range(target_col & "2:" & target_col & CStr(last_row))
End Sub

Sub check_result_sheet(sheet_name As String)
    ' check if keen sheet exist
    If Not WorksheetExists("result") Then
        Call create_sheet(sheet_name, "result")
        sheets("result").Cells(1, 1) = "row"
        sheets("result").Cells(1, 2) = "disaggregation"
        sheets("result").Cells(1, 3) = "disaggregation value"
        sheets("result").Cells(1, 4) = "disaggregation label"
        sheets("result").Cells(1, 5) = "variable"
        sheets("result").Cells(1, 6) = "variable label"
        sheets("result").Cells(1, 7) = "valid numbers"
        sheets("result").Cells(1, 8) = "measurement type"
        sheets("result").Cells(1, 9) = "measurement value"
        sheets("result").Cells(1, 10) = "count"
        sheets("result").Cells(1, 11) = "choice"
        sheets("result").Cells(1, 12) = "choice label"
        sheets("result").Cells(1, 13) = "weight"
        
        sheets("result").columns(1).ColumnWidth = 6
        sheets("result").columns(2).ColumnWidth = 15
        sheets("result").columns(3).ColumnWidth = 18
        sheets("result").columns(4).ColumnWidth = 25
        sheets("result").columns(5).ColumnWidth = 15
        sheets("result").columns(6).ColumnWidth = 45
        sheets("result").columns(7).ColumnWidth = 15
        sheets("result").columns(8).ColumnWidth = 15
        sheets("result").columns(9).ColumnWidth = 20
        sheets("result").columns(10).ColumnWidth = 10
        sheets("result").columns(11).ColumnWidth = 15
        sheets("result").columns(12).ColumnWidth = 45
        sheets("result").columns(13).ColumnWidth = 7

    End If
End Sub

Public Function unique_valuesX(rng As Range) As Collection
    Dim dic As Object, c As Range, h, tmp As String
    Dim unique_collection As New Collection
    
    Set dic = CreateObject("scripting.dictionary")
    For Each c In rng
        tmp = Trim(c.value)
        If Len(tmp) > 0 Then dic(tmp) = dic(tmp) + 1
    Next c
    
    For Each h In dic.Keys
        unique_collection.Add CStr(h)
    Next h
    
    Set unique_valuesX = unique_collection
End Function

' remove null rows from keen sheet
Sub delete_blank_rows(col_number As Long)

    Dim rng As Range
    Dim str_delete As String
    Dim last_row As Long
    Dim last_keen As Long
    Dim last_measurement As Long
    
    last_row = sheets("keen").Range("A" & sheets("keen").rows.count).End(xlUp).row
    Set rng = sheets("keen").Range("A1:C" & last_row)
    
    last_keen = rng.rows.count
    rng.Sort rng.columns(col_number), , , , , , , Header:=xlYes
    
    last_measurement = sheets("keen").Cells(rows.count, col_number).End(xlUp).row
    
    str_delete = CStr(last_measurement + 1) & ":" & last_keen
    
    If last_measurement < last_keen Then
        sheets("keen").rows(str_delete).Delete Shift:=xlUp
    End If
    
End Sub

' convert multi select column into single column in the keen sheet
' new disagregation column is E
' new measurement column is F
' new weight column is G
Sub unifier(dis As Boolean, wgh As Boolean)
    
    Dim i As Long, j As Long
    Dim arr() As String
    Dim LastRow As Long
    Dim endRow As Long
    Dim ws As Worksheet
    Set ws = sheets("keen")
    
    LastRow = ws.Cells(ws.rows.count, "A").End(xlUp).row
    
    If dis Then
        ' count total number of choices
        For i = 1 To LastRow
            arr = Split(ws.Cells(i, 2), " ")
            endRow = endRow + (UBound(arr) - LBound(arr) + 1)
        Next i
    Else
        ' count total number of choices
        For i = 1 To LastRow
            arr = Split(ws.Cells(i, 1), " ")
            endRow = endRow + (UBound(arr) - LBound(arr) + 1)
        Next i
    End If
   
    ' convert to single value based on condition
    ' write cells, begining from last
    Select Case True
    
    Case dis And wgh
        For i = LastRow To 1 Step -1
            arr = Split(ws.Cells(i, 2), " ")
            For j = LBound(arr) To UBound(arr)
                ws.Cells(endRow, 5) = ws.Cells(i, 1)
                ws.Cells(endRow, 6) = arr(j)
                ws.Cells(endRow, 7) = ws.Cells(i, 3)
                endRow = endRow - 1
            Next j
        Next i
            
    Case dis And Not wgh
        For i = LastRow To 1 Step -1
            arr = Split(ws.Cells(i, 2), " ")
            For j = LBound(arr) To UBound(arr)
                ws.Cells(endRow, 5) = ws.Cells(i, 1)
                ws.Cells(endRow, 6) = arr(j)
                endRow = endRow - 1
            Next j
        Next i

    Case Not dis And wgh
        For i = LastRow To 1 Step -1
            arr = Split(ws.Cells(i, 1), " ")
            For j = LBound(arr) To UBound(arr)
                ws.Cells(endRow, 6) = arr(j)
                ws.Cells(endRow, 7) = ws.Cells(i, 2)
                endRow = endRow - 1
            Next j
        Next i
            
    Case Else
        For i = LastRow To 1 Step -1
            arr = Split(ws.Cells(i, 1), " ")
            For j = LBound(arr) To UBound(arr)
                ws.Cells(endRow, 6) = arr(j)
                endRow = endRow - 1
            Next j
        Next i
    
    End Select

End Sub

' check if main data sheet has weight column or not
Function has_weight() As Boolean
    Dim main_ws As Worksheet
    Dim last_main_col_letter As String
    Dim cel As Variant
    
    Set main_ws = sheets(find_main_data)
    
    last_main_col_letter = Split(main_ws.Cells.Find(What:="*", After:=[a1], SearchOrder:=xlByColumns, _
                                                    SearchDirection:=xlPrevious).Cells.Address(1, 0), "$")(0)
    
    For Each cel In main_ws.Range("A1:" & last_main_col_letter & 1)
        If cel = "weight" Then
            has_weight = True
            Exit For
        Else
            has_weight = False
        End If
    Next
    
    Dim a As Long
    
End Function

Sub cc()
' testing
sheets("keen").Cells.Clear

'        If WorksheetExists("keen") Then
'            sheets("keen").Visible = xlSheetHidden
'            sheets("keen").Delete
'        End If

Call inject_data("es_nfi_vulnerability_index_cat", "province", True)
Call unifier(True, True)
End Sub


' get the required data into keen sheet and delete the blank measurement rows
Sub inject_data(measurement As String, disaggregation As String, weight As Boolean)
    Dim ws As Worksheet
    Dim measurement_col_letter As String
    Dim dis_col_lettera As String
    Dim dis_col_letter As String
    Dim weight_col_letter As String
    
'    Set ws = sheets(Public_module.DATA_SHEET)
    Set ws = sheets(find_main_data)
    
    If LCase(disaggregation) = "all" Then
        measurement_col_letter = gen_column_letter(measurement, ws.Name)
        sheets("keen").columns("A") = ws.columns(measurement_col_letter).Value2
    Else
        dis_col_letter = gen_column_letter(disaggregation, ws.Name)
        measurement_col_letter = gen_column_letter(measurement, ws.Name)
        sheets("keen").columns("A") = ws.columns(dis_col_letter).Value2
        sheets("keen").columns("B") = ws.columns(measurement_col_letter).Value2
    End If

    If weight And LCase(disaggregation) = "all" Then
        weight_col_letter = gen_column_letter("weight", ws.Name)
        sheets("keen").columns("B") = ws.columns(weight_col_letter).Value2
    ElseIf weight And LCase(disaggregation) <> "all" Then
        weight_col_letter = gen_column_letter("weight", ws.Name)
        sheets("keen").columns("C") = ws.columns(weight_col_letter).Value2
    End If
    
    If LCase(disaggregation) = "all" Then
        Call delete_blank_rows(1)
    Else
        Call delete_blank_rows(2)
    End If
    
End Sub

' return the label of main measurement
Function var_label(var As String) As String
    On Error GoTo errHandler
    
    Dim last_row_survey As Long
    Dim v_label As String
    
    last_row_survey = Worksheets("survey").Cells(rows.count, 1).End(xlUp).row
    v_label = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey), _
                                        WorksheetFunction.Match(var, sheets("survey").Range("B2:B" & last_row_survey), 0))
                
    If v_label = vbNullString Then
        var_label = var
        
    Else
        var_label = v_label
    End If
    Exit Function
                
errHandler:
    var_label = var
    
End Function

' return the label of choice, if not not found return the original choice value
Function choice_label(question As String, choice As String) As String

    On Error GoTo errHandler
    
    Dim ws_sc As Worksheet
    Set ws_sc = sheets("survey_choices")
    Dim last_row_survey_choices As Long
    Dim question_choice As String
    question_choice = question & choice
    
    last_row_survey_choices = ws_sc.Cells(rows.count, 1).End(xlUp).row
     
    choice_label = WorksheetFunction.Index(ws_sc.Range("E2:E" & last_row_survey_choices), _
                                           WorksheetFunction.Match(question_choice, ws_sc.Range("F2:F" & last_row_survey_choices), 0))

    Exit Function

errHandler:
    choice_label = choice

End Function

Sub remove_NA()
    
    Dim ws As Worksheet
    Set ws = sheets(Public_module.DATA_SHEET)
    
    ws.Cells.Replace What:="NA", replacement:="", LookAt:=xlWhole, SearchOrder _
                     :=xlByColumns, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False

End Sub


