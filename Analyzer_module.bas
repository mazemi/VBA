Attribute VB_Name = "Analyzer_module"
Option Explicit
Global WITH_WEIGHT As Boolean

Sub do_analize()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim result_sheet As Worksheet
    Dim main_ws As Worksheet
    Dim keen_ws As Worksheet
    Dim dis_arr As Variant
    Dim var_arr As Variant
    Dim header_arr() As Variant
    Dim filtered_arr() As String
    Dim m As Long
    Dim last_row_main_data As Long
    Dim last_col As Long
    Dim dis_collection As New Collection
    Dim data_rows As String
    Dim main_rng As Range
    Dim cr_rng As Range
    Dim last_dis As Long
    Dim last_col_letter As String
    Dim i As Long
    Dim j As Long
    Dim last_var As Long
    Dim var_col_letter As String
    Dim str_info As String
    Dim txt As String
    Dim last_row_result As Long

    WITH_WEIGHT = False

        analysis_form.TextInfo.value = "Starting ... "
        Application.Wait (Now + 0.00001)

        Call remove_NA

        DoEvents
        str_info = vbLf & analysis_form.TextInfo.value
        analysis_form.TextInfo.value = "Removed NAs" & str_info
        str_info = vbLf & analysis_form.TextInfo.value
    
        Set main_ws = sheets(find_main_data)
        Call remove_empty_col

        With sheets("disaggregation_setting")
            last_dis = .Cells(Rows.count, 1).End(xlUp).Row
            dis_arr = .Range("A2:B" & last_dis)
        End With

        last_row_main_data = main_ws.Cells(Rows.count, find_uuid_coln).End(xlUp).Row
        data_rows = CStr(2) & ":" & CStr(last_row_main_data)
        last_col = main_ws.Cells(1, Columns.count).End(xlToLeft).Column
        last_col_letter = number_to_letter(last_col, main_ws)

        For i = 1 To UBound(dis_arr, 1)
            If dis_arr(i, 1) <> "ALL" Then
                dis_collection.Add dis_arr(i, 1)
            End If

            If dis_arr(i, 2) = "yes" Then
                WITH_WEIGHT = True
                End If
            Next i

            If WITH_WEIGHT Then
                If Not has_weight Then
                    MsgBox "You are going to to implement wieghting in your analysis, " & vbCrLf & _
                           "but weight column dose not exist in the data!     " & vbCrLf & _
                           "Please add the weight column in the main data first.     ", vbCritical
                    End
                End If
            End If

            With sheets("analysis_list")
                last_var = .Cells(Rows.count, 1).End(xlUp).Row
                var_arr = .Range("A2:B" & last_var)
            End With

            If worksheet_exists("result") Then
                sheets("result").Delete
            End If

            Call check_result_sheet("analysis_list")
            Set result_sheet = wb.sheets("result")

            If Not worksheet_exists("keen") Then
                Call create_sheet(main_ws.Name, "keen")
                sheets("keen").Visible = xlVeryHidden
            End If

            Set keen_ws = sheets("keen")
            keen_ws.Cells.Clear

            If WITH_WEIGHT Then
                keen_ws.Columns("C:M").NumberFormat = "@"
            Else
                keen_ws.Columns("B:M").NumberFormat = "@"
            End If

            If Not worksheet_exists("temp_sheet") Then
                Call create_sheet(main_ws.Name, "temp_sheet")
                sheets("temp_sheet").Visible = xlVeryHidden
            End If

            sheets("temp_sheet").Cells.Clear
            sheets("temp_sheet").Rows(1).value = main_ws.Rows(1).value

            ' keen header:
            If WITH_WEIGHT Then
                keen_ws.Cells(1, 2) = "weight"
            End If
            If dis_collection.count > 0 Then
                If WITH_WEIGHT Then
                    For m = 1 To dis_collection.count
                        keen_ws.Cells(1, m + 2) = dis_collection.item(m)
                    Next m
                Else
                    For m = 1 To dis_collection.count
                        keen_ws.Cells(1, m + 1) = dis_collection.item(m)
                    Next m
                End If
            End If

            header_arr = main_ws.Range(main_ws.Cells(1, 1), main_ws.Cells(1, 1).End(xlToRight)).Value2
            header_arr = Application.Transpose(Application.Transpose(header_arr))

            Set main_rng = main_ws.Range("A1:" & last_col_letter & last_row_main_data)
            Set cr_rng = sheets("temp_sheet").Range("A1:" & last_col_letter & 2)

            ' main loop:
            For i = 1 To UBound(var_arr, 1)
                DoEvents
                ' show progress on the analysis user form
                If Len(str_info) > 2000 Then
                    analysis_form.TextInfo.value = left(analysis_form.TextInfo.value, 1000)
                End If

                str_info = vbLf & analysis_form.TextInfo.value
 
                txt = "Proccessing: " & CStr(var_arr(i, 1)) & str_info
                txt = Replace(txt, "0", "")
                analysis_form.TextInfo.value = txt
            
                filtered_arr = Filter(header_arr, var_arr(i, 1), True, vbTextCompare)
                If UBound(filtered_arr) = -1 Then
                    GoTo NextIteration
                End If
    
                sheets("temp_sheet").Rows(2).Clear
                var_col_letter = gen_column_letter(CStr(var_arr(i, 1)), "temp_sheet")
    
                If var_col_letter = "" Then
                    Debug.Print "column not exist in the main data."
                    GoTo NextIteration
                End If
    
                sheets("temp_sheet").Range(var_col_letter & 2) = "<>"
    
                keen_ws.Range("A1") = var_arr(i, 1)
                keen_ws.Rows(data_rows).Clear
    
                On Error GoTo criticalerrHandler
                main_rng.AdvancedFilter xlFilterCopy, cr_rng, keen_ws.Range("A1").CurrentRegion
                On Error GoTo 0
    
                If IsEmpty(keen_ws.Range("A2")) Then
                    GoTo NextIteration
                End If
    
                If var_arr(i, 2) = "select_multiple" Then
                    Call unify_data
                End If
    
                If var_arr(i, 2) = "integer" Or var_arr(i, 2) = "decimal" Then
                    Call calculate_numeric
                End If
    
                If var_arr(i, 2) = "select_one" Then
                    Call calculate_nominal
                End If
    
                If var_arr(i, 2) = "select_multiple" Then
                    Call calculate_nominal_multipe
                End If
                    
NextIteration:
            Next i

            last_row_result = result_sheet.Cells(Rows.count, 1).End(xlUp).Row

            Call delete_un_selected_choices
            
            If check_gender_column Then
                Call gender_disaggregation
            End If
            
            If last_row_result > 2 Then
                Call make_header_order
            End If

            wb.Save
            Exit Sub

criticalerrHandler:
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True

            End

        End Sub

Sub calculate_numeric()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim result_sheet As Worksheet
    Dim last_col As Long
    Dim last_col_letter As String
    Dim i As Long, j As Long
    Dim s As Long
    Dim dis_arr2 As Variant
    Dim weight_arr() As Double
    Dim simple_arr() As Double
    Dim disagregation_arr() As String
    Dim dis_value_count As Long
    Dim new_col_letter As String
    Dim last_row As Long
    Dim last_dis As Long
    Dim n As Long
    Dim col_n As Long
    Dim unique_arr As Variant
    Dim v As Variant
    Dim k As Long

    Set ws = sheets("keen")
    Set result_sheet = sheets("result")
    last_col = ws.Cells(1, Columns.count).End(xlToLeft).Column
    last_col_letter = number_to_letter(last_col, ws)
    new_col_letter = number_to_letter(last_col + 1, ws)
    last_row = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ws.Columns("C:M").NumberFormat = "@"

    With sheets("disaggregation_setting")
        last_dis = .Cells(Rows.count, 1).End(xlUp).Row
        dis_arr2 = .Range("A2:C" & last_dis)
    End With

    Dim arr As Variant
    arr = ws.Range("A1").CurrentRegion

    For i = 1 To UBound(dis_arr2, 1)

        Erase simple_arr
        ReDim simple_arr(1 To UBound(arr, 1) - 1)
        Erase weight_arr
        ReDim weight_arr(1 To UBound(arr, 1) - 1)
    
        If dis_arr2(i, 1) = "ALL" And dis_arr2(i, 2) = "yes" Then
            For j = 2 To UBound(arr, 1)
                simple_arr(j - 1) = arr(j, 1) * arr(j, 2)
                weight_arr(j - 1) = arr(j, 2)
            Next j

            With result_sheet
                n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                .Cells(n, 1) = n - 1
                .Cells(n, 2) = "ALL"
                .Cells(n, 3) = "ALL"
                .Cells(n, 4) = "ALL"
                .Cells(n, 5) = arr(1, 1)
                .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                .Cells(n, 7) = UBound(arr, 1) - 1
                .Cells(n, 8) = "average"
                .Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.sum(simple_arr) / WorksheetFunction.sum(weight_arr), 1)
                .Cells(n, 13) = "w"
            
                n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                .Cells(n, 1) = n - 1
                .Cells(n, 2) = "ALL"
                .Cells(n, 3) = "ALL"
                .Cells(n, 4) = "ALL"
                .Cells(n, 5) = arr(1, 1)
                .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                .Cells(n, 7) = UBound(arr, 1) - 1
                .Cells(n, 8) = "median"
                .Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.median(simple_arr), 1)
                .Cells(n, 13) = "w"
            End With

        
        ElseIf dis_arr2(i, 1) = "ALL" And dis_arr2(i, 2) = "no" Then
            For j = 2 To UBound(arr, 1)
                simple_arr(j - 1) = arr(j, 1)
            Next j
            With result_sheet
                n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                .Cells(n, 1) = n - 1
                .Cells(n, 2) = "ALL"
                .Cells(n, 3) = "ALL"
                .Cells(n, 4) = "ALL"
                .Cells(n, 5) = arr(1, 1)
                .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                .Cells(n, 7) = UBound(arr, 1) - 1
                .Cells(n, 8) = "average"
                .Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.Average(simple_arr), 1)
    
                n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                .Cells(n, 1) = n - 1
                .Cells(n, 2) = "ALL"
                .Cells(n, 3) = "ALL"
                .Cells(n, 4) = "ALL"
                .Cells(n, 5) = arr(1, 1)
                .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                .Cells(n, 7) = UBound(arr, 1) - 1
                .Cells(n, 8) = "median"
                .Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.median(simple_arr), 1)
            End With

        
        ElseIf dis_arr2(i, 1) <> "ALL" And dis_arr2(i, 2) = "yes" Then
            col_n = gen_column_number(CStr(dis_arr2(i, 1)), "keen")
            Erase disagregation_arr
            ReDim disagregation_arr(1 To UBound(arr, 1) - 1)

            For j = 2 To UBound(arr, 1)
                disagregation_arr(j - 1) = arr(j, col_n)
            Next j
        
            unique_arr = get_unique(disagregation_arr)
            For Each v In unique_arr
                dis_value_count = count_in_array(disagregation_arr, v)
                '            Debug.Print v, dis_value_count
                Erase simple_arr
                ReDim simple_arr(1 To dis_value_count)
                Erase weight_arr
                ReDim weight_arr(1 To dis_value_count)
            
                k = 0
                For j = 2 To UBound(arr, 1)
                    If v = arr(j, col_n) Then
                        simple_arr(k + 1) = arr(j, 1) * arr(j, 2)
                        weight_arr(k + 1) = arr(j, 2)
                        k = k + 1
                    End If
                Next j
                
                With result_sheet
                    n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                    .Cells(n, 1) = n - 1
                    .Cells(n, 2) = dis_arr2(i, 1)
                    .Cells(n, 3) = v
                    .Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                    .Cells(n, 5) = arr(1, 1)
                    .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                    .Cells(n, 7) = dis_value_count
                    .Cells(n, 8) = "average"
                    .Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.sum(simple_arr) / WorksheetFunction.sum(weight_arr), 1)
                    .Cells(n, 13) = "w"
                
                    n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                    .Cells(n, 1) = n - 1
                    .Cells(n, 2) = dis_arr2(i, 1)
                    .Cells(n, 3) = v
                    .Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                    .Cells(n, 5) = arr(1, 1)
                    .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                    .Cells(n, 7) = dis_value_count
                    .Cells(n, 8) = "median"
                    .Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.median(simple_arr), 1)
                    .Cells(n, 13) = "w"
                End With

            Next v
        
        ElseIf dis_arr2(i, 1) <> "ALL" And dis_arr2(i, 2) = "no" Then
            col_n = gen_column_number(CStr(dis_arr2(i, 1)), "keen")
            Erase disagregation_arr
            ReDim disagregation_arr(1 To UBound(arr, 1) - 1)
            For j = 2 To UBound(arr, 1)
                disagregation_arr(j - 1) = arr(j, col_n)
            Next j

            unique_arr = get_unique(disagregation_arr)
            For Each v In unique_arr
                dis_value_count = count_in_array(disagregation_arr, v)
                '            Debug.Print v, dis_value_count
                Erase simple_arr
                ReDim simple_arr(1 To dis_value_count)
                k = 0
                For j = 2 To UBound(arr, 1)
                    If v = arr(j, col_n) Then
                        simple_arr(k + 1) = arr(j, 1)
                        k = k + 1
                    End If
                Next j
                
                With result_sheet
                    n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                    .Cells(n, 1) = n - 1
                    .Cells(n, 2) = dis_arr2(i, 1)
                    .Cells(n, 3) = v
                    .Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                    .Cells(n, 5) = arr(1, 1)
                    .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                    .Cells(n, 7) = dis_value_count
                    .Cells(n, 8) = "average"
                    .Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.Average(simple_arr), 1)
                    
                    n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                    .Cells(n, 1) = n - 1
                    .Cells(n, 2) = dis_arr2(i, 1)
                    .Cells(n, 3) = v
                    .Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                    .Cells(n, 5) = arr(1, 1)
                    .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                    .Cells(n, 7) = dis_value_count
                    .Cells(n, 8) = "median"
                    .Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.median(simple_arr), 1)
                End With
            Next v
        
        End If

    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "there is err", ws.Range("A1")
    Call not_processed
    Exit Sub
End Sub

Sub calculate_nominal()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim result_sheet As Worksheet
    Dim xsc_sheet As Worksheet
    Dim last_col As Long
    Dim last_col_letter As String
    Dim i As Long, j As Long
    Dim last_row As Long
    Dim new_col_letter As String
    Dim last_dis As Long
    Dim dis_arr2 As Variant
    Dim n As Long
    Dim col_n As Long
    Dim unique_arr As Variant
    Dim v As Variant
    Dim k As Long
    Dim weight_arr() As Double
    Dim small_arr() As String
    Dim disagregation_arr() As String
    Dim dis_value_count As Long
    Dim data_arr() As String
    Dim unique_data_arr As Variant
    Dim data_count As Long
    Dim sum_weight As Single
    Dim sum_weight_in_var As Single
    Dim m As Long
    Dim small_unique_arr() As Variant
    Dim p As Variant
    Dim choice_count As Long
    Dim temp_arr As Variant
    Dim arr As Variant
    Dim all_options As Variant
    Dim xsc_arr As Variant
    Dim main_var As String

    Const Mkr As String = "!"
    Const Del As String = ","
            
    Set ws = sheets("keen")
    Set result_sheet = sheets("result")
    Set xsc_sheet = ThisWorkbook.sheets("xsurvey_choices")
    last_col = ws.Cells(1, Columns.count).End(xlToLeft).Column
    last_col_letter = number_to_letter(last_col, ws)
    new_col_letter = number_to_letter(last_col + 1, ws)
    last_row = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ws.Columns("C:M").NumberFormat = "@"

    unique_data_arr = extract_all_options()
    small_unique_arr = extract_all_options()

    With sheets("disaggregation_setting")
        last_dis = .Cells(Rows.count, 1).End(xlUp).Row
        dis_arr2 = .Range("A2:C" & last_dis)
    End With

    arr = ws.Range("A1").CurrentRegion

    For i = 1 To UBound(dis_arr2, 1)
        Erase weight_arr
        ReDim weight_arr(1 To UBound(arr, 1) - 1)
    
        If dis_arr2(i, 1) = "ALL" And dis_arr2(i, 2) = "yes" Then
        
            Erase data_arr
            ReDim data_arr(1 To UBound(arr, 1) - 1)
            sum_weight = 0
            For j = 2 To UBound(arr, 1)
                data_arr(j - 1) = arr(j, 1)
                sum_weight = sum_weight + arr(j, 2)
            Next j
        
            For Each v In unique_data_arr
                data_count = count_in_array(data_arr, v)
                sum_weight_in_var = 0
                For j = 2 To UBound(arr, 1)
                    If v = arr(j, 1) Then
                        sum_weight_in_var = sum_weight_in_var + arr(j, 2)
                    End If
                Next j
        
                With result_sheet
                    n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                    .Cells(n, 1) = n - 1
                    .Cells(n, 2) = "ALL"
                    .Cells(n, 3) = "ALL"
                    .Cells(n, 4) = "ALL"
                    .Cells(n, 5) = arr(1, 1)
                    .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                    .Cells(n, 7) = UBound(arr, 1) - 1
                    .Cells(n, 8) = "percentage"
                    .Cells(n, 9) = Application.WorksheetFunction.Round(sum_weight_in_var / sum_weight * 100, 1)
                    .Cells(n, 10) = data_count
                    .Cells(n, 11) = v
                    .Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(v))
                    .Cells(n, 13) = "w"
                End With
            Next v
 
        ElseIf dis_arr2(i, 1) = "ALL" And dis_arr2(i, 2) = "no" Then

            Erase data_arr
            ReDim data_arr(1 To UBound(arr, 1) - 1)
        
            For j = 2 To UBound(arr, 1)
                data_arr(j - 1) = arr(j, 1)
            Next j
        
            For Each v In unique_data_arr
                data_count = count_in_array(data_arr, v)
                With result_sheet
                    n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                    .Cells(n, 1) = n - 1
                    .Cells(n, 2) = "ALL"
                    .Cells(n, 3) = "ALL"
                    .Cells(n, 4) = "ALL"
                    .Cells(n, 5) = arr(1, 1)
                    .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                    .Cells(n, 7) = UBound(arr, 1) - 1
                    .Cells(n, 8) = "percentage"
                    .Cells(n, 9) = Application.WorksheetFunction.Round(data_count / (UBound(arr, 1) - 1) * 100, 1)
                    .Cells(n, 10) = data_count
                    .Cells(n, 11) = v
                    .Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(v))
                End With
            Next v
        
        ElseIf dis_arr2(i, 1) <> "ALL" And dis_arr2(i, 2) = "yes" Then
    
            If dis_arr2(i, 1) = arr(1, 1) Then
                '            Debug.Print "skip1: ", dis_arr2(i, 1)
                GoTo NextIteration
            End If
        
            col_n = gen_column_number(CStr(dis_arr2(i, 1)), "keen")
            Erase disagregation_arr
            ReDim disagregation_arr(1 To UBound(arr, 1) - 1)
            For j = 2 To UBound(arr, 1)
                disagregation_arr(j - 1) = arr(j, col_n)
            Next j

            unique_arr = get_unique(disagregation_arr)
        
            For Each v In unique_arr
                dis_value_count = count_in_array(disagregation_arr, v)

                Erase small_arr
                ReDim small_arr(1 To dis_value_count)
                k = 0
                sum_weight = 0
                For j = 2 To UBound(arr, 1)
                    If v = arr(j, col_n) Then
                        small_arr(k + 1) = arr(j, 1)
                        sum_weight = sum_weight + arr(j, 2)
                        k = k + 1
                    End If
                Next j
            
                temp_arr = Split(Mkr & Join(small_arr, Mkr & Del & Mkr) & Mkr, Del)
                'Count the items (Surrounded by markers) directly
            
                For Each p In small_unique_arr
                
                    sum_weight_in_var = 0
                    For j = 2 To UBound(arr, 1)
                        If p = arr(j, 1) And v = arr(j, col_n) Then
                            sum_weight_in_var = sum_weight_in_var + arr(j, 2)
                        End If
                    Next j
            
                    choice_count = UBound(Filter(temp_arr, Mkr & CStr(p) & Mkr, True, 1)) + 1
                    With result_sheet
                        n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                        .Cells(n, 1) = n - 1
                        .Cells(n, 2) = dis_arr2(i, 1)
                        .Cells(n, 3) = v
                        .Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                        .Cells(n, 5) = arr(1, 1)
                        .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                        .Cells(n, 7) = dis_value_count
                        .Cells(n, 8) = "percentage"
                        .Cells(n, 9) = Application.WorksheetFunction.Round(sum_weight_in_var / sum_weight * 100, 1)
                        .Cells(n, 10) = choice_count
                        .Cells(n, 11) = p
                        .Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(p))
                        .Cells(n, 13) = "w"
                    End With
                Next p

            Next v
    
        ElseIf dis_arr2(i, 1) <> "ALL" And dis_arr2(i, 2) = "no" Then
    
            If dis_arr2(i, 1) = arr(1, 1) Then
                '            Debug.Print "skip2: ", dis_arr2(i, 1)
                GoTo NextIteration
            End If
        
            col_n = gen_column_number(CStr(dis_arr2(i, 1)), "keen")
            Erase disagregation_arr
            ReDim disagregation_arr(1 To UBound(arr, 1) - 1)
            For j = 2 To UBound(arr, 1)
                disagregation_arr(j - 1) = arr(j, col_n)
            Next j

            unique_arr = get_unique(disagregation_arr)
        
            For Each v In unique_arr
                dis_value_count = count_in_array(disagregation_arr, v)
            
                '            Debug.Print v, dis_value_count

                Erase small_arr
                ReDim small_arr(1 To dis_value_count)
                k = 0
                For j = 2 To UBound(arr, 1)
                    If v = arr(j, col_n) Then
                        small_arr(k + 1) = arr(j, 1)
                        k = k + 1
                    End If
                Next j
            
                temp_arr = Split(Mkr & Join(small_arr, Mkr & Del & Mkr) & Mkr, Del)
                'Count the items (Surrounded by markers) directly
  
                For Each p In small_unique_arr
                    choice_count = UBound(Filter(temp_arr, Mkr & CStr(p) & Mkr, True, 1)) + 1
                    With result_sheet
                        n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                        .Cells(n, 1) = n - 1
                        .Cells(n, 2) = dis_arr2(i, 1)
                        .Cells(n, 3) = v
                        .Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                        .Cells(n, 5) = arr(1, 1)
                        .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                        .Cells(n, 7) = dis_value_count
                        .Cells(n, 8) = "percentage"
                        .Cells(n, 9) = Application.WorksheetFunction.Round(choice_count / dis_value_count * 100, 1)
                        .Cells(n, 10) = choice_count
                        .Cells(n, 11) = p
                        .Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(p))
                    End With
                Next p

            Next v
    
        End If
    
NextIteration:
    Next i

    Exit Sub
    
ErrorHandler:
    Debug.Print "there is err", ws.Range("A1")
    Call not_processed
    Exit Sub

End Sub

Sub calculate_nominal_multipe()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim result_sheet As Worksheet
    Dim last_col As Long
    Dim last_col_letter As String
    Dim i As Long, j As Long
    Dim last_row As Long
    Dim new_col_letter As String
    Dim last_dis As Long
    Dim dis_arr2 As Variant
    Dim n As Long
    Dim col_n As Long
    Dim unique_arr As Variant
    Dim v As Variant
    Dim k As Long
    Dim weight_arr() As Double
    Dim small_arr() As String
    Dim disagregation_arr() As String
    Dim dis_value_count As Long
    Dim dis_value_count2 As Long
    Dim data_arr() As String
    Dim unique_data_arr As Variant
    Dim data_count As Long
    Dim sum_weight As Single
    Dim sum_weight_in_var As Single
    Dim m As Long
    Dim small_unique_arr() As Variant
    Dim p As Variant
    Dim choice_count As Long
    Dim temp_arr As Variant
    Dim arr As Variant
    Dim arr2 As Variant
    Dim keen2_rng As Range

    Const Mkr As String = "!"
    Const Del As String = ","
            
    Set ws = sheets("keen")
    Set ws2 = sheets("keen2")
    Set result_sheet = sheets("result")

    last_col = ws.Cells(1, Columns.count).End(xlToLeft).Column
    last_col_letter = number_to_letter(last_col, ws)
    new_col_letter = number_to_letter(last_col + 1, ws)
    last_row = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ws.Columns("C:M").NumberFormat = "@"

    unique_data_arr = extract_all_options()
    small_unique_arr = extract_all_options()

    With sheets("disaggregation_setting")
        last_dis = .Cells(Rows.count, 1).End(xlUp).Row
        dis_arr2 = .Range("A2:C" & last_dis)
    End With

    arr = ws.Range("A1").CurrentRegion
    arr2 = ws2.Range("A1").CurrentRegion
    Set keen2_rng = ws2.Range("A1").CurrentRegion

    For i = 1 To UBound(dis_arr2, 1)
        Erase weight_arr
        ReDim weight_arr(1 To UBound(arr, 1) - 1)
    
        If dis_arr2(i, 1) = "ALL" And dis_arr2(i, 2) = "yes" Then
        
            Erase data_arr
            ReDim data_arr(1 To UBound(arr, 1) - 1)
            sum_weight = 0
            sum_weight = sum_weight_overall(arr2)
            For j = 2 To UBound(arr, 1)
                data_arr(j - 1) = arr(j, 1)
            
            Next j
        
            For Each v In unique_data_arr
                data_count = count_in_array(data_arr, v)
                sum_weight_in_var = 0
                For j = 2 To UBound(arr, 1)
                    If v = arr(j, 1) Then
                        sum_weight_in_var = sum_weight_in_var + arr(j, 2)
                    End If
                Next j
        
                With result_sheet
                    n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                    .Cells(n, 1) = n - 1
                    .Cells(n, 2) = "ALL"
                    .Cells(n, 3) = "ALL"
                    .Cells(n, 4) = "ALL"
                    .Cells(n, 5) = arr(1, 1)
                    .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                    .Cells(n, 7) = keen2_rng.Rows.count - 1
                    .Cells(n, 8) = "percentage"
                    .Cells(n, 9) = Application.WorksheetFunction.Round(sum_weight_in_var / sum_weight * 100, 1)
                    .Cells(n, 10) = data_count
                    .Cells(n, 11) = v
                    .Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(v))
                    .Cells(n, 13) = "w"
                End With
            Next v
 
        ElseIf dis_arr2(i, 1) = "ALL" And dis_arr2(i, 2) = "no" Then

            Erase data_arr
            ReDim data_arr(1 To UBound(arr, 1) - 1)
        
            For j = 2 To UBound(arr, 1)
                data_arr(j - 1) = arr(j, 1)
            Next j
        
            For Each v In unique_data_arr
                data_count = count_in_array(data_arr, v)
                With result_sheet
                    n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                    .Cells(n, 1) = n - 1
                    .Cells(n, 2) = "ALL"
                    .Cells(n, 3) = "ALL"
                    .Cells(n, 4) = "ALL"
                    .Cells(n, 5) = arr(1, 1)
                    .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                    .Cells(n, 7) = keen2_rng.Rows.count - 1
                    .Cells(n, 8) = "percentage"
                    .Cells(n, 9) = Application.WorksheetFunction.Round(data_count / (keen2_rng.Rows.count - 1) * 100, 1)
                    .Cells(n, 10) = data_count
                    .Cells(n, 11) = v
                    .Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(v))
                End With
            Next v
        
        ElseIf dis_arr2(i, 1) <> "ALL" And dis_arr2(i, 2) = "yes" Then
    
            If dis_arr2(i, 1) = arr(1, 1) Then
                '            Debug.Print "skip1: ", dis_arr2(i, 1)
                GoTo NextIteration
            End If
        
            col_n = gen_column_number(CStr(dis_arr2(i, 1)), "keen")
            Erase disagregation_arr
            ReDim disagregation_arr(1 To UBound(arr, 1) - 1)
            For j = 2 To UBound(arr, 1)
                disagregation_arr(j - 1) = arr(j, col_n)
            Next j

            unique_arr = get_unique(disagregation_arr)
        
            For Each v In unique_arr
             
                dis_value_count = count_in_array(disagregation_arr, v)
                Erase small_arr
                ReDim small_arr(1 To dis_value_count)
                k = 0
            
                sum_weight = 0
            
                For j = 2 To UBound(arr, 1)
                    If v = arr(j, col_n) Then
                        small_arr(k + 1) = arr(j, 1)
                        k = k + 1
                    End If
                Next j
            
                temp_arr = Split(Mkr & Join(small_arr, Mkr & Del & Mkr) & Mkr, Del)
                'Count the items (Surrounded by markers) directly
                sum_weight = sum_weight_when(arr2, CStr(v), col_n)
            
                dis_value_count2 = Application.WorksheetFunction.CountIf(keen2_rng.Columns(col_n), v)
                '            Debug.Print dis_value_count2
                For Each p In small_unique_arr
                
                    sum_weight_in_var = 0
                    For j = 2 To UBound(arr, 1)
                        If p = arr(j, 1) And v = arr(j, col_n) Then
                            sum_weight_in_var = sum_weight_in_var + arr(j, 2)
                        End If
                    Next j
            
                    choice_count = UBound(Filter(temp_arr, Mkr & CStr(p) & Mkr, True, 1)) + 1
                    With result_sheet
                        n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                        .Cells(n, 1) = n - 1
                        .Cells(n, 2) = dis_arr2(i, 1)
                        .Cells(n, 3) = v
                        .Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                        .Cells(n, 5) = arr(1, 1)
                        .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                        .Cells(n, 7) = dis_value_count2
                        .Cells(n, 8) = "percentage"
                        .Cells(n, 9) = Application.WorksheetFunction.Round(sum_weight_in_var / sum_weight * 100, 1)
                        .Cells(n, 10) = choice_count
                        .Cells(n, 11) = p
                        .Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(p))
                        .Cells(n, 13) = "w"
                    End With
                Next p

            Next v
    
        ElseIf dis_arr2(i, 1) <> "ALL" And dis_arr2(i, 2) = "no" Then
    
            If dis_arr2(i, 1) = arr(1, 1) Then
                '            Debug.Print "skip2: ", dis_arr2(i, 1)
                GoTo NextIteration
            End If
        
            col_n = gen_column_number(CStr(dis_arr2(i, 1)), "keen")
            Erase disagregation_arr
            ReDim disagregation_arr(1 To UBound(arr, 1) - 1)
            For j = 2 To UBound(arr, 1)
                disagregation_arr(j - 1) = arr(j, col_n)
            Next j

            unique_arr = get_unique(disagregation_arr)
        
            For Each v In unique_arr
                dis_value_count = count_in_array(disagregation_arr, v)
                dis_value_count2 = Application.WorksheetFunction.CountIf(keen2_rng.Columns(col_n), v)
                '            Debug.Print v, dis_value_count

                Erase small_arr
                ReDim small_arr(1 To dis_value_count)
                k = 0
                For j = 2 To UBound(arr, 1)
                    If v = arr(j, col_n) Then
                        small_arr(k + 1) = arr(j, 1)
                        k = k + 1
                    End If
                Next j
            
                temp_arr = Split(Mkr & Join(small_arr, Mkr & Del & Mkr) & Mkr, Del)
                'Count the items (Surrounded by markers) directly
  
                For Each p In small_unique_arr
                    choice_count = UBound(Filter(temp_arr, Mkr & CStr(p) & Mkr, True, 1)) + 1
                    With result_sheet
                        n = .Cells(Rows.count, 1).End(xlUp).Row + 1
                        .Cells(n, 1) = n - 1
                        .Cells(n, 2) = dis_arr2(i, 1)
                        .Cells(n, 3) = v
                        .Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                        .Cells(n, 5) = arr(1, 1)
                        .Cells(n, 6) = var_label(CStr(arr(1, 1)))
                        .Cells(n, 7) = dis_value_count2
                        .Cells(n, 8) = "percentage"
                        .Cells(n, 9) = Application.WorksheetFunction.Round(choice_count / dis_value_count2 * 100, 1)
                        .Cells(n, 10) = choice_count
                        .Cells(n, 11) = p
                        .Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(p))
                    End With
                Next p

            Next v
    
        End If
    
NextIteration:
    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "there is err", ws.Range("A1")
    Call not_processed
    Exit Sub

End Sub

Function get_unique(arr As Variant) As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        dict(arr(i)) = 1
    Next i
    
    get_unique = dict.Keys()
End Function

Function count_in_array(arr As Variant, item As Variant) As Long
    Dim i As Long, count As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = item Then
            count = count + 1
        End If
    Next i
    count_in_array = count
End Function

Sub count_unique()
    Dim arr() As Variant
    Dim dict As Object
    Dim i As Long, v As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = LBound(arr) To UBound(arr)
        If dict.Exists(arr(i)) Then
            dict.item(arr(i)) = dict.item(arr(i)) + 1
        Else
            dict.Add arr(i), 1
        End If
    Next i
    
End Sub

Function extract_all_options() As Variant
    On Error GoTo 0
    Dim xsc_arr As Variant
    Dim data_arr As Variant
    Dim dict As Object
    Dim i As Long
    Dim main_var As String
    Dim ws As Worksheet
    Dim count As Long
    Dim v As Variant
    Dim result() As Variant
    
    Set ws = sheets("keen")
    Set dict = CreateObject("Scripting.Dictionary")
    xsc_arr = ThisWorkbook.sheets("xsurvey_choices").Range("A1").CurrentRegion
    main_var = ws.Range("A1")
    
    For i = 1 To UBound(xsc_arr, 1)
        If xsc_arr(i, 2) = main_var Then
            dict.Add xsc_arr(i, 4), i
        End If
    Next i
    
    If dict.count > 1 Then
        extract_all_options = dict.Keys()
    Else
        data_arr = ws.Range(ws.Range("A2"), ws.Range("A2").End(xlDown))
        For i = LBound(data_arr, 1) To UBound(data_arr, 1)
            If Len(data_arr(i, 1)) > 0 Then
                dict(data_arr(i, 1)) = i
            End If
            
        Next i
        
        ' Count non-empty keys
        count = 0
        For Each v In dict.Keys()
            If Len(v) > 0 Then
                count = count + 1
            End If
        Next v
        
        ReDim result(1 To count)
        
        ' Assign non-empty keys to result array
        count = 0
        For Each v In dict.Keys()
            If Len(v) > 0 Then
                count = count + 1
                result(count) = v
            End If
        Next v
        
        extract_all_options = result
    
    End If
    
End Function

Sub unify_data()
    Dim ws As Worksheet
    Dim last_col As Long
    Dim last_col_letter As String
    Dim i As Long, j As Long
    Dim arr() As String
    Dim last_row As Long
    Dim end_row As Long
    Dim k As Long
    Dim last_dis As Long
    Dim dis_arr As Variant
    Dim data_arr As Variant
    Dim total_sum As Single
    Dim total_count As Long
    Dim unique_data_arr As Variant
    Dim v As Variant
    Dim keen2_ws As Worksheet
    
    Set ws = sheets("keen")
    ws.Columns("O:AZ").Clear
    ws.Columns("A:AZ").NumberFormat = "@"
    last_col = ws.Cells(1, Columns.count).End(xlToLeft).Column
    last_col_letter = number_to_letter(last_col, ws)
    last_row = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    For i = 1 To last_row
        arr = Split(ws.Cells(i, 1), " ")
        end_row = end_row + (UBound(arr) - LBound(arr) + 1)
    Next i
    
    For i = last_row To 1 Step -1
        arr = Split(ws.Cells(i, 1), " ")
        For j = LBound(arr) To UBound(arr)
            ws.Cells(end_row, 15) = arr(j)
            If last_col > 1 Then
                For k = 1 To last_col
                    ws.Cells(end_row, 15 + k) = ws.Cells(i, k + 1)
                Next k
            End If
            end_row = end_row - 1
        Next j
    Next i
    
    If Not worksheet_exists("keen2") Then
        Call create_sheet(find_main_data, "keen2")
        sheets("keen2").Visible = xlVeryHidden
    End If
    
    Set keen2_ws = sheets("keen2")
    keen2_ws.Cells.Clear
    
    If WITH_WEIGHT Then
        keen2_ws.Columns("A:A").NumberFormat = "@"
        keen2_ws.Columns("B:B").NumberFormat = "0.000"
        keen2_ws.Columns("C:Z").NumberFormat = "@"
    Else
        keen2_ws.Columns("A:Z").NumberFormat = "@"
    End If
    
    ws.Range("A1").CurrentRegion.Copy
    keen2_ws.Range("A1").PasteSpecial xlPasteValues

    ws.Columns("A:N").Delete
End Sub

Function sum_weight_overall(arr As Variant) As Double
    Dim i As Long
    Dim sum As Double
    sum = 0
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        sum = sum + arr(i, 2)
    Next i
    sum_weight_overall = sum
End Function

Function sum_weight_when(arr As Variant, criteria As String, col_index As Long) As Double
    Dim i As Long
    Dim sum As Double
    sum = 0
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        If arr(i, col_index) = criteria Then
            sum = sum + arr(i, 2)
        End If
    Next i
    sum_weight_when = sum
End Function

Sub check_result_sheet(SHEET_NAME As String)
    Dim wb As Workbook
    Dim resultSheet As Worksheet
    Dim i As Integer
    Dim column_widths As Variant
    
    Set wb = ActiveWorkbook
    
    If Not worksheet_exists("result") Then
        Call create_sheet(SHEET_NAME, "result")
        Set resultSheet = wb.sheets("result")
        With resultSheet
            .Cells(1, 1).Resize(1, 15).value = Array("row", "disaggregation", "disaggregation value", "disaggregation label", _
                                                      "variable", "variable label", "valid numbers", "measurement type", _
                                                      "measurement value", "count", "choice", "choice label", "weight", _
                                                      "hkey", "hkey order")
            
            ' column widths
            column_widths = Array(6, 15, 18, 25, 15, 45, 15, 15, 20, 10, 15, 45, 7, 45, 15)
            For i = 1 To 15
                .Columns(i).ColumnWidth = column_widths(i - 1)
            Next i
            
            .Columns("B:F").NumberFormat = "@"
            .Columns("K:M").NumberFormat = "@"
            .Visible = False
        End With
    End If

End Sub

' check if main data sheet has weight column or not
Private Function has_weight() As Boolean
    Dim main_ws As Worksheet
    Dim last_main_col_letter As String
    Dim cel As Variant
    
    Set main_ws = sheets(find_main_data)
    
    last_main_col_letter = Split(main_ws.Cells.Find(What:="*", after:=[a1], SearchOrder:=xlByColumns, _
                                                    SearchDirection:=xlPrevious).Cells.Address(1, 0), "$")(0)
    
    For Each cel In main_ws.Range("A1:" & last_main_col_letter & 1)
        If cel = "weight" Then
            has_weight = True
            Exit For
        Else
            has_weight = False
        End If
    Next
    
End Function

Sub make_header_order()

    Dim rng As Range
    Application.ScreenUpdating = False
    Dim res_ws As Worksheet
    Dim ws As Worksheet
    Dim last_result As Long
    Dim last_header As Long
    
    Set res_ws = sheets("result")
    last_result = res_ws.Cells(res_ws.Rows.count, "A").End(xlUp).Row
    
    res_ws.Activate
    res_ws.Range("N2:N" & last_result).Formula = "=E2&K2&H2"
    res_ws.Columns("N:N").Copy
    res_ws.Columns("N:N").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
    Set ws = sheets("temp_sheet")
    ws.Cells.Clear
    
    ws.Range("A1:A" & last_result).Value2 = res_ws.Range("N1:N" & last_result).Value2
    ws.Range("A1").CurrentRegion.RemoveDuplicates Columns:=1, header:=xlYes
    
    last_header = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
     
    ws.Range("B2") = "1"
    ws.Range("B2").AutoFill Destination:=ws.Range("B2:B" & last_header), Type:=xlFillSeries
    
    ws.Columns("B:B").Copy
    ws.Columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
         
    res_ws.Range("O2:O" & last_result).Formula = "=VLOOKUP(result!N2,temp_sheet!A:B,2,)"
    
    res_ws.Columns("O:O").Copy
    res_ws.Columns("O:O").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
End Sub

Sub delete_un_selected_choices()
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim dis_ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim var_counter As Long
    Dim rng As Range
    Dim var_rng As Range
    Dim choice_rng As Range
    Dim cell As Range
    Dim var_collection As New Collection
    Dim v As Variant
    Dim var_list As String
    
    Set ws = sheets("result")
    Set dis_ws = sheets("disaggregation_setting")
    Set rng = ws.Range("A1").CurrentRegion
    
    dis_ws.Columns("K:Q").Clear
    
    dis_ws.Range("K1") = "measurement type"
    dis_ws.Range("L1") = "measurement Value"
    dis_ws.Range("K2") = "percentage"
    dis_ws.Range("L2") = 0
    dis_ws.Range("P1") = "variable"
    
    rng.AdvancedFilter xlFilterCopy, dis_ws.Range("K1:L2"), dis_ws.Range("P1").CurrentRegion
    
    Set var_rng = dis_ws.Range("P1").CurrentRegion
    Set choice_rng = ThisWorkbook.sheets("xsurvey_choices").Range("B:B")
    
    If IsEmpty(dis_ws.Range("P2")) Then
        Debug.Print "exit function"
        Exit Sub
    End If
    
    var_rng.RemoveDuplicates Columns:=1, header:=xlYes
    
    Set var_rng = dis_ws.Range("P1").CurrentRegion
    
    var_counter = 0
    
    For Each cell In var_rng
        dis_ws.Cells(cell.Row, "Q") = Application.WorksheetFunction.CountIf(choice_rng, cell)
        If dis_ws.Cells(cell.Row, "Q") > 10 Then
            var_counter = var_counter + 1
        End If
    Next cell
    
    Debug.Print "var_counter: " & var_counter
    
    If var_counter = 0 Then
        Debug.Print "exit function"
        Exit Sub
    End If
    
    Dim var_arr() As String
    Dim j As Integer
    j = 1

    ReDim var_arr(1 To var_counter)
    
    For Each cell In var_rng
        If dis_ws.Cells(cell.Row, "Q") > 10 Then
            var_arr(j) = cell
            j = j + 1
        End If
    Next cell
    
    Call delete_zero_values(var_arr)
    dis_ws.Columns("K:Q").Clear
    
End Sub

Sub delete_zero_values(cr() As String)

    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = sheets("result")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    ws.AutoFilterMode = False
    
    With ws.Range("A1").CurrentRegion
        .AutoFilter Field:=9, Criteria1:="0"
        .AutoFilter Field:=5, Criteria1:=cr, Operator:=xlFilterValues
    End With
    
    Dim rngFiltered As Range
    Dim rngToDelete As Range
    Dim firstRow As Range
   
    If Not ws.AutoFilterMode Then
        Debug.Print "No filter applied."
        Exit Sub
    End If
    
    On Error Resume Next
    Set rngFiltered = ws.AutoFilter.Range.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If rngFiltered Is Nothing Then
        Debug.Print "No visible rows after filtering."
        Exit Sub
    End If
    
    Set firstRow = ws.Rows(1)
    Set rngToDelete = Intersect(ws.UsedRange.Offset(1), rngFiltered)
    
    If Not rngToDelete Is Nothing Then
        rngToDelete.EntireRow.Delete
    Else
        Debug.Print "No rows to delete."
    End If
   
    ws.AutoFilterMode = False
End Sub

Private Sub not_processed()
    Dim str As String
    Dim lines() As String
    Dim lastLine As String
    lines = Split(analysis_form.TextInfo, vbCrLf)
    lastLine = lines(LBound(lines))
    analysis_form.TextInfo.value = Replace(analysis_form.TextInfo, lastLine, lastLine & " !")
End Sub

Private Sub gender_disaggregation()

    Dim ws As Worksheet
    Dim g_ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim colNums(1 To 8) As Long
    Dim sumCol(1 To 8) As Long
    Dim totalRange As Range
    Dim lastRowMain As Long
    Dim totalFemale As Double
    Dim totalMale As Double
    Dim totalSum As Double
    Dim female As Double
    Dim male As Double
    
    Set ws = sheets(find_main_data)
    lastRowMain = ws.Cells(ws.Rows.count, find_uuid_coln).End(xlUp).Row
    
    If Not worksheet_exists("gender_summary") Then
        Call create_sheet("analysis_list", "gender_summary")
    End If
            
    Set g_ws = sheets("gender_summary")
    g_ws.Cells.Clear

    colNums(1) = gen_column_number("new_born_female", find_main_data)
    colNums(2) = gen_column_number("new_born_male", find_main_data)
    colNums(3) = gen_column_number("girls_6_17", find_main_data)
    colNums(4) = gen_column_number("boys_6_17", find_main_data)
    colNums(5) = gen_column_number("adult_18_59_female", find_main_data)
    colNums(6) = gen_column_number("adult_18_59_male", find_main_data)
    colNums(7) = gen_column_number("elders_60_abv_female", find_main_data)
    colNums(8) = gen_column_number("elders_60_abv_male", find_main_data)
    
    Debug.Print "WITH_WEIGHT: " & WITH_WEIGHT
    
    If WITH_WEIGHT Then
        Dim values As Variant
        Dim wcol As Long
        wcol = gen_column_number("weight", find_main_data)
        values = ws.Range(ws.Cells(2, colNums(1)), ws.Cells(lastRowMain, colNums(8))).value
        For i = 1 To UBound(values, 1)
            For j = 1 To UBound(values, 2)
                values(i, j) = values(i, j) * ws.Cells(i + 1, wcol).value
            Next j
        Next i
        g_ws.Range("A2").Resize(UBound(values, 1), UBound(values, 2)).value = values
    Else
        For i = 1 To 8
            ws.Columns(colNums(i)).Copy g_ws.Columns(i)
        Next i
    End If
    
    Set totalRange = g_ws.Range("A1:H" & lastRowMain)
    
    sumCol(1) = Application.WorksheetFunction.sum(g_ws.Range("A1:A" & lastRowMain))
    sumCol(2) = Application.WorksheetFunction.sum(g_ws.Range("B1:B" & lastRowMain))
    sumCol(3) = Application.WorksheetFunction.sum(g_ws.Range("C1:C" & lastRowMain))
    sumCol(4) = Application.WorksheetFunction.sum(g_ws.Range("D1:D" & lastRowMain))
    sumCol(5) = Application.WorksheetFunction.sum(g_ws.Range("E1:E" & lastRowMain))
    sumCol(6) = Application.WorksheetFunction.sum(g_ws.Range("F1:F" & lastRowMain))
    sumCol(7) = Application.WorksheetFunction.sum(g_ws.Range("G1:G" & lastRowMain))
    sumCol(8) = Application.WorksheetFunction.sum(g_ws.Range("H1:H" & lastRowMain))
    
    totalFemale = sumCol(1) + sumCol(3) + sumCol(5) + sumCol(7)
    totalMale = sumCol(2) + sumCol(4) + sumCol(6) + sumCol(8)
    totalSum = totalFemale + totalMale
    
    female = Application.WorksheetFunction.Round(100 * totalFemale / totalSum, 1)
    male = Application.WorksheetFunction.Round(100 * totalMale / totalSum, 1)
    
    g_ws.Range("K1:L1").value = Array("Male", CStr(male) & "%")
    g_ws.Range("N1:O1").value = Array("Female", CStr(female) & "%")
    g_ws.Range("K3:K6").value = Application.Transpose(Array("0y-5y", "6y-17y", "18y-59y", "60y+"))
    g_ws.Range("N3:N6").value = Application.Transpose(Array("0y-5y", "6y-17y", "18y-59y", "60y+"))
    
    g_ws.Range("L3:L6").value = Application.Transpose(Array( _
                                        CStr(Application.WorksheetFunction.Round(100 * sumCol(2) / totalSum, 1)) & "%", _
                                        CStr(Application.WorksheetFunction.Round(100 * sumCol(4) / totalSum, 1)) & "%", _
                                        CStr(Application.WorksheetFunction.Round(100 * sumCol(6) / totalSum, 1)) & "%", _
                                        CStr(Application.WorksheetFunction.Round(100 * sumCol(8) / totalSum, 1)) & "%"))
                                        
    g_ws.Range("O3:O6").value = Application.Transpose(Array( _
                                        CStr(Application.WorksheetFunction.Round(100 * sumCol(1) / totalSum, 1)) & "%", _
                                        CStr(Application.WorksheetFunction.Round(100 * sumCol(3) / totalSum, 1)) & "%", _
                                        CStr(Application.WorksheetFunction.Round(100 * sumCol(5) / totalSum, 1)) & "%", _
                                        CStr(Application.WorksheetFunction.Round(100 * sumCol(7) / totalSum, 1)) & "%"))
    
    Call format_gender_summary
'    g_ws.Range("F1").value = "WEIGHT: " & WITH_WEIGHT
End Sub

Sub format_gender_summary()
    On Error Resume Next
    Dim ws As Worksheet
    Dim rng1 As Range
    Dim rng2 As Range
    
    Set ws = sheets("gender_summary")
    
    ws.Columns("A:J").Delete Shift:=xlToLeft
    
    ws.Columns("B:B").ColumnWidth = 20
    ws.Columns("C:C").ColumnWidth = 5
    ws.Columns("E:E").ColumnWidth = 20
    
    With ws.Range("A1:E1").Font
        .Bold = True
    End With
    
    Set rng1 = ws.Range("B3:B6")
    Set rng2 = ws.Range("E3:E6")
    
    rng1.FormatConditions.AddDatabar
    rng1.FormatConditions(rng1.FormatConditions.count).ShowValue = True
    rng1.FormatConditions(rng1.FormatConditions.count).SetFirstPriority
    With rng1.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With rng1.FormatConditions(1).BarColor
        .Color = 2668287
        .TintAndShade = 0
    End With
    rng1.FormatConditions(1).BarFillType = xlDataBarFillSolid
    rng1.FormatConditions(1).Direction = xlContext
    rng1.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    With rng1.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With rng1.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
    
    rng2.FormatConditions.AddDatabar
    rng2.FormatConditions(rng2.FormatConditions.count).ShowValue = True
    rng2.FormatConditions(rng2.FormatConditions.count).SetFirstPriority
    With rng2.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
        .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
    End With
    With rng2.FormatConditions(1).BarColor
        .Color = 2668287
        .TintAndShade = 0
    End With
    rng2.FormatConditions(1).BarFillType = xlDataBarFillSolid
    rng2.FormatConditions(1).Direction = xlContext
    rng2.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    With rng2.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With rng2.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
    
    With ws.Range("A3:E3")
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    With ws.Range("A6:E6")
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
    
    ws.Columns("C:C").Clear
    
    ws.Rows("1:2").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Range("A1:E1").Merge
    If WITH_WEIGHT Then
        ws.Range("A1").value = "Gender Demographics (weighted)"
    Else
        ws.Range("A1").value = "Gender Demographics"
    End If
    
    
    With ws.Range("A1").Font
        .Bold = True
        .Size = 18
        .Color = RGB(150, 150, 150)
    End With
    
    With ws.Range("A3:E8").Font
        .Size = 10
    End With
    
    With ws.Range("B3:B8")
        .Style = "Percent"
        .NumberFormat = "0.0%"
    End With
    
    With ws.Range("E3:E8")
        .Style = "Percent"
        .NumberFormat = "0.0%"
    End With
    
    On Error GoTo 0
End Sub

Private Function check_gender_column() As Boolean
    Dim ws As Worksheet
    Dim header As Range
    Dim columnHeaders As Variant
    Dim columnHeader As Variant
    Dim missingHeaders As String

    ' these colums are for gender sub catagories:
    columnHeaders = Array("new_born_female", "new_born_male", _
                          "girls_6_17", "boys_6_17", _
                          "adult_18_59_female", "adult_18_59_male", _
                          "elders_60_abv_female", "elders_60_abv_male")
    
    Set ws = sheets(find_main_data)
    
    For Each columnHeader In columnHeaders
        Set header = ws.Rows(1).Find(columnHeader, LookIn:=xlValues, LookAt:=xlWhole)
        If header Is Nothing Then
            missingHeaders = missingHeaders & columnHeader & vbCrLf
        End If
    Next columnHeader

    If missingHeaders <> "" Then
        Debug.Print " missing:" & vbCrLf & missingHeaders
        check_gender_column = False
    Else
        check_gender_column = True
    End If
    
End Function



