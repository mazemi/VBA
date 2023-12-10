Attribute VB_Name = "Analyzer_module"
Option Explicit
Global WITH_ALL As Boolean
Global WITH_WEIGHT As Boolean
Global SELECT_MULTIPLE_COLL As Collection

Sub show_keen()
    sheets("keen").Visible = True
End Sub

Sub do_analize()
Dim t1 As Double
t1 = Timer

Application.ScreenUpdating = False
Application.DisplayAlerts = False

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

WITH_ALL = False
WITH_WEIGHT = False

analysis_form.TextInfo.Value = "Starting ... "
Debug.Print "new analysis0: " & Timer - t1
Application.Wait (Now + 0.00001)

Call remove_NA

DoEvents
str_info = vbLf & analysis_form.TextInfo.Value
analysis_form.TextInfo.Value = "Removed NAs" & str_info
str_info = vbLf & analysis_form.TextInfo.Value
    
Set main_ws = sheets(find_main_data)
Call remove_empty_col

With sheets("dissagregation_setting")
    last_dis = .Cells(rows.count, 1).End(xlUp).Row
    dis_arr = .Range("A2:B" & last_dis)
End With

last_row_main_data = main_ws.Cells(rows.count, uuid_coln).End(xlUp).Row
data_rows = CStr(2) & ":" & CStr(last_row_main_data)
last_col = main_ws.Cells(1, columns.count).End(xlToLeft).Column
last_col_letter = number_to_letter(last_col, main_ws)

For i = 1 To UBound(dis_arr, 1)
    If dis_arr(i, 1) <> "ALL" Then
        WITH_ALL = True
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
    last_var = .Cells(rows.count, 1).End(xlUp).Row
    var_arr = .Range("A2:B" & last_var)
End With

Debug.Print "all and weight: " & WITH_ALL, WITH_WEIGHT

If worksheet_exists("result") Then
    sheets("result").Delete
End If

Call check_result_sheet(main_ws.name)
Set result_sheet = wb.sheets("result")

If Not worksheet_exists("keen") Then
    Call create_sheet(main_ws.name, "keen")
        sheets("keen").Visible = xlVeryHidden
End If

Set keen_ws = sheets("keen")
keen_ws.Cells.Clear

If WITH_WEIGHT Then
    keen_ws.columns("C:M").NumberFormat = "@"
Else
    keen_ws.columns("B:M").NumberFormat = "@"
End If

If Not worksheet_exists("temp_sheet") Then
    Call create_sheet(main_ws.name, "temp_sheet")
     sheets("temp_sheet").Visible = xlVeryHidden
End If

sheets("temp_sheet").Cells.Clear
sheets("temp_sheet").rows(1).Value = main_ws.rows(1).Value

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
'      Debug.Print "new analysis: " & Timer - t1, var_arr(i, 1)

     ' show progress on the analysis user form
     If Len(str_info) > 2000 Then
         analysis_form.TextInfo.Value = left(analysis_form.TextInfo.Value, 1000)
     End If

     str_info = vbLf & analysis_form.TextInfo.Value
 
     txt = "Proccessing: " & CStr(var_arr(i, 1)) & str_info
     txt = Replace(txt, "0", "")
     analysis_form.TextInfo.Value = txt
            
    filtered_arr = Filter(header_arr, var_arr(i, 1), True, vbTextCompare)
    If UBound(filtered_arr) = -1 Then
'        Debug.Print "skipped: (not exist in the data)" & var_arr(i, 1)
        GoTo NextIteration
    End If
    
    sheets("temp_sheet").rows(2).Clear
    var_col_letter = gen_column_letter(CStr(var_arr(i, 1)), "temp_sheet")
    
    If var_col_letter = "" Then
        Debug.Print "column not exist in the main data."
        GoTo NextIteration
    End If
    
    sheets("temp_sheet").Range(var_col_letter & 2) = "<>"
    
    keen_ws.Range("A1") = var_arr(i, 1)
    keen_ws.rows(data_rows).Clear
    
    On Error GoTo criticalErrHandler
    main_rng.AdvancedFilter xlFilterCopy, cr_rng, keen_ws.Range("A1").CurrentRegion
    On Error GoTo 0
    
    If IsEmpty(keen_ws.Range("A2")) Then
'        Debug.Print "skipped: (no data)" & var_arr(i, 1)
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

last_row_result = result_sheet.Cells(rows.count, 1).End(xlUp).Row

If last_row_result > 2 Then
    Call make_header_order
End If

Debug.Print "new analysis: " & Timer - t1
wb.Save
Exit Sub

criticalErrHandler:
Application.ScreenUpdating = True
Application.DisplayAlerts = True

Debug.Print "critical error"
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

sheets("dissagregation_setting").Visible = True
Set ws = sheets("keen")
Set result_sheet = sheets("result")
last_col = ws.Cells(1, columns.count).End(xlToLeft).Column
last_col_letter = number_to_letter(last_col, ws)
new_col_letter = number_to_letter(last_col + 1, ws)
last_row = ws.Cells(ws.rows.count, "A").End(xlUp).Row

ws.columns("C:M").NumberFormat = "@"

With sheets("dissagregation_setting")
    last_dis = .Cells(rows.count, 1).End(xlUp).Row
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

        n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
        result_sheet.Cells(n, 1) = n - 1
        result_sheet.Cells(n, 2) = "ALL"
        result_sheet.Cells(n, 3) = "ALL"
        result_sheet.Cells(n, 4) = "ALL"
        result_sheet.Cells(n, 5) = arr(1, 1)
        result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
        result_sheet.Cells(n, 7) = UBound(arr, 1) - 1
        result_sheet.Cells(n, 8) = "average"
        result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.sum(simple_arr) / WorksheetFunction.sum(weight_arr), 1)
        result_sheet.Cells(n, 13) = "w"
        
        n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
        result_sheet.Cells(n, 1) = n - 1
        result_sheet.Cells(n, 2) = "ALL"
        result_sheet.Cells(n, 3) = "ALL"
        result_sheet.Cells(n, 4) = "ALL"
        result_sheet.Cells(n, 5) = arr(1, 1)
        result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
        result_sheet.Cells(n, 7) = UBound(arr, 1) - 1
        result_sheet.Cells(n, 8) = "median"
        result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.median(simple_arr), 1)
        result_sheet.Cells(n, 13) = "w"
        
    ElseIf dis_arr2(i, 1) = "ALL" And dis_arr2(i, 2) = "no" Then
        For j = 2 To UBound(arr, 1)
            simple_arr(j - 1) = arr(j, 1)
        Next j
        
        n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
        result_sheet.Cells(n, 1) = n - 1
        result_sheet.Cells(n, 2) = "ALL"
        result_sheet.Cells(n, 3) = "ALL"
        result_sheet.Cells(n, 4) = "ALL"
        result_sheet.Cells(n, 5) = arr(1, 1)
        result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
        result_sheet.Cells(n, 7) = UBound(arr, 1) - 1
        result_sheet.Cells(n, 8) = "average"
        result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.Average(simple_arr), 1)

        n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
        result_sheet.Cells(n, 1) = n - 1
        result_sheet.Cells(n, 2) = "ALL"
        result_sheet.Cells(n, 3) = "ALL"
        result_sheet.Cells(n, 4) = "ALL"
        result_sheet.Cells(n, 5) = arr(1, 1)
        result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
        result_sheet.Cells(n, 7) = UBound(arr, 1) - 1
        result_sheet.Cells(n, 8) = "median"
        result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.median(simple_arr), 1)
        
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
            
            n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
            result_sheet.Cells(n, 1) = n - 1
            result_sheet.Cells(n, 2) = dis_arr2(i, 1)
            result_sheet.Cells(n, 3) = v
            result_sheet.Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
            result_sheet.Cells(n, 5) = arr(1, 1)
            result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
            result_sheet.Cells(n, 7) = dis_value_count
            result_sheet.Cells(n, 8) = "average"
            result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.sum(simple_arr) / WorksheetFunction.sum(weight_arr), 1)
            result_sheet.Cells(n, 13) = "w"
            
            n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
            result_sheet.Cells(n, 1) = n - 1
            result_sheet.Cells(n, 2) = dis_arr2(i, 1)
            result_sheet.Cells(n, 3) = v
            result_sheet.Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
            result_sheet.Cells(n, 5) = arr(1, 1)
            result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
            result_sheet.Cells(n, 7) = dis_value_count
            result_sheet.Cells(n, 8) = "median"
            result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.median(simple_arr), 1)
            result_sheet.Cells(n, 13) = "w"
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
            
            n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
            result_sheet.Cells(n, 1) = n - 1
            result_sheet.Cells(n, 2) = dis_arr2(i, 1)
            result_sheet.Cells(n, 3) = v
            result_sheet.Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
            result_sheet.Cells(n, 5) = arr(1, 1)
            result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
            result_sheet.Cells(n, 7) = dis_value_count
            result_sheet.Cells(n, 8) = "average"
            result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.Average(simple_arr), 1)
                
            n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
            result_sheet.Cells(n, 1) = n - 1
            result_sheet.Cells(n, 2) = dis_arr2(i, 1)
            result_sheet.Cells(n, 3) = v
            result_sheet.Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
            result_sheet.Cells(n, 5) = arr(1, 1)
            result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
            result_sheet.Cells(n, 7) = dis_value_count
            result_sheet.Cells(n, 8) = "median"
            result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(WorksheetFunction.median(simple_arr), 1)
            
        Next v
        
    End If

Next i

Exit Sub

ErrorHandler:
Debug.Print "there is err", ws.Range("A1")
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
            
sheets("dissagregation_setting").Visible = True
Set ws = sheets("keen")
Set result_sheet = sheets("result")
Set xsc_sheet = ThisWorkbook.sheets("xsurvey_choices")
last_col = ws.Cells(1, columns.count).End(xlToLeft).Column
last_col_letter = number_to_letter(last_col, ws)
new_col_letter = number_to_letter(last_col + 1, ws)
last_row = ws.Cells(ws.rows.count, "A").End(xlUp).Row

ws.columns("C:M").NumberFormat = "@"

unique_data_arr = extract_all_options()
small_unique_arr = extract_all_options()

With sheets("dissagregation_setting")
    last_dis = .Cells(rows.count, 1).End(xlUp).Row
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
        
'            Debug.Print v, data_count, sum_weight, sum_weight_in_var, UBound(arr, 1) - 1
            n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
            result_sheet.Cells(n, 1) = n - 1
            result_sheet.Cells(n, 2) = "ALL"
            result_sheet.Cells(n, 3) = "ALL"
            result_sheet.Cells(n, 4) = "ALL"
            result_sheet.Cells(n, 5) = arr(1, 1)
            result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
            result_sheet.Cells(n, 7) = UBound(arr, 1) - 1
            result_sheet.Cells(n, 8) = "percentage"
            result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(sum_weight_in_var / sum_weight * 100, 1)
            result_sheet.Cells(n, 10) = data_count
            result_sheet.Cells(n, 11) = v
            result_sheet.Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(v))
            result_sheet.Cells(n, 13) = "w"

        Next v
 
    ElseIf dis_arr2(i, 1) = "ALL" And dis_arr2(i, 2) = "no" Then

        Erase data_arr
        ReDim data_arr(1 To UBound(arr, 1) - 1)
        
        For j = 2 To UBound(arr, 1)
            data_arr(j - 1) = arr(j, 1)
        Next j
        
        For Each v In unique_data_arr
            data_count = count_in_array(data_arr, v)
'            Debug.Print v, data_count, UBound(arr, 1) - 1
            n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
            result_sheet.Cells(n, 1) = n - 1
            result_sheet.Cells(n, 2) = "ALL"
            result_sheet.Cells(n, 3) = "ALL"
            result_sheet.Cells(n, 4) = "ALL"
            result_sheet.Cells(n, 5) = arr(1, 1)
            result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
            result_sheet.Cells(n, 7) = UBound(arr, 1) - 1
            result_sheet.Cells(n, 8) = "percentage"
            result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(data_count / (UBound(arr, 1) - 1) * 100, 1)
            result_sheet.Cells(n, 10) = data_count
            result_sheet.Cells(n, 11) = v
            result_sheet.Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(v))
     
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
'            Debug.Print v, dis_value_count, sum_weight
            
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
'                Debug.Print p, choice_count, " ", sum_weight_in_var
                n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
                result_sheet.Cells(n, 1) = n - 1
                result_sheet.Cells(n, 2) = dis_arr2(i, 1)
                result_sheet.Cells(n, 3) = v
                result_sheet.Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                result_sheet.Cells(n, 5) = arr(1, 1)
                result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
                result_sheet.Cells(n, 7) = dis_value_count
                result_sheet.Cells(n, 8) = "percentage"
                result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(sum_weight_in_var / sum_weight * 100, 1)
                result_sheet.Cells(n, 10) = choice_count
                result_sheet.Cells(n, 11) = p
                result_sheet.Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(p))
                result_sheet.Cells(n, 13) = "w"
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
'                Debug.Print p, choice_count
                n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
                result_sheet.Cells(n, 1) = n - 1
                result_sheet.Cells(n, 2) = dis_arr2(i, 1)
                result_sheet.Cells(n, 3) = v
                result_sheet.Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                result_sheet.Cells(n, 5) = arr(1, 1)
                result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
                result_sheet.Cells(n, 7) = dis_value_count
                result_sheet.Cells(n, 8) = "percentage"
                result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(choice_count / dis_value_count * 100, 1)
                result_sheet.Cells(n, 10) = choice_count
                result_sheet.Cells(n, 11) = p
                result_sheet.Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(p))
            Next p

        Next v
    
    End If
    
NextIteration:
Next i

Exit Sub

ErrorHandler:
Debug.Print "there is err", ws.Range("A1")
Exit Sub

End Sub


Sub calculate_nominal_multipe()
'On Error GoTo ErrorHandler

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
            
sheets("dissagregation_setting").Visible = True
Set ws = sheets("keen")
Set ws2 = sheets("keen2")
Set result_sheet = sheets("result")

last_col = ws.Cells(1, columns.count).End(xlToLeft).Column
last_col_letter = number_to_letter(last_col, ws)
new_col_letter = number_to_letter(last_col + 1, ws)
last_row = ws.Cells(ws.rows.count, "A").End(xlUp).Row

ws.columns("C:M").NumberFormat = "@"

unique_data_arr = extract_all_options()
small_unique_arr = extract_all_options()

With sheets("dissagregation_setting")
    last_dis = .Cells(rows.count, 1).End(xlUp).Row
    dis_arr2 = .Range("A2:C" & last_dis)
End With

arr = ws.Range("A1").CurrentRegion
arr2 = ws2.Range("A1").CurrentRegion
Set keen2_rng = ws2.Range("A1").CurrentRegion

'Debug.Print keen2_rng.rows.count

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
        
'            Debug.Print v, data_count, sum_weight, sum_weight_in_var, UBound(arr, 1) - 1
            n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
            result_sheet.Cells(n, 1) = n - 1
            result_sheet.Cells(n, 2) = "ALL"
            result_sheet.Cells(n, 3) = "ALL"
            result_sheet.Cells(n, 4) = "ALL"
            result_sheet.Cells(n, 5) = arr(1, 1)
            result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
            result_sheet.Cells(n, 7) = keen2_rng.rows.count - 1
            result_sheet.Cells(n, 8) = "percentage"
            result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(sum_weight_in_var / sum_weight * 100, 1)
            result_sheet.Cells(n, 10) = data_count
            result_sheet.Cells(n, 11) = v
            result_sheet.Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(v))
            result_sheet.Cells(n, 13) = "w"

        Next v
 
    ElseIf dis_arr2(i, 1) = "ALL" And dis_arr2(i, 2) = "no" Then

        Erase data_arr
        ReDim data_arr(1 To UBound(arr, 1) - 1)
        
        For j = 2 To UBound(arr, 1)
            data_arr(j - 1) = arr(j, 1)
        Next j
        
        For Each v In unique_data_arr
            data_count = count_in_array(data_arr, v)
'            Debug.Print v, data_count, UBound(arr, 1) - 1
            n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
            result_sheet.Cells(n, 1) = n - 1
            result_sheet.Cells(n, 2) = "ALL"
            result_sheet.Cells(n, 3) = "ALL"
            result_sheet.Cells(n, 4) = "ALL"
            result_sheet.Cells(n, 5) = arr(1, 1)
            result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
            result_sheet.Cells(n, 7) = keen2_rng.rows.count - 1
            result_sheet.Cells(n, 8) = "percentage"
            result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(data_count / (keen2_rng.rows.count - 1) * 100, 1)
            result_sheet.Cells(n, 10) = data_count
            result_sheet.Cells(n, 11) = v
            result_sheet.Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(v))
     
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
                    ' sum_weight = sum_weight + arr(j, 2)
                    k = k + 1
                End If
            Next j
'            Debug.Print "hoho!", dis_arr2(i, 1), col_n, v, dis_value_count, sum_weight
            
            temp_arr = Split(Mkr & Join(small_arr, Mkr & Del & Mkr) & Mkr, Del)
            'Count the items (Surrounded by markers) directly
            sum_weight = sum_weight_when(arr2, CStr(v), col_n)
            
            dis_value_count2 = Application.WorksheetFunction.CountIf(keen2_rng.columns(col_n), v)
'            Debug.Print dis_value_count2
            For Each p In small_unique_arr
                
                sum_weight_in_var = 0
                For j = 2 To UBound(arr, 1)
                    If p = arr(j, 1) And v = arr(j, col_n) Then
                        sum_weight_in_var = sum_weight_in_var + arr(j, 2)
                    End If
                Next j
            
                choice_count = UBound(Filter(temp_arr, Mkr & CStr(p) & Mkr, True, 1)) + 1
'                Debug.Print p, choice_count, " ", sum_weight_in_var
                n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
                result_sheet.Cells(n, 1) = n - 1
                result_sheet.Cells(n, 2) = dis_arr2(i, 1)
                result_sheet.Cells(n, 3) = v
                result_sheet.Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                result_sheet.Cells(n, 5) = arr(1, 1)
                result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
                result_sheet.Cells(n, 7) = dis_value_count2
                result_sheet.Cells(n, 8) = "percentage"
                result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(sum_weight_in_var / sum_weight * 100, 1)
                result_sheet.Cells(n, 10) = choice_count
                result_sheet.Cells(n, 11) = p
                result_sheet.Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(p))
                result_sheet.Cells(n, 13) = "w"
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
            dis_value_count2 = Application.WorksheetFunction.CountIf(keen2_rng.columns(col_n), v)
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
'                Debug.Print p, choice_count
                n = result_sheet.Cells(rows.count, 1).End(xlUp).Row + 1
                result_sheet.Cells(n, 1) = n - 1
                result_sheet.Cells(n, 2) = dis_arr2(i, 1)
                result_sheet.Cells(n, 3) = v
                result_sheet.Cells(n, 4) = choice_label(CStr(dis_arr2(i, 1)), CStr(v))
                result_sheet.Cells(n, 5) = arr(1, 1)
                result_sheet.Cells(n, 6) = var_label(CStr(arr(1, 1)))
                result_sheet.Cells(n, 7) = dis_value_count2
                result_sheet.Cells(n, 8) = "percentage"
                result_sheet.Cells(n, 9) = Application.WorksheetFunction.Round(choice_count / dis_value_count2 * 100, 1)
                result_sheet.Cells(n, 10) = choice_count
                result_sheet.Cells(n, 11) = p
                result_sheet.Cells(n, 12) = choice_label(CStr(arr(1, 1)), CStr(p))
            Next p

        Next v
    
    End If
    
NextIteration:
Next i

Exit Sub

ErrorHandler:
Debug.Print "there is err", ws.Range("A1")
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
    
    For Each v In dict.Keys
'        Debug.Print v & " " & dict.item(v)
    Next v
End Sub

Function extract_all_options() As Variant
    On Error GoTo 0
    Dim xsc_arr As Variant
    Dim data_arr As Variant
    Dim dict As Object
    Dim i As Long, v As Variant
    Dim main_var As String
    Dim ws As Worksheet
    
    Set ws = sheets("keen")
    Set dict = CreateObject("Scripting.Dictionary")
    xsc_arr = ThisWorkbook.sheets("xsurvey_choices").Range("A1").CurrentRegion
    main_var = ws.Range("A1")
    
    For i = 1 To UBound(xsc_arr, 1)
        If xsc_arr(i, 2) = main_var Then
            dict.Add xsc_arr(i, 4), i
        End If
    Next i
    
'    Debug.Print "count dict: " & dict.count
    
    If dict.count > 1 Then
        extract_all_options = dict.Keys()
    Else
        data_arr = ws.Range(ws.Range("A2"), ws.Range("A2").End(xlDown))
        For i = LBound(data_arr, 1) To UBound(data_arr, 1)
            dict(data_arr(i, 1)) = i
        Next i
        extract_all_options = dict.Keys()
    End If
     
'    For Each v In dict.Keys
'        Debug.Print v & " " & dict.item(v)
'    Next v
'
'    Debug.Print main_var, dict.count
    
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
    
    Set ws = sheets("keen")
    ws.columns("O:AZ").Clear
    ws.columns("A:AZ").NumberFormat = "@"
    last_col = ws.Cells(1, columns.count).End(xlToLeft).Column
    last_col_letter = number_to_letter(last_col, ws)
    last_row = ws.Cells(ws.rows.count, "A").End(xlUp).Row
    
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
    
    Set SELECT_MULTIPLE_COLL = New Collection
   
    Dim dis_arr As Variant
    Dim data_arr As Variant
    Dim total_sum As Single
    Dim total_count As Long
    Dim unique_data_arr As Variant
    Dim v As Variant
    Dim keen2_ws As Worksheet

    'Set SELECT_MULTIPLE_COLL = Nothing
    If Not worksheet_exists("keen2") Then
        Call create_sheet(find_main_data, "keen2")
            sheets("keen2").Visible = xlVeryHidden
    End If
    
    Set keen2_ws = sheets("keen2")
    keen2_ws.Cells.Clear
    
    If WITH_WEIGHT Then
        keen2_ws.columns("A:A").NumberFormat = "@"
        keen2_ws.columns("B:B").NumberFormat = "0.000"
        keen2_ws.columns("C:Z").NumberFormat = "@"
    Else
        keen2_ws.columns("A:Z").NumberFormat = "@"
    End If
    
    ws.Range("A1").CurrentRegion.Copy
    keen2_ws.Range("A1").PasteSpecial xlPasteValues

    ws.columns("A:N").Delete
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
'            Debug.Print arr(i, col_index)
            sum = sum + arr(i, 2)
        End If
    Next i
    sum_weight_when = sum
End Function

Sub check_result_sheet(sheet_name As String)
    ' check if keen sheet exist
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    If Not worksheet_exists("result") Then
        Call create_sheet(sheet_name, "result")
        wb.sheets("result").Cells(1, 1) = "row"
        wb.sheets("result").Cells(1, 2) = "disaggregation"
        wb.sheets("result").Cells(1, 3) = "disaggregation value"
        wb.sheets("result").Cells(1, 4) = "disaggregation label"
        wb.sheets("result").Cells(1, 5) = "variable"
        wb.sheets("result").Cells(1, 6) = "variable label"
        wb.sheets("result").Cells(1, 7) = "valid numbers"
        wb.sheets("result").Cells(1, 8) = "measurement type"
        wb.sheets("result").Cells(1, 9) = "measurement value"
        wb.sheets("result").Cells(1, 10) = "count"
        wb.sheets("result").Cells(1, 11) = "choice"
        wb.sheets("result").Cells(1, 12) = "choice label"
        wb.sheets("result").Cells(1, 13) = "weight"
        wb.sheets("result").Cells(1, 14) = "hkey"
        wb.sheets("result").Cells(1, 15) = "hkey order"
        
        wb.sheets("result").columns(1).ColumnWidth = 6
        wb.sheets("result").columns(2).ColumnWidth = 15
        wb.sheets("result").columns(3).ColumnWidth = 18
        wb.sheets("result").columns(4).ColumnWidth = 25
        wb.sheets("result").columns(5).ColumnWidth = 15
        wb.sheets("result").columns(6).ColumnWidth = 45
        wb.sheets("result").columns(7).ColumnWidth = 15
        wb.sheets("result").columns(8).ColumnWidth = 15
        wb.sheets("result").columns(9).ColumnWidth = 20
        wb.sheets("result").columns(10).ColumnWidth = 10
        wb.sheets("result").columns(11).ColumnWidth = 15
        wb.sheets("result").columns(12).ColumnWidth = 45
        wb.sheets("result").columns(13).ColumnWidth = 7
        wb.sheets("result").columns(14).ColumnWidth = 45
        wb.sheets("result").columns(15).ColumnWidth = 15
    End If
    
    wb.sheets("result").columns("B:F").NumberFormat = "@"
    wb.sheets("result").columns("K:M").NumberFormat = "@"
    wb.sheets("result").Visible = False
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

'Dim t1 As Double
't1 = Timer

    Dim rng As Range
    Application.ScreenUpdating = False
    Dim res_ws As Worksheet
    Dim ws As Worksheet
    Dim last_result As Long
    Dim last_header As Long
    
    Set res_ws = sheets("result")
    last_result = res_ws.Cells(res_ws.rows.count, "A").End(xlUp).Row
    
    res_ws.Activate
    res_ws.Range("N2:N" & last_result).Formula = "=E2&K2&H2"
    res_ws.columns("N:N").Copy
    res_ws.columns("N:N").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
    Set ws = sheets("temp_sheet")
    ws.Cells.Clear
    
    ws.Range("A1:A" & last_result).Value2 = res_ws.Range("N1:N" & last_result).Value2
    ws.Range("A1").CurrentRegion.RemoveDuplicates columns:=1, Header:=xlYes
    
    last_header = ws.Cells(ws.rows.count, "A").End(xlUp).Row
     
    ws.Range("B2") = "1"
    ws.Range("B2").AutoFill Destination:=ws.Range("B2:B" & last_header), Type:=xlFillSeries
    
    ws.columns("B:B").Copy
    ws.columns("B:B").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
         
    res_ws.Range("O2:O" & last_result).Formula = "=VLOOKUP(result!N2,temp_sheet!A:B,2,)"
    
    res_ws.columns("O:O").Copy
    res_ws.columns("O:O").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
' Debug.Print Timer - t1
End Sub

