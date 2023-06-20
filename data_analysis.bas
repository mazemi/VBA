Attribute VB_Name = "data_analysis"
Sub abc()
Dim new_row, sum_w2, sum_weight, mean_w_value, last_row_survey_sheet As Long

Dim simple_mean, simple_count As Long
Dim main_ws As Worksheet
Set main_ws = sheets("RAM2")
Dim weighting, numeric As Boolean
Dim unique_diss_options As Range, Item As Range
Dim ws As Worksheet
Dim diss As String

Dim keenSheet As Worksheet

diss = "province"
weighting = False
numeric = True

'* check if diss var is a select on or not?
diss_type = match_type(diss)

If Left(diss_type, 10) <> "select_one" Then
    Debug.Print "exit sub"
    Exit Sub
End If

'* find diss labels and all choices:
Dim SourceRange As Range, cl As Range, rng As Range
Dim DestSheet As Worksheet
Dim Criteria As String
Dim filtered_col As Long
Dim choice_coll As New Collection
Dim i As Integer
        
key_name_arr = Split(diss_type, " ")
key_name = key_name_arr(UBound(key_name_arr))

last_row_choice = Worksheets("choices").Cells(Rows.Count, 1).End(xlUp).Row
last_row_survey_sheet = Worksheets("survey").Cells(Rows.Count, 1).End(xlUp).Row

Set SourceRange = Worksheets("choices").Range("A1:C" & last_row_choice)
SourceRange.AutoFilter
SourceRange.AutoFilter Field:=1, Criteria1:=key_name
SourceRange.SpecialCells (xlCellTypeVisible)

Set unique_diss_options = SourceRange.Columns(SourceRange.Columns.Count - 1).Cells

' check if xx sheet exist
If WorksheetExists("result") <> True Then
    Call create_sheet(main_ws.Name, "result")
    sheets("result").Cells(1, 1) = "disaggregation"
    sheets("result").Cells(1, 2) = "disaggregation value"
    sheets("result").Cells(1, 3) = "variable"
    sheets("result").Cells(1, 4) = "variable label"
    sheets("result").Cells(1, 5) = "measurement type"
    sheets("result").Cells(1, 6) = "measurement value"
    sheets("result").Cells(1, 7) = "measurement numbers"
End If

Set ws = sheets("result")

new_row = ws.UsedRange.Rows(ws.UsedRange.Rows.Count).Row + 1

If WorksheetExists("keen") <> True Then
    Call create_sheet(main_ws.Name, "keen")
End If

Set keenSheet = sheets("keen")

' copy necesory data to sheet keen sheet
sheets("RAM2").Range("A:A,I:I,AE:AE").Copy sheets("keen").Range("A:C")
last_row_keen = sheets("keen").Cells(Rows.Count, 1).End(xlUp).Row

Dim small_dt As Range
Set small_dt = sheets("keen").Cells(1, 1).CurrentRegion

Call remove_rows(small_dt)

last_row_keen = sheets("keen").Cells(Rows.Count, 1).End(xlUp).Row
'* populate the result sheet
If weighting And numeric Then

    Call add_mulitipication
    
    For Each Item In unique_diss_options.SpecialCells(xlCellTypeVisible)
        If Item <> "name" Then
            sum_w2 = Application.WorksheetFunction.SumIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), Item, keenSheet.Range("D2:D" & CStr(last_row_keen)))
            sum_weight = Application.WorksheetFunction.SumIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), Item, keenSheet.Range("C2:C" & CStr(last_row_keen)))
            count_weight = Application.WorksheetFunction.CountIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), Item)
            If sum_w2 > 0 Then
                ws.Cells(new_row, 1) = diss
                ws.Cells(new_row, 2) = Item
                mean_w_value = Round(sum_w2 / sum_weight, 1)
                ws.Cells(new_row, 3) = Worksheets("keen").Cells(1, 2)
                ws.Cells(new_row, 4) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), WorksheetFunction.Match(keenSheet.Cells(1, 2), _
                    sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))
                
                ws.Cells(new_row, 5) = "mean"
                ws.Cells(new_row, 6) = mean_w_value
                ws.Cells(new_row, 7) = count_weight
                
                Debug.Print sum_w2, sum_weight, mean_w_value
                
            Else
                ws.Cells(new_row, 1) = diss
                ws.Cells(new_row, 2) = Item
                ws.Cells(new_row, 3) = Worksheets("keen").Cells(1, 2)
                ws.Cells(new_row, 4) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), WorksheetFunction.Match(keenSheet.Cells(1, 2), _
                    sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))
                ws.Cells(new_row, 5) = "mean"
                ws.Cells(new_row, 7) = 0
            End If
            
            new_row = new_row + 1
        End If
    Next
ElseIf Not weighting And numeric Then
    For Each Item In unique_diss_options.SpecialCells(xlCellTypeVisible)
        Debug.Print Item
        If Item <> "name" Then
            simple_count = WorksheetFunction.CountIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), Item)
            If simple_count > 0 Then
                simple_mean = WorksheetFunction.AverageIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), Item, keenSheet.Range("B2:B" & CStr(last_row_keen)))
                ws.Cells(new_row, 1) = diss
                ws.Cells(new_row, 2) = Item
                         
                ws.Cells(new_row, 3) = Worksheets("keen").Cells(1, 2)
                ws.Cells(new_row, 4) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), WorksheetFunction.Match(keenSheet.Cells(1, 2), _
                    sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))
                
                ws.Cells(new_row, 5) = "mean"
                ws.Cells(new_row, 6) = Round(simple_mean, 1)
                ws.Cells(new_row, 7) = simple_count
            Else
                ws.Cells(new_row, 1) = diss
                ws.Cells(new_row, 2) = Item
                ws.Cells(new_row, 3) = Worksheets("keen").Cells(1, 2)
                ws.Cells(new_row, 4) = WorksheetFunction.Index(sheets("survey").Range("C2:C" & last_row_survey_sheet), WorksheetFunction.Match(keenSheet.Cells(1, 2), _
                    sheets("survey").Range("B2:B" & last_row_survey_sheet), 0))
                ws.Cells(new_row, 5) = "mean"
                ws.Cells(new_row, 7) = 0
            End If
            
            new_row = new_row + 1
        End If
    Next
End If

sheets("keen").Cells.Clear
'Debug.Print last_row_keen

'* check if weight in the diss


'* check the mesurement type (number, catagory or multi catagory)


'* if the mesurement type is number


'* if the mesurement type is catagory


'* if the mesurement type is multi catagory
End Sub

Private Sub remove_rows(rng As Range)

    Dim Lrange As Range
    Dim n As Long
    
    Set Lrange = rng

    For n = Lrange.Rows.Count To 1 Step -1
        If Lrange.Cells(n, 2).Value = "" Then
            Lrange.Rows(n).Delete
        End If
    Next
End Sub

Sub add_mulitipication()
    Dim last_row As Long
    last_row = Worksheets("keen").Cells(Rows.Count, 1).End(xlUp).Row
    Worksheets("keen").Range("D1").FormulaR1C1 = "w2"
    Application.CutCopyMode = False
    Worksheets("keen").Range("D2").FormulaR1C1 = "=RC[-2]*RC[-1]"
    Worksheets("keen").Range("D2").AutoFill Destination:=Worksheets("keen").Range("D2:D" & CStr(last_row))
End Sub

Sub kk()
simple_mean = WorksheetFunction.AverageIf(keenSheet.Range("A2:A" & CStr(last_row_keen)), Item, keenSheet.Range("B2:" & CStr(last_row_keen)))
End Sub


