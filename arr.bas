Attribute VB_Name = "arr"

Sub zz()
    Dim arr() As String
    Dim multi_df As New Array2D
    multi_df.data = Null
    multi_df.insertColumnsBlank 1, 2
    
    Dim rng As Range, cel As Range
    Set rng = sheets("Sheet3").Range("A1").CurrentRegion
    
    For Each cel In rng.columns(2).Cells
        arr = Split(cel, " ")
        For Each Item In arr
            multi_df.insertRowsBlank multi_df.rowCount + 1
            multi_df.value(multi_df.rowCount, 1) = rng.Cells(cel.row, 1)
            multi_df.value(multi_df.rowCount, 2) = Item
            multi_df.value(multi_df.rowCount, 3) = rng.Cells(cel.row, 3)
        Next
    Next
    
multi_df.writeDataToRange sheets("Sheet3").Range("E1")
Debug.Print "array done.   "
    
Dim unique_values_col As New Collection

last = sheets("Sheet3").Range("F" & rows.count).End(xlUp).row

Set unique_values_col = unique_values(sheets("Sheet3").Range("F2:F" & last))

Debug.Print "collection done.   "
Debug.Print unique_values_col.count

For Each i In unique_values_col
    Debug.Print i
Next

'multi_df.dataFromRange rg:=rng, removeHeader:=True

End Sub

Sub make_survey_choice()

Dim survey_df As New Array2D
Dim choices_df As New Array2D
Dim full_df As New Array2D
Dim int_df As New Array2D, dec_df As New Array2D, calc_df As New Array2D
Dim select_one_df As New Array2D, select_multiple_df As New Array2D
Dim good_df As New Array2D

full_df.data = Null

full_df.insertColumnsBlank 1, 4

Dim s_rng As Range, c_rng As Range

Set s_rng = sheets("survey").Range("A1").CurrentRegion
Set c_rng = sheets("choices").Range("A1").CurrentRegion

survey_df.dataFromRange rg:=s_rng, removeHeader:=True
survey_df.insertColumnsBlank survey_df.columnCount + 1
choices_df.dataFromRange rg:=c_rng, removeHeader:=True

For i = 1 To survey_df.rowCount
    If InStr(survey_df.value(i, 1), " ") Then
        Leftx = Left(survey_df.value(i, 1), InStrRev(survey_df.value(i, 1), " ") - 1)
        Rightx = Right(survey_df.value(i, 1), Len(survey_df.value(i, 1)) - InStrRev(survey_df.value(i, 1), " "))
        survey_df.value(i, 1) = Leftx
        survey_df.value(i, survey_df.columnCount) = Rightx
    End If
Next i

Dim k
k = 1

While k <= survey_df.rowCount
    m = full_df.rowCount
    
    For j = 1 To choices_df.rowCount
        
        If choices_df.value(j, 1) = survey_df.value(k, 4) Then
            n = full_df.rowCount
            full_df.value(n, 1) = survey_df.value(k, 1)
            full_df.value(n, 2) = survey_df.value(k, 2)
            full_df.value(n, 3) = survey_df.value(k, 3)
            full_df.value(n, 4) = choices_df.value(j, 2)
            full_df.value(n, 5) = choices_df.value(j, 3)
            full_df.insertRowsBlank full_df.rowCount + 1
       End If
       
    Next
    
    If survey_df.value(k, 1) <> "" Then
        full_df.value(m, 1) = survey_df.value(k, 1)
        full_df.value(m, 2) = survey_df.value(k, 2)
        full_df.value(m, 3) = survey_df.value(k, 3)
        full_df.insertRowsBlank full_df.rowCount + 1
    End If
    k = k + 1
Wend

int_df.data = full_df.Filter("integer")
dec_df.data = full_df.Filter("decimal")
calc_df.data = full_df.Filter("calculate")
select_one_df.data = full_df.Filter("select_one")
select_multiple_df.data = full_df.Filter("select_multiple")

full_df.data = full_df.Filter("select_one")

full_df.insertRows int_df.rowCount, int_df.data

good_df.data = int_df.data
'good_df.insertRows dec_df.rowCount, dec_df.data
good_df.insertRows calc_df.rowCount, calc_df.data
good_df.insertRows select_one_df.rowCount, select_one_df.data
good_df.insertRows select_multiple_df.rowCount, select_multiple_df.data

good_df.writeDataToRange sheets("survey_choices").Range("A2")
    
Set survey_df = Nothing
Set choices_df = Nothing
Set full_df = Nothing
End Sub

Sub mm()

Dim col_number As Long
col_number = 3

Dim rng As Range
Dim str_delete As String

Set rng = sheets("keen").Range("A1").CurrentRegion

last_keen = rng.rows.count

rng.Sort columns(col_number), , , , , , , Header:=xlYes

last_measurement = Worksheets("keen").Cells(rows.count, col_number).End(xlUp).row

str_delete = CStr(last_measurement + 1) & ":" & last_keen

If last_measurement < last_keen Then

    Debug.Print str_delete

End If

rows(str_delete).Delete Shift:=xlUp

End Sub


