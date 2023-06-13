Attribute VB_Name = "log_replacer"
Sub replace_log()

    Dim ws_main As Worksheet
    Dim ws_log As Worksheet
    Set ws_main = ActiveWorkbook.Sheets("RAM2")
    Set ws_log = ActiveWorkbook.Sheets("log_book")
    Dim cell As Range
    Dim log_uuid_col_number, log_last_col, log_question_name_col_number, _
        log_new_value_col_number, log_changed_col_number, row_number As Long
    Dim question As String
    
    ' turn off screen updating and automatic calculation
    Dim savedScreenUpdating As Boolean
    Dim savedCalculation As Long
    savedScreenUpdating = Application.ScreenUpdating
    savedCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

'    Application.ScreenUpdating = False
    
    ws_log.Select
    log_uuid_col_number = column_number("uuid")
    log_question_name_col_number = column_number("question.name")
    log_new_value_col_number = column_number("new.value")
    log_changed_col_number = column_number("changed")
    
    log_last_col = Cells(1, Columns.Count).End(xlToLeft).Column
    log_last_row = ws_log.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).row
    
    ws_main.Select
    main_uuid_col_number = column_letter("_uuid")
    
    For i = 2 To log_last_row
        Application.StatusBar = "Replacing clean data : " & i
        ' if changes in log book is yes:
        If LCase(ws_log.Cells(i, log_changed_col_number)) = "yes" Then
            uuid_value = ws_log.Cells(i, log_uuid_col_number)
            question = ws_log.Cells(i, log_question_name_col_number)
            clean_value = ws_log.Cells(i, log_new_value_col_number)
                     
            uuid_exist = False
            
            ' find row number in the main sheet based on the uuid value:
'            For Each cell In ws_main.Range(main_uuid_col_number & ":" & main_uuid_col_number)
'                If cell.Value = uuid_value Then
'                    row_number = cell.row
'                    uuid_exist = True
'                    Exit For
'                End If
'            Next cell
            
            ' use an array to store the values in the main uuid column
            Dim main_uuid_values() As Variant
            main_uuid_values = ws_main.Range(main_uuid_col_number & ":" & main_uuid_col_number).Value2
            
            ' loop through the array instead of the cells
            Dim k As Long
            For k = 1 To UBound(main_uuid_values)
                If main_uuid_values(k, 1) = uuid_value Then
                    row_number = ik
                    uuid_exist = True
                    Exit For
                End If
            Next k
            
            If Not uuid_exist Then
'                Debug.Print uuid_value
                ws_log.Cells(1, log_last_col + 1) = "remarks"
                ws_log.Cells(i, log_last_col + 1) = "uuid not found"
            End If
            
            col_number = column_number(question)
            
            If col_nunber = 0 Then
                ' Debug.Print "  ***  " & question
                
            Else
                ' replace clean value in the cell:
                ws_main.Cells(row_number, col_number) = clean_value
            End If
            
        End If

    Next i
    
'    Application.ScreenUpdating = True
    MsgBox "All cleaned data has been replaced.                  ", vbInformation
    
    ' turn on screen updating and automatic calculation
    Application.ScreenUpdating = savedScreenUpdating
    Application.Calculation = savedCalculation
    
End Sub


' find row number in the main sheet based on the uuid value:







