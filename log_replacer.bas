Attribute VB_Name = "log_replacer"
Sub replace_log()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim ws_main As Worksheet
    Dim ws_log As Worksheet
    Set ws_main = ActiveWorkbook.Sheets("RAM2")
    Set ws_log = ActiveWorkbook.Sheets("log_book")
    Dim cell As Range
    Dim log_uuid_col_number, log_last_col, log_question_name_col_number, _
        log_new_value_col_number, log_changed_col_number, row_number As Long
    Dim question As String
    Dim lastColLetter As String
    Dim rng_log_question As Range, rng_header As Range
    Dim result As Variant
    
    ws_log.Select
    log_uuid_col_number = column_number("uuid")
    log_question_name_col_number = column_number("question.name")
    log_question_name_col_letter = column_letter("question.name")
    log_new_value_col_number = column_number("new.value")
    log_changed_col_number = column_number("changed")
    
    log_last_col = Cells(1, Columns.Count).End(xlToLeft).Column
    log_last_row = ws_log.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).row
    
    lastCol = ws_main.UsedRange.Columns.Count
    lastColLetter = Split(Cells(, lastCol).Address, "$")(1)

    ' new column for remarks in the log book shet:
    ws_log.Cells(1, log_last_col + 1) = "remarks"

    ws_main.Select
    main_uuid_col_number = column_letter("_uuid")
    
    For i = 2 To 44
        Application.StatusBar = "Replacing clean data : " & i
        ' if changes in log book is yes:
        If LCase(ws_log.Cells(i, log_changed_col_number)) = "yes" Then
            uuid_value = ws_log.Cells(i, log_uuid_col_number)
            question = ws_log.Cells(i, log_question_name_col_number)
            clean_value = ws_log.Cells(i, log_new_value_col_number)
                     
            uuid_exist = False
            
            ' find row number in the main sheet based on the uuid value:
            ' use an array to store the values in the main uuid column
            Dim main_uuid_values() As Variant
            main_uuid_values = ws_main.Range(main_uuid_col_number & ":" & main_uuid_col_number).Value2
            
            ' loop through the array instead of the cells
            Dim k As Long
            For k = 1 To UBound(main_uuid_values)
                If main_uuid_values(k, 1) = uuid_value Then
                    row_number = k
                    uuid_exist = True
                    Exit For
                End If
            Next k
            
            If Not uuid_exist Then
                ws_log.Cells(i, log_last_col + 1) = ws_log.Cells(i, log_last_col + 1) & " uuid not found"
            End If
            
            col_number = column_number(question)
             
            If col_number = 0 Then
'                 Debug.Print "  ***  " & question
                 ws_log.Cells(i, log_last_col + 1) = ws_log.Cells(i, log_last_col + 1) & " question not found"
            Else
                ' replace clean value in the cell:
                ws_main.Cells(row_number, col_number).Interior.Color = xlNone
                ws_main.Cells(row_number, col_number) = clean_value
            End If
            
        End If

    Next i
    
    ' Application.ScreenUpdating = True
    MsgBox "All cleaned data has been replaced.                  ", vbInformation
    ws_log.Activate
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub








