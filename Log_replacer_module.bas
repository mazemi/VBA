Attribute VB_Name = "Log_replacer_module"
Option Explicit

Sub replace_log()

    On Error GoTo errHandler
    Application.ScreenUpdating = False
    Dim ws_main As Worksheet
    Dim ws_log As Worksheet
    Dim cell As Range
    Dim log_uuid_col_number, log_last_col, log_question_name_col_number As Long
    Dim log_new_value_col_number, log_changed_col_number, row_number As Long
    Dim question As String
    Dim lastColLetter As String
    Dim rng_log_question As Range, rng_header As Range
    Dim Result As Variant
    Dim i As Long, j As Long
    Dim dt As String
    Dim log_question_name_col_letter As String
    Dim log_last_row As Long
    Dim LastCol As Long
    Dim main_uuid_col_numbera As String
    Dim main_uuid_col_number As String
    Dim uuid_value As String
    Dim clean_value As String
    Dim uuid_exist As Boolean
    Dim cn As Long
    
    ' check if the KOBO tool exist
    If ThisWorkbook.Worksheets("xsurvey").Range("A1") = vbNullString Then
        MsgBox "Please import the KOBO tools.    ", vbInformation
        Exit Sub
    End If

    dt = find_main_data
    Set ws_main = ActiveWorkbook.sheets(dt)
    
    If Not worksheet_exists("log_book") Then
        MsgBox "The log_book sheet dose not exist!     ", vbInformation
        Exit Sub
    End If
    
    Call remove_empty_col
    
    Set ws_log = ActiveWorkbook.sheets("log_book")
      
    ws_log.Select
    log_uuid_col_number = column_number("uuid")
    log_question_name_col_number = column_number("question.name")
    log_question_name_col_letter = column_letter("question.name")
    log_new_value_col_number = column_number("new.value")
    log_changed_col_number = column_number("changed")
    
    log_last_col = ws_log.Cells(1, columns.count).End(xlToLeft).column
    log_last_row = ws_log.UsedRange.rows(ws_log.UsedRange.rows.count).row
    
    LastCol = ws_main.UsedRange.columns.count
    lastColLetter = Split(ws_main.Cells(, LastCol).Address, "$")(1)

    ' new column for remarks in the log book shet:
    If ws_log.Cells(1, log_last_col) <> "remarks" Then
        ws_log.Cells(1, log_last_col + 1) = "remarks"
    Else
        log_last_col = log_last_col - 1
    End If
    
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
            ' use an array to store the values in the main uuid column
            Dim main_uuid_values() As Variant
            main_uuid_values = ws_main.Range(main_uuid_col_number & ":" & main_uuid_col_number).Value2
            
            ' loop through the array of main uuid
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
            
            cn = column_number(question)
             
            If cn = 0 Then
                ws_log.Cells(i, log_last_col + 1) = ws_log.Cells(i, log_last_col + 1) & " question not found"
            Else
                If is_multi_select(question) Then
                    Dim pos1 As Integer
                    Dim pos2 As Integer
                    Dim m_select_question_col As Long
                    Dim choice_val As String
                    Dim combined_str As String
                    
                    pos1 = InStr(question, ".")
                    pos2 = InStr(question, "/")
                    
                    m_select_question_col = column_number(left(question, Abs(pos1 - pos2) - 1))
                    choice_val = Right(question, Len(question) - Abs(pos1 - pos2))
                     
                    If clean_value = 0 Then
                        ws_main.Cells(row_number, m_select_question_col) = _
                            Replace(ws_main.Cells(row_number, m_select_question_col), choice_val, "")
                        Trim (ws_main.Cells(row_number, m_select_question_col))
                    ElseIf clean_value = 1 Then
                        combined_str = ws_main.Cells(row_number, m_select_question_col) & " " & choice_val
                        ws_main.Cells(row_number, m_select_question_col) = _
                            clean_multi_select(combined_str)
                    End If
                     
                End If
                ' replace clean value in the cell:
                ws_main.Cells(row_number, cn).Interior.Color = xlNone
                ws_main.Cells(row_number, cn) = clean_value
                j = j + 1
            End If
            
        End If

    Next i
    
    Application.ScreenUpdating = True
    If j > 0 Then
        MsgBox "The cleaning logs have been replaced.                      ", vbInformation
    Else
        MsgBox "No replacement!" & vbCrLf & _
               "Please check your log_book, if the 'changed' column has been set or not.", vbInformation
    End If
    
    Application.StatusBar = False
    Exit Sub
    
errHandler:
    MsgBox "The log replacement failed! Pleae check your logbook and main data set and the integrated tool, then try again.", vbInformation
    Exit Sub
    
End Sub

Function is_multi_select(str As String) As Boolean
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim q_type As String
    pos1 = InStr(str, ".")
    pos2 = InStr(str, "/")
    If pos1 = 0 And pos2 = 0 Then
        is_multi_select = False
    Else
        q_type = question_type(left(str, Abs(pos1 - pos2) - 1))
        If q_type = "select_multiple" Then
            ' Debug.Print Abs(pos1 - pos2)
            is_multi_select = True
            Exit Function
        End If
    End If
    is_multi_select = False
End Function

Function clean_multi_select(str As String) As String
    Dim arr_strings() As String
    Dim d As Object, c As Variant
    Dim res, k, tmp As String
    str = Trim(str)
    arr_strings = Split(str, " ")
    
    Set d = CreateObject("scripting.dictionary")
    For Each c In arr_strings
        tmp = Trim(c)
        If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
    Next c

    For Each k In d.Keys
        ' Debug.Print k, d(k)
        res = res + " " & k
    Next k
    
    clean_multi_select = Trim(res)
    
End Function

Sub remove_duplicated_logs()

    sheets("log_book").Range("A1").CurrentRegion.RemoveDuplicates columns:=9, Header:=xlYes
    sheets("log_book").columns("H:I").Delete Shift:=xlToLeft

End Sub