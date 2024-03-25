VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} extra_logs_form 
   Caption         =   "Logbook option"
   ClientHeight    =   1866
   ClientLeft      =   -300
   ClientTop       =   -1344
   ClientWidth     =   7470
   OleObjectBlob   =   "extra_logs_form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "extra_logs_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandAdd_Click()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    If Me.ComboQuestion.value = vbNullString Then
        Exit Sub
    End If
    
    Dim log_ws As Worksheet
    Dim dt_ws As Worksheet
    Dim last_log As Long, last_col As Long
    Dim res As Variant
    Dim header_rng As Range
    Dim i As Long
    Dim uuid_coln As Long
    
    Set dt_ws = sheets(find_main_data)
    Set log_ws = sheets("log_book")
    
    Call clear_filter(dt_ws)
    Call clear_filter(log_ws)
    
    Set header_rng = log_ws.Range(log_ws.Cells(1, 1), log_ws.Cells(1, 1).End(xlToRight))
    i = 0
    For i = 1 To header_rng.Columns.count
    
        If Me.ComboQuestion.value = log_ws.Cells(1, i) Then
            new_col = i
            Exit For
        Else
            new_col = log_ws.Cells(1, Columns.count).End(xlToLeft).Column + 1
        End If
        
    Next
    
    If new_col > 14 Then
        MsgBox "You have reached to maximum extra columns in the logbook. ", vbInformation
        Exit Sub
    End If
    
    question_col_letter = data_column_letter(Me.ComboQuestion)
    uuid_col_letter = data_column_letter("_uuid")
    last_log = log_ws.Cells(Rows.count, 1).End(xlUp).Row
    
    uuid_coln = gen_column_number("_uuid", find_main_data)
    last_dt = dt_ws.Cells(Rows.count, uuid_coln).End(xlUp).Row
  
    new_col_letter = Split(log_ws.Cells(1, new_col).Address, "$")(1)
'    Debug.Print new_col, new_col_letter
    
    Dim uuid_rng As Range
    Set uuid_rng = log_ws.Range("A1:A" & last_log)
    
    log_ws.Cells(1, new_col) = Me.ComboQuestion.value
    
    For j = 2 To last_log
        res = Application.index(dt_ws.Range(question_col_letter & "2:" & question_col_letter & last_dt), _
                                Application.Match(log_ws.Cells(j, 1), _
                                                  dt_ws.Range(uuid_col_letter & "2:" & uuid_col_letter & last_dt), 0))
                                           
        If IsError(res) Then
            log_ws.Cells(j, new_col) = ""
        Else
            log_ws.Cells(j, new_col) = res
        End If
    Next j
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CommandLogDuplicate_Click()
    On Error Resume Next
    Call find_duplicate_log
End Sub

Private Sub CommandRemoveduplicatedLogs_Click()
    On Error Resume Next
    Call remove_duplicate_log
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    If Not worksheet_exists("log_book") Then
        MsgBox "The logbook dose not exist!   ", vbInformation
        End
    End If
    
    sheets("log_book").Activate
    
    PopulateComboBox
    
End Sub

Private Sub PopulateComboBox()

    Dim header_arr() As Variant
    Dim ws As Worksheet
    
    Set ws = sheets(find_main_data)

    header_arr = ws.Range(ws.Cells(1, 1), ws.Cells(1, 1).End(xlToRight)).Value2
    
    With Application
        header_arr = .Transpose(.Transpose(header_arr))
    End With
    
    Me.ComboQuestion.List = header_arr
    
End Sub

