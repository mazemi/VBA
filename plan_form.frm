VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} plan_form 
   Caption         =   "Logical Data Cleaning "
   ClientHeight    =   3528
   ClientLeft      =   -306
   ClientTop       =   -1344
   ClientWidth     =   10434
   OleObjectBlob   =   "plan_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "plan_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Sub ComboOp1_Change()
    If Me.ComboOp1.value = "is empty" Or Me.ComboOp1.value = "is not empty" Then
        Me.ComboAns1.value = vbNullString
        Me.ComboAns1.Enabled = False
    Else
        Me.ComboAns1.Enabled = True
    End If
End Sub

Private Sub ComboOp2_Change()
    If Me.ComboOp2.value = "is empty" Or Me.ComboOp2.value = "is not empty" Then
        Me.ComboAns2.value = vbNullString
        Me.ComboAns2.Enabled = False
    Else
        Me.ComboAns2.Enabled = True
    End If
End Sub

Private Sub ComboQ1_Change()
    Dim last_row As Long
    Dim i As Long
    
    If Me.ComboQ1.value <> vbNullString Then
        Do While Me.ComboAns1.ListCount > 0
            Me.ComboAns1.RemoveItem (0)
        Loop
        
        Me.ComboAns1.value = vbNullString
        Call extract_choice(Me.ComboQ1)
        
        If ThisWorkbook.sheets("xsurvey_choices").Cells(2, "K") <> vbNullString Then
            last_row = ThisWorkbook.sheets("xsurvey_choices").Cells(Rows.count, "K").End(xlUp).Row
            For i = 2 To last_row
                Me.ComboAns1.AddItem (ThisWorkbook.sheets("xsurvey_choices").Cells(i, "K"))
            Next i
            Me.ComboAns1.Style = fmStyleDropDownList
        Else
            Me.ComboAns1.Style = fmStyleDropDownCombo
        End If
    End If

End Sub

Private Sub ComboQ2_Change()
    Dim last_row As Long
    Dim i As Long
    
    If Me.ComboQ2.value <> vbNullString Then
        Do While Me.ComboAns2.ListCount > 0
            Me.ComboAns2.RemoveItem (0)
        Loop
        
        Me.ComboAns2.value = vbNullString
        Call extract_choice(Me.ComboQ2)
        
        If ThisWorkbook.sheets("xsurvey_choices").Cells(2, "K") <> vbNullString Then
            last_row = ThisWorkbook.sheets("xsurvey_choices").Cells(Rows.count, "K").End(xlUp).Row
            For i = 2 To last_row
                Me.ComboAns2.AddItem (ThisWorkbook.sheets("xsurvey_choices").Cells(i, "K"))
            Next i
            Me.ComboAns2.Style = fmStyleDropDownList
        Else
            Me.ComboAns2.Style = fmStyleDropDownCombo
        End If
    End If

End Sub

Private Sub CommandCancel_Click()
    Unload Me
    If Public_module.PLAN_NUMBER > 0 Then
    plan_list_form.Show
    End If
End Sub

Private Sub CommanSave_Click()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim new_row As Long
    Dim main_operation As String
   
    Set ws = ThisWorkbook.sheets("xlogical_checks")
    
    If Me.OptionAnd Then
        main_operation = "and"
    ElseIf Me.OptionOr Then
        main_operation = "or"
    Else
        main_operation = vbNullString
    End If
   
    If Me.TextMassage = vbNullString Then
        MsgBox "This is not a valid checking role! Pleas set the message .   ", vbCritical
        Exit Sub

    ElseIf Me.ComboQ1 = vbNullString Or Me.ComboOp1 = vbNullString Then
        MsgBox "This is not a valid checking role! Complete in the first part.   ", vbCritical
        Exit Sub
        
    ElseIf Len(Me.ComboAns1) = 0 And _
        (Me.ComboOp1.value = "is equal" Or Me.ComboOp1.value = "is not equal" Or _
         Me.ComboOp1.value = "is greater than" Or Me.ComboOp1.value = "is greater than or equal" Or _
         Me.ComboOp1.value = "is less than" Or Me.ComboOp1.value = "is less than or equal") Then
        MsgBox "This is not a valid checking role! Complete in the first part. ans  ", vbCritical
        Exit Sub
        
    ElseIf (Me.OptionAnd Or Me.OptionOr) And (Len(Me.ComboQ2) = 0 Or Len(Me.ComboOp2) = 0) Then
        MsgBox "This is not a valid checking role! Complete in the second part operation.   ", vbCritical
        Exit Sub
    
    ElseIf Len(Me.ComboAns2) = 0 And _
        (Me.ComboOp2.value = "is equal" Or Me.ComboOp2.value = "is not equal" Or _
         Me.ComboOp2.value = "is greater than" Or Me.ComboOp2.value = "is greater than or equal" Or _
         Me.ComboOp2.value = "is less than" Or Me.ComboOp2.value = "is less than or equal") Then
        MsgBox "This is not a valid checking role! Complete in the second part. ans  ", vbCritical
        Exit Sub
    End If
  
    new_row = ws.Cells(Rows.count, 1).End(xlUp).Row + 1
    
    If new_row = 2 And ws.Cells(1, 1) = "" Then
        new_row = 1
    End If
    
    If Public_module.PLAN_NUMBER > 0 Then
        new_row = Public_module.PLAN_NUMBER
    End If
    
    ws.Cells(new_row, 1) = Me.ComboQ1
    ws.Cells(new_row, 2) = Me.ComboOp1
    
    If IsNumeric(Me.ComboAns1) And Len(Me.ComboAns1) > 0 Then
    
    End If
    
    If Me.ComboOp1.value = "is empty" Or Me.ComboOp1.value = "is not empty" Then
        ws.Cells(new_row, 3) = vbNullString
    Else
        If IsNumeric(Me.ComboAns1) And Len(Me.ComboAns1) > 0 Then
            ws.Cells(new_row, 3) = CSng(Me.ComboAns1)
        Else
            ws.Cells(new_row, 3) = Me.ComboAns1
        End If
    End If
    
    ws.Cells(new_row, 8) = Me.TextMassage
    
    ws.Cells(new_row, 4) = main_operation
    
    ws.Cells(new_row, 5) = Me.ComboQ2
    ws.Cells(new_row, 6) = Me.ComboOp2
    
    If Me.ComboOp2.value = "is empty" Or Me.ComboOp2.value = "is not empty" Then
        ws.Cells(new_row, 7) = vbNullString
    Else
        If IsNumeric(Me.ComboAns2) And Len(Me.ComboAns2) > 0 Then
            ws.Cells(new_row, 7) = CSng(Me.ComboAns2)
        Else
            ws.Cells(new_row, 7) = Me.ComboAns2
        End If
    End If
    
    If Me.ComboQ2.value = vbNullString Then
        ws.Cells(new_row, 4) = vbNullString
        ws.Cells(new_row, 5) = vbNullString
        ws.Cells(new_row, 6) = vbNullString
        ws.Cells(new_row, 7) = vbNullString
    End If
    
    Call remove_duplicated_plan

    If Public_module.PLAN_NUMBER > 0 Or is_loaded("plan_list_form") Then
        Unload plan_list_form
        plan_list_form.Show
    End If
    
    Public_module.PLAN_NUMBER = 0
    
    Unload plan_form
    ThisWorkbook.Save
    Application.ScreenUpdating = True
    Exit Sub

End Sub

Sub remove_duplicated_plan()
    Dim last_row  As Long
    last_row = ThisWorkbook.sheets("xlogical_checks").Cells(Rows.count, 1).End(xlUp).Row
    ThisWorkbook.sheets("xlogical_checks").Range("$A$1:$H$" & last_row).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8), header:=xlNo
End Sub

Private Sub OptionAnd_Click()
    Me.ComboQ2.Enabled = True
    Me.ComboOp2.Enabled = True
    Me.ComboAns2.Enabled = True
End Sub

Private Sub OptionOr_Click()
    Me.ComboQ2.Enabled = True
    Me.ComboOp2.Enabled = True
    Me.ComboAns2.Enabled = True
End Sub

Private Sub OptionSimple_Click()
    Me.ComboQ2.value = Null
    Me.ComboOp2.value = Null
    Me.ComboAns2.value = Null
    Me.ComboQ2.Enabled = False
    Me.ComboOp2.Enabled = False
    Me.ComboAns2.Enabled = False
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim ws As Worksheet
    Dim op1 As String
    Dim op2 As String
    
    Set ws = ThisWorkbook.sheets("xlogical_checks")
    Me.ComboOp1.AddItem ("is equal")
    Me.ComboOp1.AddItem ("is not equal")
    Me.ComboOp1.AddItem ("is empty")
    Me.ComboOp1.AddItem ("is not empty")
    Me.ComboOp1.AddItem ("is greater than")
    Me.ComboOp1.AddItem ("is greater than or equal")
    Me.ComboOp1.AddItem ("is less than")
    Me.ComboOp1.AddItem ("is less than or equal")
   
    Me.ComboOp2.AddItem ("is equal")
    Me.ComboOp2.AddItem ("is not equal")
    Me.ComboOp2.AddItem ("is empty")
    Me.ComboOp2.AddItem ("is not empty")
    Me.ComboOp2.AddItem ("is greater than")
    Me.ComboOp2.AddItem ("is greater than or equal")
    Me.ComboOp2.AddItem ("is less than")
    Me.ComboOp2.AddItem ("is less than or equal")
    Call PopulateComboBox
    If Public_module.PLAN_NUMBER > 0 Then
        Application.EnableEvents = False
        Me.ComboQ1 = ws.Cells(Public_module.PLAN_NUMBER, 1)
        Me.ComboOp1 = ws.Cells(Public_module.PLAN_NUMBER, 2)
        If ws.Cells(Public_module.PLAN_NUMBER, 3) <> vbNullString Then
            Me.ComboAns1 = ws.Cells(Public_module.PLAN_NUMBER, 3)
        End If
        
        If ws.Cells(Public_module.PLAN_NUMBER, 3) <> vbNullString Then
            Me.ComboAns1 = ws.Cells(Public_module.PLAN_NUMBER, 3)
        End If
        
        If ws.Cells(Public_module.PLAN_NUMBER, 5) <> vbNullString Then
            Me.ComboQ2 = ws.Cells(Public_module.PLAN_NUMBER, 5)
        End If
        
        If ws.Cells(Public_module.PLAN_NUMBER, 6) <> vbNullString Then
            Me.ComboOp2 = ws.Cells(Public_module.PLAN_NUMBER, 6)
        End If
        
        If ws.Cells(Public_module.PLAN_NUMBER, 7) <> vbNullString Then
            Me.ComboAns2 = ws.Cells(Public_module.PLAN_NUMBER, 7)
        End If
        

        Me.TextMassage = ws.Cells(Public_module.PLAN_NUMBER, 8)
        
        If ws.Cells(Public_module.PLAN_NUMBER, 4) = "and" Then
            Me.OptionAnd = True
        ElseIf ws.Cells(Public_module.PLAN_NUMBER, 4) = "or" Then
            Me.OptionOr = True
        Else
            Me.OptionSimple = True
        End If
        Application.EnableEvents = True
    Else
        Me.OptionSimple = True

    End If
    
    
    
End Sub

Private Sub PopulateComboBox()
    On Error Resume Next
    Dim header_arr() As Variant
    Dim filtered_arr() As String
    Dim ws As Worksheet
    
    Set ws = sheets(find_main_data)

    header_arr = ws.Range(ws.Cells(1, 1), ws.Cells(1, 1).End(xlToRight)).Value2
    
    With Application
        header_arr = .Transpose(.Transpose(header_arr))
    End With
   
    Me.ComboQ1.List = header_arr
    Me.ComboQ2.List = header_arr

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    If CloseMode = vbFormControlMenu Then
        Public_module.PLAN_NUMBER = 0
    End If
End Sub

Private Function is_loaded(form_name As String) As Boolean
    On Error Resume Next
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = form_name Then
            is_loaded = True
            Exit Function
        End If
    Next frm
    is_loaded = False
End Function

