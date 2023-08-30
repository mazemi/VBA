VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} plan_form 
   Caption         =   "Logical Data Cleaning "
   ClientHeight    =   3300
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   7698
   OleObjectBlob   =   "plan_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "plan_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommanSave_Click()
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.sheets("logical_checks")
    
    If Me.ComboQ1 = "" Or Me.TextAnswer1 = "" Or Me.TextMassage = "" Then
        MsgBox "The logical check is empty!   ", vbCritical
        Exit Sub
    End If
    
    Dim str As String
    str = Me.ComboQ1 & Me.TextAnswer1 & Me.OptionAnd & Me.OptionOr & Me.ComboQ2 & _
          Me.TextAnswer2 & Me.TextMassage

    new_row = ws.Cells(rows.count, 1).End(xlUp).row + 1
    
    If new_row = 2 And ws.Cells(1, 1) = "" Then
        new_row = 1
    End If
    
    If Public_module.PLAN_NUMBER > 0 Then
        new_row = Public_module.PLAN_NUMBER
    End If
    
    ws.Cells(new_row, 1) = Me.ComboQ1
    ws.Cells(new_row, 2) = Me.TextAnswer1
    ws.Cells(new_row, 6) = Me.TextMassage
    
    If Me.OptionAnd Then
        op = "and"
    Else
        op = "or"
    End If
    
    If Me.ComboQ2 <> vbNullString And Me.TextAnswer2 <> vbNullString And (Me.OptionAnd Or Me.OptionOr) Then
        ws.Cells(new_row, 3) = op
        ws.Cells(new_row, 4) = Me.ComboQ2
        ws.Cells(new_row, 5) = Me.TextAnswer2
    Else
        ws.Cells(new_row, 3) = vbNullString
        ws.Cells(new_row, 4) = vbNullString
        ws.Cells(new_row, 5) = vbNullString
    
    End If
    
    Public_module.PLAN_STRING = Me.ComboQ1 & Me.TextAnswer1 & Me.OptionAnd & Me.OptionOr & Me.ComboQ2 & _
                  Me.TextAnswer2 & Me.TextMassage
    
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
    last_row = ThisWorkbook.sheets("logical_checks").Cells(rows.count, 1).End(xlUp).row
    ThisWorkbook.sheets("logical_checks").Range("$A$1:$F$" & last_row).RemoveDuplicates columns:=Array(1, 2, 3, 4, 5, 6), Header:=xlNo
End Sub

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
       
    Set ws = ThisWorkbook.sheets("logical_checks")
    
    If Public_module.PLAN_NUMBER > 0 Then
        
        Me.CommanSave.Caption = "Save"
        
        Me.ComboQ1 = ws.Cells(Public_module.PLAN_NUMBER, 1)
        Me.TextAnswer1 = ws.Cells(Public_module.PLAN_NUMBER, 2)
        If ws.Cells(Public_module.PLAN_NUMBER, 3) = "and" Then
            Me.OptionAnd = True
            Me.OptionOr = False
        ElseIf ws.Cells(Public_module.PLAN_NUMBER, 3) = "or" Then
            Me.OptionAnd = False
            Me.OptionOr = True
        End If
        Me.ComboQ2 = ws.Cells(Public_module.PLAN_NUMBER, 4)
        Me.TextAnswer2 = ws.Cells(Public_module.PLAN_NUMBER, 5)
        Me.TextMassage = ws.Cells(Public_module.PLAN_NUMBER, 6)
    End If
    Call PopulateComboBox
   
End Sub

Private Sub PopulateComboBox()

    Dim header_arr() As Variant
    Dim filtered_arr() As String
    Dim ws As Worksheet
    
    Set ws = sheets(find_main_data)

    header_arr = ws.Range(ws.Cells(1, 1), ws.Cells(1, 1).End(xlToRight)).Value2
    
    With Application
        header_arr = .transpose(.transpose(header_arr))
    End With
   
    Me.ComboQ1.List = header_arr
    Me.ComboQ2.List = header_arr

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  
    If CloseMode = vbFormControlMenu Then
        Public_module.PLAN_NUMBER = 0
    End If
End Sub

Private Function is_loaded(form_name As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = form_name Then
            is_loaded = True
            Exit Function
        End If
    Next frm
    is_loaded = False
End Function

