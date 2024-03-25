VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} plan_list_form 
   Caption         =   "Cleaning Plan"
   ClientHeight    =   5532
   ClientLeft      =   -288
   ClientTop       =   -1338
   ClientWidth     =   10824
   OleObjectBlob   =   "plan_list_form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "plan_list_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandDelete_Click()
    If Me.ListPlan.List(0) = "NO CEALNING PLAN!" Then
        Exit Sub
    End If
   
    n = Me.ListPlan.ListCount
    Dim i As Long
    For i = 0 To n - 1
        
        If Me.ListPlan.Selected(i) Then
            Me.ListPlan.RemoveItem (i)
            Call delete_row(i + 1)
        End If

    Next
    Me.ListPlan.Clear
    Call UserForm_Initialize
    ThisWorkbook.Save
End Sub

Private Sub CommandDeleteAll_Click()
    
    Dim answer As Integer
    
    If ThisWorkbook.sheets("xlogical_checks").Range("A1") = vbNullString Then Exit Sub
    
    answer = MsgBox("All the cleaning roles will be removed." & vbCrLf & _
                    "Do you want to Continue?", vbQuestion + vbYesNo)
    
    If answer = vbYes Then

        ThisWorkbook.sheets("xlogical_checks").Cells.Clear
        Me.ListPlan.Clear
        Call UserForm_Initialize
    End If

End Sub



Private Sub ListPlan_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    n = Me.ListPlan.ListCount
    Dim i As Long
    For i = 0 To n - 1
        If Me.ListPlan.Selected(i) Then
            Call single_check(i + 1)
        End If
    Next
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.sheets("xlogical_checks")
    
    last_row = ws.Cells(Rows.count, 1).End(xlUp).Row
     
    If ws.Cells(1, 1) <> "" Then
    
        Dim str As String
        For i = 1 To last_row
            str = "If ( " & ws.Cells(i, 1) & " " & ws.Cells(i, 2) & " " & ws.Cells(i, 3) & " " & _
                ws.Cells(i, 4) & " " & ws.Cells(i, 5) & " " & ws.Cells(i, 6) & " " & _
                ws.Cells(i, 7) & " ) -> flag: " & ws.Cells(i, 8)
            str = Replace(str, "   ", " ")
            str = Replace(str, "  ", " ")
            Me.ListPlan.AddItem str, i - 1
        Next
    Else
        Me.ListPlan.AddItem "NO CEALNING PLAN!", 0
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub CommandEdit_Click()
    On Error Resume Next

    If Me.ListPlan.List(0) = "NO CEALNING PLAN!" Then
        Exit Sub
    End If
    
    n = Me.ListPlan.ListCount
    Dim i As Long
    For i = 0 To n - 1
        If Me.ListPlan.Selected(i) Then
            Public_module.PLAN_NUMBER = i + 1
            plan_form.Show
            Unload plan_list_form
        End If
    Next
End Sub

Sub delete_row(n As Long)
    On Error Resume Next
    ThisWorkbook.sheets("xlogical_checks").Rows(n).EntireRow.Delete
End Sub


