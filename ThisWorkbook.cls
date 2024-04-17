VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    On Error Resume Next
    Call create_log_shortcut
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    Call delete_log_shortcut
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
End Sub

