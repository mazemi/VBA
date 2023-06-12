VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} settingFrm 
   Caption         =   "Setting"
   ClientHeight    =   7308
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   11136
   OleObjectBlob   =   "settingFrm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "settingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandAdd_Click()
    If Me.ComboDataColumns.Value = "" Then
        Exit Sub
    End If
    If Me.ListBoxLog.ListCount < 5 Then
        Me.ListBoxLog.AddItem Me.ComboDataColumns.Value
    End If
End Sub

Private Sub CommandClear_Click()
    Me.ListBoxLog.Clear
End Sub

Private Sub CommandSave_Click()
    Dim koboLog As String
    koboLog = log_cols()

    ' save to registry:
    user = SaveRegistrySetting("ramSetting", "koboUserReg", Me.TextUser.Value)
    Password = SaveRegistrySetting("ramSetting", "koboPasswordReg", Me.TextPassword.Value)
    audit = SaveRegistrySetting("ramSetting", "koboAuditReg", Me.ComboAudit.Value)
    photo = SaveRegistrySetting("ramSetting", "koboPhotoReg", Me.ComboPhoto.Value)
    tools = SaveRegistrySetting("ramSetting", "koboToolsReg", Me.TextTools.Value)
    logs = SaveRegistrySetting("ramSetting", "koboLogReg", koboLog)
    
    Unload Me
    
End Sub

Private Sub CommandTools_Click()
    Dim objFSO As New FileSystemObject
    Set myFile = Application.FileDialog(msoFileDialogOpen)
    With myFile
        .Title = "Choose File"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Exit Sub
        End If
        FileSelected = .SelectedItems(1)
    End With
    Me.TextTools = FileSelected
End Sub

Private Sub UserForm_Initialize()
    Me.TextUser.Value = GetRegistrySetting("ramSetting", "koboUserReg")
    Me.TextPassword.Value = GetRegistrySetting("ramSetting", "koboPasswordReg")
    Me.ComboAudit.Value = GetRegistrySetting("ramSetting", "koboAuditReg")
    Me.ComboPhoto.Value = GetRegistrySetting("ramSetting", "koboPhotoReg")
    Me.TextTools.Value = GetRegistrySetting("ramSetting", "koboToolsReg")
    logs = GetRegistrySetting("ramSetting", "koboLogReg")
    
'    MsgBox logs
    Dim SubStringArr() As String
    Dim SrcString As String
    
    If logs <> "" Then
        SrcString = logs
        SubStringArr = Split(SrcString, ",")
        Me.ListBoxLog.List = SubStringArr
    End If
    
    Call PopulateComboBoxEdited
    
End Sub

Private Sub PopulateComboBoxEdited()

    Dim header_arr() As Variant
    Dim filtered_arr() As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.ActiveSheet  'change the sheet name accordingly

    header_arr = ws.Range(Cells(1, 1), Cells(1, 1).End(xlToRight)).Value2
    
    With Application
        header_arr = .Transpose(.Transpose(header_arr))
    End With

    filtered_arr = Filter(header_arr, "URL", True, vbTextCompare)
    
    Me.ComboAudit.List = filtered_arr
    Me.ComboPhoto.List = filtered_arr
    Me.ComboDataColumns.List = header_arr
    
End Sub

Function log_cols() As String
    Dim Size As Integer
'    Dim koboLog As String
    Size = Me.ListBoxLog.ListCount - 1
    If Size >= 0 Then
        ReDim ListBoxContents(0 To Size) As String
        Dim i As Integer
        
        For i = 0 To Size
            ListBoxContents(i) = Me.ListBoxLog.List(i)
        Next i
      
        log_cols = Join(ListBoxContents, ",")
    Else
        log_cols = ""
    End If
    
End Function
