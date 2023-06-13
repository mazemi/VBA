VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} settingFrm 
   Caption         =   "Setting"
   ClientHeight    =   6450
   ClientLeft      =   84
   ClientTop       =   390
   ClientWidth     =   10644
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
'    tools = SaveRegistrySetting("ramSetting", "koboToolsReg", Me.TextTools.Value)
    logs = SaveRegistrySetting("ramSetting", "koboLogReg", koboLog)
    
    Unload Me
    
End Sub



Private Sub UserForm_Initialize()
    Me.TextUser.Value = GetRegistrySetting("ramSetting", "koboUserReg")
    Me.TextPassword.Value = GetRegistrySetting("ramSetting", "koboPasswordReg")
    Me.ComboAudit.Value = GetRegistrySetting("ramSetting", "koboAuditReg")
    Me.ComboPhoto.Value = GetRegistrySetting("ramSetting", "koboPhotoReg")
'    Me.TextTools.Value = GetRegistrySetting("ramSetting", "koboToolsReg")
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

Private Sub CommandTools_Click()
    Dim objFSO As New FileSystemObject
    Dim FileSelected As String
    Set myFile = Application.FileDialog(msoFileDialogOpen)
    With myFile
        .Title = "Choose File"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Exit Sub
        End If
        FileSelected = .SelectedItems(1)
    End With
    
    Call import_survey(FileSelected)
    Call import_choices(FileSelected)
    
    ' hide sheets:
    Worksheets("survey").Visible = False
    Worksheets("choices").Visible = False
    
    MsgBox "KOBO tools has beed integared.   ", vbInformation

End Sub

Sub import_survey(tools_path As String)
'    On Error Resume Next
    Application.DisplayAlerts = False

    Application.ScreenUpdating = False
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim main_ws As Worksheet
    Set main_ws = ActiveWorkbook.ActiveSheet

    'check if log_book sheet exist
    If WorksheetExists("survey") <> True Then
        Call create_sheet(main_ws.Name, "survey")
    Else
        ' clear survey sheet
        Sheets("survey").Cells.Clear
    End If

    Set ImportWorkbook = Workbooks.Open(Filename:=tools_path)
    ImportWorkbook.Worksheets("survey").UsedRange.Copy
    ThisWorkbook.Worksheets("survey").Range("A1").PasteSpecial _
    Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    ImportWorkbook.Close

    ' trime all three columns:
    Dim rng As Range
        For i = 1 To 3
        Set rng = Columns(i)
        rng.Value = Application.Trim(rng)
    Next i

    Call deleteIrrelevantColumns("survey")

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub import_choices(tools_path As String)
'    On Error Resume Next
    Application.DisplayAlerts = False

    Application.ScreenUpdating = False
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim main_ws As Worksheet
    Set main_ws = ActiveWorkbook.ActiveSheet

    'check if temp_choices sheet exist
    If WorksheetExists("choices") <> True Then
        Call create_sheet(main_ws.Name, "choices")
    Else
        ' clear chioces sheet
        Sheets("choices").Cells.Clear
    End If

    Set ImportWorkbook = Workbooks.Open(Filename:=tools_path)

    ImportWorkbook.Worksheets("choices").UsedRange.Copy
    ThisWorkbook.Worksheets("choices").Range("A1").PasteSpecial _
    Paste:=xlPasteValues, SkipBlanks:=False

    ImportWorkbook.Close

    Call deleteIrrelevantColumns("choices")

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub deleteIrrelevantColumns(sheet_name As String)
    Dim keepColumn As Boolean
    Dim currentColumn As Integer
    Dim columnHeading As String
    Dim temp_ws As Worksheet
    Set temp_ws = Worksheets(sheet_name)

    currentColumn = 1

    While currentColumn <= temp_ws.UsedRange.Columns.Count
        columnHeading = temp_ws.UsedRange.Cells(1, currentColumn).Value

        'CHECK WHETHER TO KEEP THE COLUMN
        keepColumn = False
        If columnHeading = "list_name" Then keepColumn = True
        If columnHeading = "type" Then keepColumn = True
        If columnHeading = "name" Then keepColumn = True
        If columnHeading = "label::English" Then keepColumn = True

        If keepColumn Then
            'IF YES THEN SKIP TO THE NEXT COLUMN,
            currentColumn = currentColumn + 1
        Else
            'IF NO DELETE THE COLUMN
            temp_ws.Columns(currentColumn).Delete
        End If

        'LASTLY AN ESCAPE IN CASE THE SHEET HAS NO COLUMNS LEFT
        If (temp_ws.UsedRange.Address = "$A$1") And (temp_ws.Range("$A$1").Text = "") Then Exit Sub
    Wend

End Sub

