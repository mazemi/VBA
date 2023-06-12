Attribute VB_Name = "Public_module"

Public Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Public Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function

Public Function column_number(column_value As String) As Long
    Dim colNum As Long
    Dim worksheetName As String
    
    worksheetName = ActiveSheet.Name
    
    colNum = Application.Match(column_value, ActiveWorkbook.Sheets(worksheetName).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        column_number = colNum
    Else
        column_number = 0
    End If
    
End Function

Public Function column_letter(column_value As String) As String
    On Error Resume Next
    Dim colNum As Long
    Dim vArr
    worksheetName = ActiveSheet.Name

    colNum = Application.Match(column_value, ActiveWorkbook.Sheets(worksheetName).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        column_letter = Replace(Cells(1, colNum).Address(False, False), "1", "")
    Else
        column_letter = "" '
    End If
End Function

Public Sub create_sheet(sheet_name_base As String, new_sheet_name As String)
    
    Sheets.Add(After:=Sheets(sheet_name_base)).Name = new_sheet_name
    
End Sub

Sub get_files()
    On Error GoTo errHandler:
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object, sf
    Dim i As Integer, colFolders As New Collection, ws As Worksheet
    Set ws = ActiveSheet
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(ThisWorkbook.path & "\audit")
    
    colFolders.Add oFolder          'start with this folder
    Do While colFolders.Count > 0      'process all folders
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
    
        For Each oFile In oFolder.Files
            ws.Cells(i + 1, 1) = oFolder.path
            ws.Cells(i + 1, 2) = oFile.Name
            i = i + 1
        Next oFile

        'add any subfolders to the collection for processing
        For Each sf In oFolder.SubFolders
            colFolders.Add sf
        Next sf
    Loop
    
Exit Sub
errHandler:
    MsgBox "There is an issue, please make suere the audit files are existed!", vbCritical
End Sub
