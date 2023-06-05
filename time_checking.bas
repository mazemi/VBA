Attribute VB_Name = "time_checking"
Sub find_duplicate()
    Application.ScreenUpdating = False
    Dim xWs As Worksheet
    Set xWs = Worksheets("simple")
    For m = 1 To 10444
        If Application.WorksheetFunction.CountIf(xWs.Range("AO1:AO10444"), xWs.Range("AO" & m)) > 1 Then
            xWs.Range("AP" & m).value = True
        Else
            xWs.Range("AP" & m).value = False
        End If
    Next m
        Application.ScreenUpdating = True
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

Private Sub csv_import(path As String)
On Error Resume Next
Dim ws As Worksheet, strFile As String
Set ws = ActiveWorkbook.Sheets("temp_sheet") 'set to current worksheet name

With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A1"))
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .Refresh
End With
End Sub

Private Sub remove_rows()
On Error Resume Next
Sheets("temp_sheet").Select
    With Cells(1, 1).CurrentRegion
        .AutoFilter 1, "<>*question*"                '<~~ Filter for any instance of ""<>*question*" in column A (1)
        .Offset(1).EntireRow.Delete
        .AutoFilter
    End With
End Sub

Private Function add_calculation()
    On Error Resume Next
    Sheets("temp_sheet").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]-RC[-2])/1000"
    
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & CStr(lrow))
    
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-1])/60"
    Range("F2").Select
    add_calculation = Range("F2")
    
End Function

Private Sub clear_sheet()
    On Error Resume Next
    Sheets("temp_sheet").Select
    Cells.Select
    Selection.QueryTable.Delete
    Selection.ClearContents
End Sub

Sub time_check()
Call check_uuid
progress_form.Show
Application.ScreenUpdating = False
main_sheet = ActiveSheet.Name
Sheets.Add.Name = "temp_sheet"
Dim iCell As Range

new_col = Cells(1, Columns.Count).End(xlToLeft).Column + 1

base_path = ThisWorkbook.path & "\audit\"
Sheets(main_sheet).Select

uuid_col = column_letter("_uuid")

record_count = Cells(Rows.Count, 1).End(xlUp).Row
percentage_value = Round(record_count / 100, 0)
progress_value = record_count / 270

For Each iCell In Range(uuid_col & "2:" & uuid_col & CStr(record_count)).Cells
    progress_form.percentage.Caption = CStr(Round(iCell.Row / percentage_value, 0)) & " %"
    progress_form.bar.Width = CDec(iCell.Row / progress_value)
    DoEvents
    Call csv_import(base_path & iCell & "\audit.csv")
    Call remove_rows
    
    Duration = add_calculation()
    Sheets(main_sheet).Select
    Range("B" & CStr(iCell.Row)).value = Duration
'    MsgBox iCell.Row
    Call clear_sheet
    Sheets(main_sheet).Select
Next iCell
Unload progress_form
Application.DisplayAlerts = False
Sheets("temp_sheet").Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Public Sub last_col_number()
last_col = Cells(1, Columns.Count).End(xlToLeft).Column
MsgBox last_col
End Sub


Sub check_uuid()
On Error GoTo errHandler:

With Sheets(ActiveSheet.Name)
    col = WorksheetFunction.Match("_uuid", .Rows(1), 0)
End With
'MsgBox col

Exit Sub
errHandler:
End
End Sub

