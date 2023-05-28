Attribute VB_Name = "RAM"
Sub FindDuplicate()
    Application.ScreenUpdating = False
    Dim xWs As Worksheet
    Set xWs = Worksheets("simple")
    For m = 1 To 10444
        If Application.WorksheetFunction.CountIf(xWs.Range("AO1:AO10444"), xWs.Range("AO" & m)) > 1 Then
            xWs.Range("AP" & m).Value = True
        Else
            xWs.Range("AP" & m).Value = False
        End If
    Next m
        Application.ScreenUpdating = True
End Sub

Sub getfiles()
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFile As Object, sf
    Dim i As Integer, colFolders As New Collection, ws As Worksheet
    Set ws = ActiveSheet
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.getfolder("C:\Users\Mohammad AZIMI\Desktop\test-addins\audit")
    
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
        For Each sf In oFolder.subfolders
            colFolders.Add sf
        Next sf
    Loop
End Sub

Sub csv_import(path As String)
Dim ws As Worksheet, strFile As String
Set ws = ActiveWorkbook.Sheets("temp_sheet") 'set to current worksheet name
'strFile = "C:\Users\Mohammad AZIMI\Desktop\test-addins\audit\000bec19-7ccf-4935-8ce8-d7a12c02ba0a\audit.csv"

With ws.QueryTables.Add(Connection:="TEXT;" & path, Destination:=ws.Range("A1"))
     .TextFileParseType = xlDelimited
     .TextFileCommaDelimiter = True
     .Refresh
End With
End Sub

Sub remove_rows()
Sheets("temp_sheet").Select
    With Cells(1, 1).CurrentRegion
        .AutoFilter 1, "<>*question*"                '<~~ Filter for any instance of ""<>*question*" in column A (1)
        .Offset(1).EntireRow.Delete
        .AutoFilter
    End With
End Sub

Function add_calculation()
    
    Sheets("temp_sheet").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=(RC[-1]-RC[-2])/1000"
    
    'Find the last non-blank cell in column A(1)
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E" & CStr(lrow))
    
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=SUM(C[-1])/60"
    Range("F2").Select
    add_calculation = Range("F2")
    
End Function

Sub clear_sheet()
    Sheets("temp_sheet").Select
    Cells.Select
    Selection.QueryTable.Delete
    Selection.ClearContents
End Sub

Sub time_check()
progress_form.Show
Application.ScreenUpdating = False
Sheets.Add.Name = "temp_sheet"
Dim iCell As Range
Sheets("path_info").Select

record_count = Cells(Rows.Count, 1).End(xlUp).Row

percentage_value = Round(record_count / 100, 0)
progress_value = record_count / 270

For Each iCell In Range("A1:A" & CStr(record_count)).Cells
    progress_form.percentage.Caption = CStr(Round(iCell.Row / percentage_value, 0)) & " %"
    progress_form.bar.Width = CDec(iCell.Row / progress_value)
    DoEvents
    Call csv_import(iCell & "\audit.csv")
    Call remove_rows
    qq = add_calculation()
    Sheets("path_info").Select
    Range("C" & CStr(iCell.Row)).Select
    ActiveCell.Value = qq
    Call clear_sheet
    Sheets("path_info").Select
Next iCell
Unload progress_form
Application.DisplayAlerts = False
Sheets("temp_sheet").Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

