Attribute VB_Name = "Public_module"
Global Data_sheet As String
Global Sample_sheet As String
Global Data_strata As String
Global Sample_strata As String
Global Sample_pop As String
Global ana As Integer
Global cancel_proc As Boolean


Public Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Function col_number(ColName As String)
    col_number = Range(ColName & 1).column
End Function

Public Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")
End Function

Public Function column_number(column_value As String) As Long
    On Error Resume Next
    Dim colNum As Long
    Dim worksheetName As String
    
    worksheetName = ActiveSheet.Name
    
    colNum = Application.Match(column_value, ActiveWorkbook.sheets(worksheetName).Range("1:1"), 0)
    
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

    colNum = Application.Match(column_value, ActiveWorkbook.sheets(worksheetName).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        column_letter = Replace(Cells(1, colNum).Address(False, False), "1", "")
    Else
        column_letter = "" '
    End If
End Function

Public Function gen_column_number(column_value As String, sheet_name As String) As Long
    On Error Resume Next
    Dim colNum As Long
    Dim worksheetName As String

    colNum = Application.Match(column_value, sheets(sheet_name).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        gen_column_number = colNum
    Else
        gen_column_number = 0
    End If
    
End Function
Public Function gen_column_letter(column_value As String, sheet_name As String) As String
    On Error Resume Next
    Dim colNum As Long
    Dim vArr

    colNum = Application.Match(column_value, sheets(sheet_name).Range("1:1"), 0)
    
    If Not IsError(colNum) Then
        gen_column_letter = Replace(sheets(sheet_name).Cells(1, colNum).Address(False, False), "1", "")
    Else
        gen_column_letter = ""
    End If
End Function
Public Sub create_sheet(sheet_name_base As String, new_sheet_name As String)
    sheets.Add(After:=sheets(sheet_name_base)).Name = new_sheet_name
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
    Do While colFolders.count > 0      'process all folders
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

Function UnmatchedElements(array1 As Variant, array2 As Variant, check_both As Boolean) As Collection
    Dim arr1() As Variant
    Dim arr2() As Variant
    Dim unmatched As New Collection
    Dim i As Long
        
    With Application
        arr1 = .transpose(array1)
    End With
    
    With Application
        arr2 = .transpose(array2)
    End With
    
    
    ' Find elements in arr1 that are not in arr2
    For i = LBound(arr1) To UBound(arr1)
        If Not IsInArray(arr1(i), arr2) Then
            unmatched.Add arr1(i)
        End If
    Next i
    
    If check_both Then
        ' Find elements in arr2 that are not in arr1
        For i = LBound(arr2) To UBound(arr2)
            If Not IsInArray(arr2(i), arr1) Then
                unmatched.Add arr2(i)
            End If
        Next i
    End If
    
    ' Print the unmatched elements
'    For i = 1 To unmatched.Count
'        Debug.Print unmatched(i)
'    Next i
    
   Set UnmatchedElements = unmatched
    
End Function

Function IsInArray(val As Variant, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If val = arr(i) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Sub clear_filter(ws As Worksheet)

    Dim filtered_col As Long

    If ws.FilterMode Then
        With ws.AutoFilter
            For filtered_col = 1 To .Filters.count
                If .Filters(filtered_col).On Then
                    ws.AutoFilter.Sort.SortFields.Clear
                    ws.ShowAllData
                End If
            Next filtered_col
        End With
        ws.UsedRange.AutoFilter
    End If
    
End Sub

Sub GetUniqueAndCount()
    Dim d As Object, c As Range, k, tmp As String

    Set d = CreateObject("scripting.dictionary")
    For Each c In Selection
        tmp = Trim(c.value)
        If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
    Next c

    For Each k In d.Keys
        Debug.Print k, d(k)
    Next k

End Sub
Function sheet_list() As Collection
    'This function returns a collection of worksheet names in the workbook
    Dim sheets As Collection
    Set sheets = New Collection
    Dim ws As Worksheet 'Use a worksheet variable instead of an index
    For Each ws In ThisWorkbook.Worksheets 'Loop through each worksheet in the workbook
        sheets.Add ws.Name 'Add the worksheet name to the collection
    Next ws
    Set sheet_list = sheets
End Function

Public Function Contains(col As Collection, key As Variant) As Boolean
Dim obj As Variant
On Error GoTo err
    Contains = True
    obj = col(key)
    Exit Function
err:

    Contains = False
End Function


Function unique_values(rng As Range) As Collection
    Dim d As Object, c As Range, h, tmp As String
    Dim unique_collection As New Collection
    
    Set d = CreateObject("scripting.dictionary")
    For Each c In rng
        tmp = Trim(c.value)
        If Len(tmp) > 0 Then d(tmp) = d(tmp) + 1
    Next c

    For Each h In d.Keys
'        Debug.Print h
         unique_collection.Add CStr(h)
    Next h
    Set unique_values = unique_collection
End Function

