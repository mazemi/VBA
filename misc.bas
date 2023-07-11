Attribute VB_Name = "misc"
Sub compare_strata()
    Dim r1 As Range
    Dim r2 As Range
    Dim res As Boolean
    Dim d As Variant
    
    a = Cells(rows.count, 1).End(xlUp).row
    b = Cells(rows.count, 2).End(xlUp).row
    
    ' Debug.Print a, b

    Dim col1 As New Collection
    Dim col2 As New Collection
    
    Set r1 = sheets("inpro").Range("A2:A" & a)
    Set r2 = sheets("inpro").Range("B2:B" & b)
    
    For Each i In r1
        col1.Add CStr(i)
    Next
    
    For Each j In r2
        col2.Add CStr(j)
    Next
    
    For Each d In col2
        res = HasKey(col1, CStr(d))
        If Not res Then
            Debug.Print d
            
        End If
    Next
    Debug.Print "done!"
End Sub

Function HasKey(coll As Collection, strKey As String) As Boolean
    Dim var As Variant
    On Error Resume Next
    var = coll(strKey)
    HasKey = (err.Number = 0)
    err.Clear
End Function

Public Function InCollection(col As Collection, key As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long

  InCollection = False
  Set var = Nothing

  err.Clear
  On Error Resume Next
    var = col.Item(key)
    errNumber = CLng(err.Number)
  On Error GoTo 0

  '5 is not in, 0 and 438 represent incollection
  If errNumber = 5 Then ' it is 5 if not in collection
    InCollection = False
  Else
    InCollection = True
  End If

End Function

Function SetDifference(rng1 As Range, rng2 As Range) As Range
'    On Error Resume Next
    
    If Intersect(rng1, rng2) Is Nothing Then
        'if there is no common area then we will set both areas as result
        Set SetDifference = Union(rng1, rng2)
        'alternatively
        'set SetDifference = Nothing
        Exit Function
    End If
    
'    On Error GoTo 0
    Dim aCell As Range
    For Each aCell In rng1
        Dim Result As Range
        If Application.Intersect(aCell, rng2) Is Nothing Then
            If Result Is Nothing Then
                Set Result = aCell
            Else
                Set Result = Union(Result, aCell)
            End If
        End If
    Next aCell
    Set SetDifference = Result

End Function

Sub show_last()

    Dim header_arr() As Variant
    
    last_row = sheets("RAM").Cells(rows.count, 1).End(xlUp).row
    last_row2 = sheets("RAM").UsedRange.rows(ActiveSheet.UsedRange.rows.count).row
    last_col = sheets("RAM").Cells(1, columns.count).End(xlToLeft).column
    
    ' below needs to be improved
    header_arr = sheets("RAM").Range(Cells(1, 1), Cells(1, 1).End(xlToRight)).Value2
    Debug.Print last_row, last_row2, last_col, LBound(header_arr), UBound(header_arr)
    
End Sub

Sub Extractor()
    t = Timer
    Dim i As Long, j As Long
    Dim arr() As String
    Dim LastRow As Long
    Dim endRow As Long
    
    LastRow = Cells(rows.count, "B").End(xlUp).row
    
    ' count total number of employees
    For i = 1 To LastRow
        arr = Split(Cells(i, 2), " ")
        endRow = endRow + (UBound(arr) - LBound(arr) + 1)
    Next i

    ' write cells, begining from last
    For i = LastRow To 1 Step -1
        arr = Split(Cells(i, 2), " ")
        For j = LBound(arr) To UBound(arr)
            Cells(endRow, 1) = Cells(i, 1)
            Cells(endRow, 2) = arr(j)
            endRow = endRow - 1
        Next j
    Next i
    MsgBox Timer - t
End Sub

Sub Extractor_one()
    t = Timer
    Dim i As Long, j As Long
    Dim arr() As String
    Dim LastRow As Long
    Dim endRow As Long
    
    LastRow = Cells(rows.count, "A").End(xlUp).row
    
    ' count total number of employees
    For i = 1 To LastRow
        arr = Split(Cells(i, 1), " ")
        endRow = endRow + (UBound(arr) - LBound(arr) + 1)
    Next i

    ' write cells, begining from last
    For i = LastRow To 1 Step -1
        arr = Split(Cells(i, 1), " ")
        For j = LBound(arr) To UBound(arr)
            Cells(endRow, 6) = arr(j)
            endRow = endRow - 1
        Next j
    Next i
    MsgBox Timer - t
End Sub

Sub Extractor_three()
    t = Timer
    Dim i As Long, j As Long
    Dim arr() As String
    Dim LastRow As Long
    Dim endRow As Long
    
    LastRow = Cells(rows.count, "B").End(xlUp).row
    
    ' count total number of employees
    For i = 1 To LastRow
        arr = Split(Cells(i, 2), " ")
        endRow = endRow + (UBound(arr) - LBound(arr) + 1)
    Next i

    ' write cells, begining from last
    For i = LastRow To 1 Step -1
        arr = Split(Cells(i, 2), " ")
        For j = LBound(arr) To UBound(arr)
            Cells(endRow, 5) = Cells(i, 1)
            Cells(endRow, 6) = arr(j)
            Cells(endRow, 7) = Cells(i, 3)
            endRow = endRow - 1
        Next j
    Next i
    MsgBox Timer - t
End Sub
