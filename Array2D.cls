VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Array2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' clsArray
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: This class provides extended functionality for using VBA Arrays
' YouTube video: https://youtu.be/ZYxa-Q-nSX0

Enum arrayErrors
    errColumnDoesNotExist = vbObjectError + 513
    errParameterNotArray = vbObjectError + 514
    errIncorrectColumnNumber = vbObjectError + 515
    errIncorrectRowNumber = vbObjectError + 516
    errInvalidArrayPosition = vbObjectError + 517
    errIncorrectNumberofRows = vbObjectError + 518
    errIncorrectNumberofColumns = vbObjectError + 519
    errParameterArrayEmpty = vbObjectError + 520
    errNotArray = vbObjectError + 521
    errNot2DArray = vbObjectError + 522
    errArrayNotSet = vbObjectError + 523
End Enum

Private m_data As Variant

' PROPERTIES
Public Property Get rowStart() As Variant
    Call CheckArrayIsSet("RowStart()")
    rowStart = LBound(m_data, 1)
End Property
Public Property Get rowEnd() As Variant
    Call CheckArrayIsSet("rowEnd()")
    rowEnd = UBound(m_data, 1)
End Property
Public Property Get rowCount() As Variant
    rowCount = UBound(m_data, 1) - LBound(m_data, 1) + 1
End Property
Public Property Get columnStart() As Variant
    Call CheckArrayIsSet("columnStart()")
    columnStart = LBound(m_data, 2)
End Property
Public Property Get columnEnd() As Variant
    Call CheckArrayIsSet("columnEnd()")
    columnEnd = UBound(m_data, 2)
End Property
Public Property Get columnCount() As Variant
    Call CheckArrayIsSet("columnEnd()")
    columnCount = UBound(m_data, 2) - LBound(m_data, 2) + 1
End Property

Public Property Let SetValue(row As Long, column As Long, value As Variant)
    m_data(row, column) = value
End Property


Public Property Get data() As Variant
    data = m_data
End Property

Public Property Let data(ByRef data As Variant)
   
    ' If one value then convert to an array
    If IsEmpty(data) = True Or IsArray(data) = False Then
        ReDim m_data(1 To 1, 1 To 1)
        m_data(1, 1) = data
    ElseIf IsArray(data) = True Then
        If CheckDimensions(data) <> 2 Then
           err.Raise errNot2DArray, "Let data Property" _
                , "The array to be assigned to the data property must be a two-dimensional array."
        End If
        m_data = data
    End If
    
End Property

' Set the value in the array
Public Property Let value(ByVal row As Long, ByVal column As Long, ByVal data As Variant)
    m_data(row, column) = data
End Property

' Read the value from the array
Public Property Get value(ByVal row As Long, ByVal column As Long) As Variant
    value = m_data(row, column)
End Property

' Coverts the range to an array. The main purpose of this sub is to remove the header from
' the original data. If not removing the header then use the "data" property to assign the range to the array.
Public Sub dataFromRange(ByVal rg As Range _
                        , Optional removeHeader As Boolean = False _
                        , Optional headerCount As Long = 1 _
                        , Optional removeHeaderLeft As Boolean = False _
                        , Optional headerLeftCount As Long = 1)

    
    If removeHeader = True Or removeHeaderLeft = True Then
    
        ' Get the new size for the rows based on removeHeader being true or false
        Dim rowCount As Long, rowOffset As Long
        rowCount = IIf(removeHeader = True, rg.rows.count - headerCount, rg.rows.count)
        rowOffset = IIf(removeHeader = True, headerCount, 0)
        
        ' Get the new size for the rows based on removeHeaderLeft being true or false
        Dim columnCount As Long, columnOffset As Long
        columnCount = IIf(removeHeaderLeft = True, rg.columns.count - headerLeftCount, rg.columns.count)
        columnOffset = IIf(removeHeaderLeft = True, headerLeftCount, 0)
        
        ' remove top and/or left header
        Set rg = rg.Offset(rowOffset, columnOffset).Resize(rowCount, columnCount)
        
    End If
    
    ' Copy the range data to the array
    m_data = rg.value
    

End Sub

'Private Sub UpdateStartValues()
'
'    ' row and column start will always be 1 but if this class is updated in future to accomdate 0 based arrays then
'    ' this code will still be valid.
'    rowStart = LBound(m_data, 1)
'    rowEnd = UBound(m_data, 1)
'    rowCount = UBound(m_data, 1) - LBound(m_data, 1) + 1
'    columnStart = LBound(m_data, 2)
'    columnEnd = UBound(m_data, 2)
'    m_columnCount = UBound(m_data, 2) - LBound(m_data, 2) + 1
'
'End Sub

' GET MEMBER PROCEDURES

' getColumnsFromRows
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Retrieves one or more rows and converts them to a column
' Parameters:
'           row: the row to retrieve
'           numberOfRows: [optional]the number of rows to retrieve
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Function getColumnsFromRows(ByVal row As Long, Optional ByVal numberOfRows As Long = 1) As Variant
    
    If row - 1 + numberOfRows > rowEnd Then
       err.Raise errIncorrectNumberofRows, "getRows" _
            , "The number of rows reads past the number of rows in the array. Please change the row or numberOfRows parameter values."
    End If

    Dim newColumns As Variant
    ReDim newColumns(1 To columnEnd, 1 To numberOfRows)
    
    Dim lastRow As Long: lastRow = ((row - 1) + numberOfRows)
    Dim currentRow As Long: currentRow = rowStart
    
    Dim i As Long, j As Long
    For i = row To lastRow
        For j = 1 To columnEnd
            newColumns(j, currentRow) = m_data(i, j)
        Next j
        currentRow = currentRow + 1
    Next i

    getColumnsFromRows = newColumns

End Function

' getRowsFromColumns
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Retrieves one or more columns and converts them to a row
' Parameters:
'           column: the column to retrieve
'           numberOfColumns: [optional]the number of columns to retrieve
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Function getRowsFromColumns(ByVal column As Long, Optional ByVal numberOfColumns As Long = 1) As Variant
    
    If column - 1 + numberOfColumns > columnEnd Then
       err.Raise errIncorrectNumberofColumns, "getRowFromColumns" _
            , "The number of columns reads past the number of columns in the array. Please change the column or numberOfColumns parameter values."
    End If

    Dim newRows As Variant
    ReDim newRows(columnStart To numberOfColumns, rowStart To rowEnd)
    
    Dim lastColumn As Long: lastColumn = ((column - 1) + numberOfColumns)
    Dim currentColumn As Long: currentColumn = columnStart
    
    Dim i As Long, j As Long
    For j = column To lastColumn
        For i = rowStart To rowEnd
            newRows(currentColumn, i) = m_data(i, j)
        Next i
        currentColumn = currentColumn + 1
    Next j

    getRowsFromColumns = newRows

End Function


' getRows
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Retrieves one or more rows
' Parameters:
'           row: the row to retrieve
'           numberOfRows: [optional]the number of rows to retrieve
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Function getRows(ByVal row As Long, Optional ByVal numberOfRows As Long = 1) As Variant

    If row - 1 + numberOfRows > rowCount Then
       err.Raise errIncorrectNumberofRows, "getRows" _
            , "The number of rows reads past the number of rows in the array. Please change the row or numberOfRows parameter values."
    End If
    If numberOfRows < 1 Then
       err.Raise errIncorrectNumberofRows, "getRows" _
            , "The numberOfRows must be one or greater. Please change the numberOfRows parameter value."
    End If
    

    Dim newRow As Variant
    ReDim newRow(rowStart To numberOfRows, columnStart To columnEnd)
    
    Dim lastRow As Long: lastRow = ((row - 1) + numberOfRows)
    Dim currentRow As Long: currentRow = rowStart
    
    Dim i As Long, j As Long
    For i = row To lastRow
        For j = columnStart To columnEnd
            newRow(currentRow, j) = m_data(i, j)
        Next j
        currentRow = currentRow + 1
    Next i
    
    getRows = newRow

End Function

' getRows
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Retrieves one or more rows
' Parameters:
'           column: the column to retrieve
'           numberOfColumns: [optional]the number of columns to retrieve
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Function GetColumns(ByVal column As Long, Optional ByVal numberOfColumns As Long = 1) As Variant

    If column - 1 + numberOfColumns > columnEnd Then
       err.Raise errIncorrectNumberofColumns, "getColumns" _
            , "The number of columns reads past the number of columns in the array. Please change the column or numberOfColumns parameter values."
    End If
    If column < columnStart Or column > columnEnd Then
       err.Raise errIncorrectColumnNumber, "getColumns" _
            , "The column number is invalid as it does not exist in the array."
    
    End If
    
    Dim newColumn As Variant
    ReDim newColumn(rowStart To rowEnd, columnStart To numberOfColumns)
    
    Dim lastColumn As Long: lastColumn = ((column - 1) + numberOfColumns)
    Dim currentColumn As Long: currentColumn = columnStart
    
    Dim i As Long, j As Long
    For j = column To lastColumn
        For i = rowStart To rowEnd
            newColumn(i, currentColumn) = m_data(i, j)
        Next i
        currentColumn = currentColumn + 1
    Next j
    
    GetColumns = newColumn

End Function


' WRITE MEMBER PROCEDURES

' getRows
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Writes the array to a given range
' Parameters: rg as range
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Sub writeDataToRange(ByVal rg As Range, Optional ByVal transpose As Boolean = False)
    
    If transpose = True Then
        rg.Resize(columnEnd, rowEnd).value = getRowsFromColumns(1, columnCount)
    Else
        rg.Resize(rowEnd, columnEnd).value = m_data
    End If
    
End Sub

' writeRowsToRange
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Writes rows from the array to a given range
' Parameters:
'           row: the row to write to the range
'           writeRange:  the range where the rows will be written
'           numberOfRows: the number of rows to write
'           Transpose: True to convert to columns
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Sub writeRowsToRange(ByVal writeRange As Range _
                                , ByVal row As Long _
                                , Optional numberOfRows As Long = 1 _
                                , Optional transpose As Boolean = False)
    
    If row < rowStart Or row > rowEnd Then
       err.Raise errIncorrectRowNumber, "AddRows" _
            , "The row number does not exist in the array."
    End If
    
    If transpose = True Then
        writeRange.Resize(columnEnd, numberOfRows).value = getColumnsFromRows(row, numberOfRows)
    Else
        writeRange.Resize(numberOfRows, columnEnd).value = getRows(row, numberOfRows)
    End If
    
End Sub

' writeColumnsToRange
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Writes columns from the array to a given range
' Parameters:
'           column: the column to write to the range
'           writeRange: the range where the columns will be written
'           numberOfColumns: the number of columns to write
'           Transpose: True to convert to columns
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Sub writeColumnsToRange(ByVal writeRange As Range _
                                , ByVal column As Long _
                                , Optional numberOfColumns As Long = 1 _
                                , Optional transpose As Boolean = False)
    
    If transpose = True Then
        writeRange.Resize(numberOfColumns, rowEnd).value = getRowsFromColumns(column, numberOfColumns)
    Else
        writeRange.Resize(rowEnd, numberOfColumns).value = GetColumns(column, numberOfColumns)
    End If
    
End Sub

' PRIVATE COPY PROCEDURES
' CopyArray
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Copy one array to another
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Private Function CopyArray(ByRef arrSource As Variant _
                            , ByRef arrDestination As Variant)
    
    Dim i As Long, j As Long
    For i = LBound(arrDestination, 1) To UBound(arrDestination, 1)
        For j = LBound(arrDestination, 2) To UBound(arrDestination, 2)
            arrDestination(i, j) = arrSource(i, j)
        Next j
    Next i
    
End Function

' CopyRow
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Copies the row contents from one range to another
' Parameters:
'           row: the row array where the data will be copies
'           currentRow: the row in the array that will be copies
'           newRow: the position of the row in the new array
'           numberColumns: True to convert to columns
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Private Sub CopyRow(ByRef rowData As Variant _
                            , ByVal currentRow As Long _
                            , ByVal newRow As Long)
    
    
    Dim i As Long
    For i = LBound(m_data, 2) To UBound(m_data, 2)
        rowData(newRow, i) = m_data(currentRow, i)
    Next i
    
End Sub

' Utility Procedures
Private Function CheckDimensions(ByRef arr As Variant) As Long
    If IsArray(arr) = False Then
       err.Raise errIncorrectNumberofColumns, "CheckDimensions" _
            , "The arr parameter is not a valid array."
    End If
    
    On Error GoTo eh
    
    Dim dimensions As Long
    Dim count As Long: count = 0
    
    Do
        count = count + 1
        dimensions = UBound(arr, count)
    Loop While True
    
eh:
    CheckDimensions = count - 1
End Function



' PROCEDURES FOR UPDATING THE ARRAY


' insertColumnsBlank
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Inserts a given number of blank columns
' Parameters:
'           insertColumn: the column where the new column(s) will be inserted
'           numberOfColumns:[optional] the number of columns to be inserted
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Sub insertColumnsBlank(ByVal insertColumn As Long, Optional ByVal numberOfColumns As Long = 1)
    
    ' If the insertColumn is greater than the number of columns then add new columns at the end
    If insertColumn > columnEnd + 1 Then
        insertColumn = columnEnd + 1
    End If
    
    ' Create a new array
    Dim temp As Variant
    
    ' Create the new array with the extra columns
    ReDim temp(rowStart To rowEnd, columnStart To columnEnd + numberOfColumns)
    
    ' Copy the columns before insert
    Dim i As Long, j As Long
    For j = columnStart To insertColumn - 1
        For i = rowStart To rowEnd
            temp(i, j) = m_data(i, j)
        Next i
    Next j
        
    ' Copy the columns after insert
    For j = insertColumn To columnEnd
        For i = rowStart To rowEnd
            temp(i, j + numberOfColumns) = m_data(i, j)
        Next i
    Next j
    
    m_data = temp
        
End Sub

' insertColumns
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Inserts an array of columns into the array
' Parameters:
'           insertColumn: the column where the new column(s) will be inserted
'           newColumns: The array of columns to copy to the array.
' Note: If the newColumns array has less rows than the main array then an error will be raised.
'       If the newColumns has more rows than the main array there will not be an error but extra rows will not
'       be copied.
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Sub insertColumns(ByVal insertColumn As Long, ByRef newColumns As Variant)
    
    ' Check that the newColumns array is not empty
    If IsEmpty(newColumns) Then
        err.Raise errParameterArrayEmpty, "insertColumn" _
            , "The newColumns array parameter is empty. It should be a 2D array"
    End If
    ' Check the newColumns array is 2D
    If CheckDimensions(newColumns) <> 2 Then
           err.Raise errNot2DArray, "Let data Property" _
                , "The newColumns array must be a two-dimensional array."
    End If
    ' Check that the newColumns array has the correct number of rows
    If UBound(newColumns, 1) < rowEnd Then
        err.Raise errIncorrectNumberofColumns, "insertColumn" _
            , "There are not enough rows in the newColumns array. There must be the same number as in the main Array."
    End If
    ' Check that the insertColumn is valid
    If insertColumn < 1 Then
        err.Raise errInvalidArrayPosition, "insertColumn" _
            , "The insertColumn for the array is invalid. Please use a valid column number."
    End If
    
    ' If the insertColumn is greater than the number of columns then add the new columns to the end of the array.
    If insertColumn > columnEnd + 1 Then
        insertColumn = columnEnd + 1
    End If
    
    Dim newColumnCount As Long
    newColumnCount = UBound(newColumns, 2)
    
    ' Create a new array
    Dim temp As Variant
    
    ' Create the new array with the extra columns
    ReDim temp(rowStart To rowEnd, columnStart To columnEnd + newColumnCount)
    
    ' Copy the new column data to the new array
    Dim i As Long, j As Long, currentColumn As Long
    currentColumn = insertColumn
    For j = columnStart To newColumnCount
        For i = rowStart To rowEnd
            temp(i, currentColumn) = newColumns(i, j)
        Next i
        currentColumn = currentColumn + 1
    Next j

    ' Copy the columns before insert position
    For j = columnStart To insertColumn - 1
        For i = rowStart To rowEnd
            temp(i, j) = m_data(i, j)
        Next i
    Next j
        
    ' Copy the columns after the insertColumn
    For j = insertColumn To columnEnd
        For i = rowStart To rowEnd
            temp(i, j + newColumnCount) = m_data(i, j)
        Next i
    Next j
    
    m_data = temp
        
End Sub


' insertRowsBlank
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Inserts a given number of blank rows
' Parameters:
'           insertRow: the row where the new row(s) will be inserted
'           numberOfRows:[optional] the number of rows to be inserted
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Sub insertRowsBlank(ByVal insertRow As Long, Optional ByVal numberOfRows As Long = 1)
    
    ' If the insertRow is greater than the number of rows then add new rows at the end
    If insertRow > rowEnd + 1 Then
        insertRow = rowEnd + 1
    End If
    
    ' Create a new array
    Dim temp As Variant
    
    ' Create the new array with the extra rows
    ReDim temp(rowStart To rowEnd + numberOfRows, columnStart To columnEnd)
    
    ' Copy the rows before insert
    Dim i As Long, j As Long
    For i = rowStart To insertRow - 1
        For j = columnStart To columnEnd
            temp(i, j) = m_data(i, j)
        Next j
    Next i
        
    ' Copy the rows after insert
    For i = insertRow To rowEnd
        For j = columnStart To columnEnd
            temp(i + numberOfRows, j) = m_data(i, j)
        Next j
    Next i
    
    m_data = temp
        
End Sub

' insertRows
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Inserts an array of rows into the array
' Parameters:
'           insertRow: the row where the new row(s) will be inserted
'           newRows: The array of rows to copy to the array.
' Note: If the newRows array has less columns than the main array then an error will be raised.
'       If the newRows has more columns than the main array there will not be an error but extra columns will not
'       be copied.
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Sub insertRows(ByVal insertRow As Long, ByRef newRows As Variant)
    
    ' Check that the newRows array is not empty
    If IsEmpty(newRows) Then
        err.Raise errParameterArrayEmpty, "insertRow" _
            , "The newRows array parameter is empty. It should be a 2D array"
    End If
    ' Check the newRows array is 2D
    If CheckDimensions(newRows) <> 2 Then
           err.Raise errNot2DArray, "Let data Property" _
                , "The newRows array must be a two-dimensional array."
    End If
    ' Check that the newRows array has the correct number of columns
    If UBound(newRows, 2) < columnEnd Then
        err.Raise errIncorrectNumberofRows, "insertRow" _
            , "There are not enough columnsw in the newRows array. There must be the same number as in the main Array."
    End If
    ' Check that the insertRow is
    If insertRow < 1 Then
        err.Raise errInvalidArrayPosition, "insertRow" _
            , "The insertRow for the array is invalid. Please use a valid row number."
    End If
    
    ' If the insertRow is greater than the number of rows then add the new rows to the end of the array.
    If insertRow > rowEnd + 1 Then
        insertRow = rowEnd + 1
    End If
    
    Dim newRowCount As Long
    newRowCount = UBound(newRows, 1)
    
    ' Create a new array
    Dim temp As Variant
    
    ' Create the new array with the extra rows
    ReDim temp(rowStart To rowEnd + newRowCount, columnStart To columnEnd)
    
    ' Copy the new row data to the new array
    Dim i As Long, j As Long, currentRow As Long
    currentRow = insertRow
    For i = rowStart To newRowCount
        For j = columnStart To columnEnd
            temp(currentRow, j) = newRows(i, j)
        Next j
        currentRow = currentRow + 1
    Next i

    ' Copy the rows before insert position
    For i = rowStart To insertRow - 1
        For j = columnStart To columnEnd
            temp(i, j) = m_data(i, j)
        Next j
    Next i
        
    ' Copy the rows after the insertRow
    For i = insertRow To rowEnd
        For j = columnStart To columnEnd
            temp(i + newRowCount, j) = m_data(i, j)
        Next j
    Next i
    
    m_data = temp
        
End Sub

' RemoveColumns
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Inserts an array of rows into the array
' Parameters:
'           removeColumn: the row where the removal will start from
'           numberOfColumns:[optional] the number of rows that will be removed
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Sub RemoveColumns(ByVal removeColumn As Long, Optional ByVal numberOfColumns As Long = 1)
    
    If removeColumn < 1 Or removeColumn > columnEnd Then
       err.Raise errIncorrectColumnNumber, "RemoveColumns" _
            , "The removeColumn column number for the array is invalid. Please use a valid column number."
    End If
    If removeColumn = 1 And numberOfColumns >= columnEnd Then
       err.Raise errIncorrectNumberofColumns, "RemoveColumns" _
            , "The values you have passed as parameters will delete all the columns. At least one column must be left in the array."
    End If
    
    ' If the number of columns is greater than the columns to be deleted the set the number to the max possible
    If removeColumn - 1 + numberOfColumns > columnEnd Then
        numberOfColumns = columnEnd - removeColumn + 1
    End If
    
    
    ' Create a new array
    Dim temp As Variant
    
    ' Create the new array with one less column
    ReDim temp(rowStart To rowEnd, columnStart To columnEnd - numberOfColumns)
    
    Dim i As Long, j As Long
    ' Copy the columns before deleted column
    For j = columnStart To removeColumn - 1
        For i = rowStart To rowEnd
            temp(i, j) = m_data(i, j)
        Next i
    Next j
    
    ' Copy the rows after deleted removeColumn
    For j = removeColumn + numberOfColumns To columnEnd
        For i = rowStart To rowEnd
            temp(i, j - numberOfColumns) = m_data(i, j)
        Next i
    Next j
    
    m_data = temp

End Sub


' RemoveRows
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Inserts an array of rows into the array
' Parameters:
'           removeRow: the row where the removal will start from
'           numberOfRows:[optional] the number of rows that will be removed
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Sub RemoveRows(ByVal removeRow As Long, Optional ByVal numberOfRows As Long = 1)
    
    If removeRow < 1 Or removeRow > rowEnd Then
       err.Raise errIncorrectRowNumber, "RemoveRows" _
            , "The removeRow for the array is invalid. Please use a valid row number."
    End If
    If removeRow - 1 + numberOfRows > rowEnd Then
       err.Raise errIncorrectNumberofRows, "RemoveRows" _
            , "The number of rows to delete from the removeRow is greater than the actual number of rows in the array."
    End If
    
    ' Create a new array
    Dim temp As Variant
    
    ' Create the new array with one less removeRow
    ReDim temp(rowStart To rowEnd - numberOfRows, columnStart To columnEnd)
    
    Dim i As Long, j As Long
    ' Copy the rows before deleted removeRow
    For i = rowStart To removeRow - 1
        For j = columnStart To columnEnd
            temp(i, j) = m_data(i, j)
        Next j
    Next i
    
    ' Copy the rows after deleted removeRow
    For i = removeRow + numberOfRows To rowEnd
        For j = columnStart To columnEnd
            temp(i - numberOfRows, j) = m_data(i, j)
        Next j
    Next i
    
    m_data = temp
    
End Sub


' PROCEDURES FOR SEARCHING

' IndexOf
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Finds an item in the array and returns the row
'
' Parameters:
'           searchItem: the item to search for
'           column:[optional] the column to search
'           compareType:[optional] vbBinary is case sensitive, vbTextCompare is not case sensitive
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Function IndexOf(ByVal searchItem As Variant _
                        , Optional ByVal column As Long = 1 _
                        , Optional ByVal startPosition As Long = 1 _
                        , Optional ByVal compareType As VbCompareMethod = vbTextCompare) As Long
    
    If column > columnEnd Or column < 1 Then
        err.Raise errColumnDoesNotExist _
                    , "indexOf" _
                    , "The column supplied does not exist in the array."
    End If
    If startPosition > rowEnd Then
        err.Raise errIncorrectRowNumber _
                    , "indexOf" _
                    , "The startPosition is not a valid row in the array."
    End If
    
    IndexOf = -1
    
    Dim i As Long
    For i = startPosition To rowEnd
        
        If StrComp(m_data(i, column), searchItem, compareType) = 0 Then
            ' If found then return the index
            IndexOf = i
            Exit For
        End If
        
    Next i
        
End Function

' indexofByFunction
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Finds an item in the array and returns the row.
'
' Parameters:
'           compareFunction: The class which contains the IndexOf function that is used search.
'                            The compareFunction parameter is a class that Implements iCompareFunction class.
'                            This means it has to create an IndexOf function which returns true or false.
'
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
'Public Function indexofByFunction(compareFunction As iCompareFunction, Optional ByVal startPosition As Long = 1) As Long
'
'    If startPosition > rowEnd Then
'        err.Raise errIncorrectRowNumber _
'                    , "indexofByFunction" _
'                    , "The startPosition is not a valid row in the array."
'    End If
'
'
'    indexofByFunction = -1
'
'    Dim i As Long
'    For i = startPosition To rowEnd
'
'        If compareFunction.Compare(m_data, i) Then
'            ' If found then return the index
'            indexofByFunction = i
'            Exit For
'        End If
'
'    Next i
'
'End Function


' Filter
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Filter the array based on match an item.
'              Similar to IndexOf but this returns a new array with the results. Note that it doesn't
'              change the original array.
'
' Parameters:
'           searchItem: the item to search for.
'           column:[optional] the column to search.
'           compareType:[optional] vbBinary is case sensitive, vbTextCompare is not case sensitive.
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Function Filter(ByVal searchItem As Variant _
                        , Optional ByVal column As Long = 1 _
                        , Optional ByVal startPosition As Long = 1 _
                        , Optional ByVal compareType As VbCompareMethod = vbTextCompare) As Variant
    
    If column > columnEnd Then
        err.Raise errColumnDoesNotExist _
                    , "indexOf" _
                    , "The column parameter has a number greater than the numbe of parameters"
    End If
    If startPosition > rowEnd Then
        err.Raise errIncorrectRowNumber _
                    , "Filter" _
                    , "The startPosition is not a valid row in the array."
    End If
    
    
    
    Dim filteredData As Variant
    ReDim filteredData(1 To UBound(m_data, 1), 1 To UBound(m_data, 2))
    
    Dim count As Long: count = 1
        
    Dim i As Long, newRow As Long, filterEmpty As Boolean
    newRow = 1
    filterEmpty = True
    For i = startPosition To rowEnd
        
        If StrComp(m_data(i, column), searchItem, compareType) = 0 Then
            ' IF filter is true then copy
            CopyRow filteredData, i, newRow
            newRow = newRow + 1
            filterEmpty = False
            
        End If
        
    Next i
    
    If filterEmpty = True Then
        ' No data found
        Filter = Empty
    Else
        ' Copy the filteredData array which is the same size as the original array
        ' to a smaller array
        Dim finalData As Variant
        ReDim finalData(1 To newRow - 1, 1 To UBound(filteredData, 2))
        CopyArray filteredData, finalData
    End If
    
    Filter = finalData
    
End Function

' filterByFunction
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description: Finds an item in the array and returns the row
'
' Parameters:
'           compareFunction: The class which contains the Filter function that contains the filter criteria.
'                           This compareFunction parameter is a class that Implements iCompareFunction class.
'                           This means it has to create an Filter function which returns true or false.
'
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
'Public Function filterByFunction(compareFunction As iCompareFunction, Optional ByVal startPosition As Long = 1) As Variant
'
'    If startPosition > rowEnd Then
'        err.Raise errIncorrectRowNumber _
'                    , "Filter" _
'                    , "The startPosition is not a valid row in the array."
'    End If
'
'    Dim filteredData As Variant
'    ReDim filteredData(1 To UBound(m_data, 1), 1 To UBound(m_data, 2))
'
'    Dim count As Long: count = 1
'
'    Dim i As Long, newRow As Long, filterEmpty As Boolean
'    newRow = 1
'    filterEmpty = True
'    For i = startPosition To rowEnd
'
'        If compareFunction.Compare(m_data, i) Then
'            ' if filter is true then copy
'            CopyRow filteredData, i, newRow
'            newRow = newRow + 1
'            filterEmpty = False
'
'        End If
'
'    Next i
'
'    If filterEmpty = True Then
'        ' No data found
'        filterByFunction = Empty
'    Else
'        ' Copy the filteredData array which is the same size as the original array
'        ' to a smaller array
'        Dim finalData As Variant
'        ReDim finalData(1 To newRow - 1, 1 To UBound(filteredData, 2))
'        CopyArray filteredData, finalData
'    End If
'
'    ' return the new array
'    filterByFunction = finalData
'
'End Function

' SORTING PROCEDURES
' Reverse
' https://ExcelMacroMastery.com/
' Author: Paul Kelly
' Description:  Reverses the array. If you want the array in descending order then Sort
'               and Reverse.
'
' YouTube video: https://youtu.be/ZYxa-Q-nSX0
Public Sub Reverse()
    
    Dim rows As Long, columns As Long
    rows = rowEnd
    columns = columnEnd
    
    Dim reversedData As Variant
    ReDim reversedData(1 To rows, 1 To columns)

    Dim i As Long, j As Long, row As Long
    row = 1
    For i = rows To 1 Step -1
        For j = 1 To columns
            reversedData(row, j) = m_data(i, j)
        Next j
        row = row + 1
    Next i
    
    m_data = reversedData
    
End Sub

' Requires Office 365
' Pass sort_order as -1 for descending
Public Sub Sort2(ByVal column As Long _
                    , Optional ByVal sort_order As Long = 1 _
                    , Optional ByVal byColumn As Boolean = False)
    m_data = WorksheetFunction.Sort(m_data, column, sort_order)

End Sub

Public Sub Sort(ByVal column As Long)
    QuickSort SortArray:=m_data, lngColumn:=column
End Sub

' A QuickSort algorithm for sorting the Array.
Private Sub QuickSort(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSort arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSort(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSort(SortArray, i, lngMax, lngColumn)
    
End Sub


' Checks

Public Sub CheckArrayIsSet(ByVal procedure As String)

    If IsEmpty(m_data) Then
        err.Raise errArrayNotSet, procedure _
            , "The array has not been set for the clsArray2D class. Please set the array after you create the class object."
    End If

End Sub





