Attribute VB_Name = "download_media"
Sub download_audit()
    Dim FileUrl As String
    Dim objXmlHttpReq As Object
    Dim objStream As Object
    record_count = Cells(Rows.Count, 1).End(xlUp).Row
    audit_col = column_letter("audit_URL")
    uuid_col_number = column_number("_uuid")
    
    base_path = ThisWorkbook.path
    If Dir(base_path & "\audit", vbDirectory) = "" Then
        MkDir base_path & "\audit"
    End If
    
    For Each iCell In Range(audit_col & "2:" & audit_col & CStr(record_count)).Cells
    
        FileUrl = iCell
        
        If FileUrl = "" Then
            GoTo NextIteration
        End If
        
        Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
        objXmlHttpReq.Open "GET", FileUrl, False, "azemireach", "MUsfoi^%$sja"
        objXmlHttpReq.send
        
        If objXmlHttpReq.Status = 200 Then
            Application.StatusBar = "Downloding audit files: " & iCell.Row - 1
            DoEvents
            Set objStream = CreateObject("ADODB.Stream")
            objStream.Open
            objStream.Type = 1
            
            objStream.Write objXmlHttpReq.responseBody
            Call make_folder(Cells(iCell.Row, uuid_col_number), "audit")
            
            objStream.SaveToFile ThisWorkbook.path & "\audit\" & Cells(iCell.Row, uuid_col_number) & "\audit.csv", 2
            objStream.Close
        End If
NextIteration:
    Next iCell
    Application.StatusBar = False
End Sub

Sub download_photo()
    Dim FileUrl As String
    Dim objXmlHttpReq As Object
    Dim objStream As Object
    

    
    progress_form.Show
    Application.ScreenUpdating = False
    
    record_count = Cells(Rows.Count, 1).End(xlUp).Row
    percentage_value = Round(record_count / 100, 0)
    progress_value = record_count / 270
    
    pic_col = column_letter("shelter_photo_URL")
    
    pic_name_col_number = column_number("shelter_photo_URL") - 1
    uuid_col_number = column_number("_uuid")
    base_path = ThisWorkbook.path
    If Dir(base_path & "\photo", vbDirectory) = "" Then
        MkDir base_path & "\photo"
    End If
     
     For Each iCell In Range(pic_col & "2:" & pic_col & CStr(record_count)).Cells
         FileUrl = iCell
         
        progress_form.percentage.Caption = CStr(Round(iCell.Row / percentage_value, 0)) & " %"
        progress_form.bar.Width = CDec(iCell.Row / progress_value)
        DoEvents
         
        If FileUrl = "" Or Cells(iCell.Row, pic_name_col_number) = "" Then
            GoTo NextIteration
        End If
         
         Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
         objXmlHttpReq.Open "GET", FileUrl, False, "azemireach", "MUsfoi^%$sja"
         objXmlHttpReq.send
    
         If objXmlHttpReq.Status = 200 Then
            Application.StatusBar = "Downloading photos: " & iCell.Row - 1
            DoEvents
              Set objStream = CreateObject("ADODB.Stream")
              objStream.Open
              objStream.Type = 1
              objStream.Write objXmlHttpReq.responseBody
              
              Call make_folder(Cells(iCell.Row, uuid_col_number), "photo")
              
              objStream.SaveToFile ThisWorkbook.path & "\photo\" & Cells(iCell.Row, uuid_col_number) & "\" & Cells(iCell.Row, pic_name_col_number), 2
              objStream.Close
         End If
         
NextIteration:
    Next iCell
    Unload progress_form

    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub



Public Function column_letter(column_value As String)
Dim colNum As Integer
Dim vArr
worksheetName = ActiveSheet.Name
colNum = WorksheetFunction.Match(column_value, ActiveWorkbook.Sheets(worksheetName).Range("1:1"), 0)
vArr = Split(Cells(1, colNum).Address(True, False), "$")
col_letter = vArr(0)
column_letter = col_letter
End Function

Public Function column_number(column_value As String)
Dim colNum As Integer
worksheetName = ActiveSheet.Name
colNum = WorksheetFunction.Match(column_value, ActiveWorkbook.Sheets(worksheetName).Range("1:1"), 0)
column_number = colNum
End Function

Public Sub make_folder(uuid As String, folder_name As String)
base_path = ThisWorkbook.path
If Dir(base_path & "\" & folder_name & "\" & uuid, vbDirectory) = "" Then
    MkDir base_path & "\" & folder_name & "\" & uuid
End If
End Sub



