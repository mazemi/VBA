Attribute VB_Name = "download_media"
Sub kobo_account_frm()
    koboForm.Show
End Sub

Sub kobo_media_url_frm()
    UrlForm.Show
End Sub

Sub download_audit()
On Error GoTo errHandler:
    Dim FileUrl As String
    Dim audit_url As String
    Dim objXmlHttpReq As Object
    Dim objStream As Object
    record_count = Cells(Rows.Count, 1).End(xlUp).Row
    
    audit_url = ""
    audit_url = GetRegistrySetting("ramSetting", "koboAuditReg")
    
    If audit_url = "" Then
        MsgBox "Please check the audit URL column!      ", vbCritical
        Exit Sub
    End If
    
    audit_col = column_letter(audit_url)
    uuid_col_number = column_number("_uuid")
    
    base_path = ThisWorkbook.path
    If Dir(base_path & "\audit", vbDirectory) = "" Then
        MkDir base_path & "\audit"
    End If
    
    UserName = ""
    Password = ""
    UserName = GetRegistrySetting("ramSetting", "koboUserReg")
    Password = GetRegistrySetting("ramSetting", "koboPasswordReg")
    
    If UserName = "" Or Password = "" Then
        MsgBox "Please check your KOBO account info!      ", vbCritical
        Exit Sub
    End If
    err_counter = 0
    For Each iCell In Range(audit_col & "2:" & audit_col & CStr(record_count)).Cells
    
        FileUrl = iCell
        
        If FileUrl = "" Then
            GoTo NextIteration
        End If
        
        strFileName = ThisWorkbook.path & "\audit\" & Cells(iCell.Row, uuid_col_number) & "\audit.csv"
        strFileExists = Dir(strFileName)
        
        If strFileExists = "" Then
            ' MsgBox "The selected file doesn't exist"
            Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
            objXmlHttpReq.Open "GET", FileUrl, False, UserName, Password
            objXmlHttpReq.SetRequestHeader "Cache-Control", "no-store"
            objXmlHttpReq.SetRequestHeader "Pragma", "no-cache"
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
            Else
                err_counter = err_counter + 1
                If err_counter > 3 Then
                    Dim answer As Integer
                    answer = MsgBox("There is an issue, please check your KOBO account and audit URL." & vbNewLine & "Do you want to countinue to download?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
                    If answer = vbYes Then
                        err_counter = 0
                    Else
                        Exit Sub
                    End If
                End If
            End If
         End If
        
NextIteration:
    Next iCell
    Application.StatusBar = False
     Set objXmlHttpReq = Nothing
Exit Sub
errHandler:
    MsgBox "There is an issue, please check your KOBO account and audit URL!      ", vbCritical
    Application.StatusBar = False
     Set objXmlHttpReq = Nothing
End Sub

Sub download_photo()
    On Error GoTo errHandler:
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim FileUrl, photo_url As String
    Dim objXmlHttpReq As Object
    Dim objStream As Object

    progress_form.Show
    DoEvents

    
    record_count = Cells(Rows.Count, 1).End(xlUp).Row
    percentage_value = record_count / 100
    progress_value = record_count / 270
    
    photo_url = ""
    photo_url = GetRegistrySetting("ramSetting", "koboPhotoReg")
    
    If photo_url = "" Then
        MsgBox "Please check the photo URL column!      ", vbCritical
        Unload progress_form
        Exit Sub
    End If
    
    pic_col = column_letter(photo_url)
    
    pic_name_col_number = column_number(photo_url) - 1
    uuid_col_number = column_number("_uuid")
    base_path = ThisWorkbook.path
    If Dir(base_path & "\photo", vbDirectory) = "" Then
        MkDir base_path & "\photo"
    End If
     
    UserName = ""
    Password = ""
    UserName = GetRegistrySetting("ramSetting", "koboUserReg")
    Password = GetRegistrySetting("ramSetting", "koboPasswordReg")
    
    err_counter = 0
    
    For Each iCell In Range(pic_col & "2:" & pic_col & CStr(record_count)).Cells
        
        FileUrl = iCell
         
        progress_form.percentage.Caption = CStr(Round(iCell.Row / percentage_value, 0)) & " %"
        progress_form.bar.Width = CDec(iCell.Row / progress_value)
        
         
        If FileUrl = "" Or Cells(iCell.Row, pic_name_col_number) = "" Then
            GoTo NextIteration
        End If
        
        strFileName = ThisWorkbook.path & "\photo\" & Cells(iCell.Row, uuid_col_number) & "\" & Cells(iCell.Row, pic_name_col_number)
        strFileExists = Dir(strFileName)
        
        If strFileExists = "" Then
         
             Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
             objXmlHttpReq.Open "GET", FileUrl, False, UserName, Password
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
            Else
                err_counter = err_counter + 1
                If err_counter > 3 Then
                    Dim answer As Integer
                    answer = MsgBox("There is an issue, please check your KOBO account and photo URL." & vbNewLine & "Do you want to countinue to download?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
                    If answer = vbYes Then
                        err_counter = 0
                    Else
                        Exit Sub
                    End If
                End If
             End If
          End If
         
NextIteration:
    Next iCell
    Unload progress_form

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Set objXmlHttpReq = Nothing
    Application.StatusBar = False
    Exit Sub
    
errHandler:
    Unload progress_form
    MsgBox "There is an issue, please check your KOBO account and photo URL!      ", vbCritical
    Application.StatusBar = False
    Set objXmlHttpReq = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
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



