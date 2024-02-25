Attribute VB_Name = "Download_media_module"
Sub download_audit()
    On Error GoTo errhandler:
    Dim FileUrl As String
    Dim audit_url As String
    Dim objXmlHttpReq As Object
    Dim objStream As Object
    
    Dim ws As Worksheet
    Set ws = sheets(find_main_data)
    
    audit_url = ""
    audit_url = GetRegistrySetting("ramSetting", "koboAuditReg")
    
    If audit_url = "" Then
        MsgBox "Please check the audit URL column!      ", vbCritical
        Exit Sub
    End If
    sheets(find_main_data).Activate
    audit_col = column_letter(audit_url)
    
    If audit_col = vbNullString Then
        MsgBox "Please check the audit URL column!      ", vbCritical
        Exit Sub
    End If
    
    uuid_col_number = column_number("_uuid")
    record_count = ws.Cells(ws.Rows.count, uuid_col_number).End(xlUp).Row
    base_path = ActiveWorkbook.path
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
    For Each iCell In ws.Range(audit_col & "2:" & audit_col & CStr(record_count)).Cells
    
        FileUrl = iCell
        
        If FileUrl = "" Then
            GoTo NextIteration
        End If
        
        strFileName = ActiveWorkbook.path & "\audit\" & ws.Cells(iCell.Row, uuid_col_number) & "\audit.csv"
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
                Call make_folder(ws.Cells(iCell.Row, uuid_col_number), "audit")
                
                objStream.SaveToFile ActiveWorkbook.path & "\audit\" & ws.Cells(iCell.Row, uuid_col_number) & "\audit.csv", 2
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
    MsgBox "Audit files downloaded!            ", vbInformation
    Exit Sub

errhandler:
        MsgBox "There is an issue, please check your KOBO account and audit URL!      ", vbCritical
        Application.StatusBar = False
         Set objXmlHttpReq = Nothing
End Sub


Sub make_folder(uuid As String, folder_name As String)
    On Error Resume Next
    base_path = ActiveWorkbook.path
    If Dir(base_path & "\" & folder_name & "\" & uuid, vbDirectory) = "" Then
        MkDir base_path & "\" & folder_name & "\" & uuid
    End If
End Sub

