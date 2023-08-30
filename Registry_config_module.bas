Attribute VB_Name = "Registry_config_module"
Private Const sAPPNAME = "ramApplicationConfig"

Function SaveRegistrySetting(sSectionName As String, sKeyName As String, sSettingValue As String) As Boolean
    On Error GoTo Error_Handler

    Call SaveSetting(sAPPNAME, sSectionName, sKeyName, sSettingValue)
    SaveRegistrySetting = True
    
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Source: SaveRegistrySetting" & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Function GetRegistrySetting(sSectionName As String, sKeyName As String) As String
    'Returns "" if app, section or key are not found
    'Always returns a VarType = vbString / TypeName = String (stored as REG_SZ), so use conversion function after retrieving as req'd
    On Error GoTo Error_Handler

    GetRegistrySetting = GetSetting(sAPPNAME, sSectionName, sKeyName)
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Source: GetRegistrySetting" & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Function GetAllRegistrySettings(sSectionName As String) As Variant
    On Error GoTo Error_Handler
    Dim aSectionSettings      As Variant
    Dim iCounter              As Long

    aSectionSettings = GetAllSettings(sAPPNAME, sSectionName)
    GetAllRegistrySettings = aSectionSettings

    If IsEmpty(aSectionSettings) = True Then
        GetAllRegistrySettings = Null
    Else
        For iCounter = LBound(GetAllRegistrySettings, 1) To UBound(GetAllRegistrySettings, 1)
            Debug.Print aSectionSettings(iCounter, 0), aSectionSettings(iCounter, 1)
        Next iCounter
    End If
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Source: GetAllRegistrySettings" & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Function DeleteRegistrySetting(sSectionName As String, Optional sKeyName As String) As Boolean
    On Error GoTo Error_Handler

    If sKeyName = "" Then
        'Delete the entire section and all its keys
        Call DeleteSetting(sAPPNAME, sSectionName)
    Else
        'Delete a specific section/key
        Call DeleteSetting(sAPPNAME, sSectionName, sKeyName)
    End If
    DeleteRegistrySetting = True

Error_Handler_Exit:
    On Error Resume Next
    Exit Function

Error_Handler:
    If err.Number <> 5 Then
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
               "Error Source: DeleteRegistrySetting" & vbCrLf & _
               "Error Number: " & err.Number & vbCrLf & _
               "Error Description: " & err.Description & _
               Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
               , vbOKOnly + vbCritical, "An Error has Occurred!"
    End If
    Resume Error_Handler_Exit
End Function


