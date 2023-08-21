'*******************************************************************************
' Purpose   : Setup the specified Trusted Location (TL)
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/

'*******************************************************************************
Const HKEY_CLASSES_ROOT              = &H80000000
Const HKEY_CURRENT_USER 			 = &H80000001

Call CreateTrustedLocation

Sub CreateTrustedLocation()
	Dim oRegistry
	Dim sKeyName				'Registry Key Name - default is Location1, Location2, ...
	Dim sPath					'Path to set as a Trusted Location	
	Dim sDescription			'Description of the Trusted Location
	Dim bAllowSubFolders		'Enable subFolders as Trusted Locations
	Dim bAllowNetworkLocations 	'Enable Network Locations as Trusted
								'	Locations
	Dim sOverWriteExistingTL	'Should this routine overwrite the entry if it already
								'	options are: Overwrite, New, Exit
	Dim bAlreadyExists			'Does the path already have an entry?
	Dim sParentKey
	Dim iLocCounter				'Counter
	Dim aChildKeys				'Array of Child Registry Keys
	Dim sChildKey				'Individual Registry Key
	Dim sValue					'Value
	Dim sNewKey					'New Key to Create
	Dim sAppName   				'Name of the application to create the Trusted Location for Access, Excel, Word
	
 
'User defined values for the script 
'*******************************************************************************

	Set oWS = WScript.CreateObject("WScript.Shell")
	userProfile = oWS.ExpandEnvironmentStrings( "%userprofile%" )
	Dim strFolder : strFolder =  userProfile & "\AppData\Roaming\Microsoft\AddIns\" & "direct.xlam"
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	currentFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
	Dim strFile : strFile = currentFolder & "\direct.xlam"
	Const Overwrite = True
	Dim oFSO
	Set oFSO = CreateObject("Scripting.FileSystemObject")

	On Error Resume Next
	'On Error Goto 0
	oFSO.CopyFile strFile, strFolder, True

	If Err.Number <> 0 Then
		MsgBox "Installation failed."  & vbCrLf  & _
			"- Please make sure the MS Excel application 2016 (64 bit) or higher is installed."  & vbCrLf  & _
			"- Please close the MS Excel application and try again.",,"Fail" 
		Exit Sub
		'WScript.Echo "Error in DoStep1: " & Err.Description
		'Err.Clear
	End If

	'Name of the application to create the Trusted Location for Access, Excel, Word
	sAppName = "Excel"
	'Name of the Trusted Location registry key, normally Location, Location1, ...
	sKeyName = "Direct" 
	'Path to be added as a Trusted Location - ie: c:\databases\
	sPath = userProfile & "\AppData\Roaming\Microsoft\AddIns\"
	'Description of the Trusted Location
	sDescription = "Direct utility"
	'Should sub-folders of this Trusted Location also be trusted?
	bAllowSubFolders = True
	'Should network paths be allowed to be Trusted Locations?  Typically, No = False
	bAllowNetworkLocations = False
	'Should this routine overwrite the entry if it already exist
	sOverWriteExistingTL = "Overwrite" '"New", "Overwrite", "Exit"

'Do NOT Alter Anything Beyond This Point Unless You Know What You Are Doing!!!!!
'*******************************************************************************
	bAlreadyExists = False
	
	Set oRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
	oRegistry.GetStringValue HKEY_CLASSES_ROOT, sAppName & ".Application\CurVer", "", sValue
	If IsNull(sValue) Then
		'This message box is optional, feel free to comment it out
		MsgBox "Microsoft " & sAppName & " does not appear to be installed on this computer?!  Cannot continue with the Trusted Location configuration."
	Else
		sValue = Mid(sValue, InStr(sValue, "n.") + 2)
		If sValue >= 12 Then 'Only need to define Trusted Location for Office 2007 Apps or later
			sParentKey = "Software\Microsoft\Office\" & sValue  & ".0\" & sAppName & "\Security\Trusted Locations"	'Trusted Location Reg Key
			
			'Allow Usage of Networked Trusted Locations.  This is NOT recommended
			If bAllowNetworkLocations = True Then
    			oRegistry.SetDWORDValue HKEY_CURRENT_USER, sParentKey, "AllowNetworkLocations", 1
			End If
			
			'Check and see if the Key already exists
			If KeyExists(oRegistry, sParentKey, sKeyName) Then
				If sOverWriteExistingTL = "Exit" Then Exit Sub
				If sOverWriteExistingTL = "New" Then
					sKeyName = sKeyName & GetNextKeySequenceNo(oRegistry, sParentKey, sKeyName)
				End If
				oRegistry.DeleteKey HKEY_CURRENT_USER, sParentKey & "\" & sKeyName
			End If
			
			'Actual Trusted Location Creation in the Registry
			sNewKey = sParentKey & "\" & sKeyName
			oRegistry.CreateKey HKEY_CURRENT_USER, sNewKey
			oRegistry.SetStringValue HKEY_CURRENT_USER, sNewKey, "Date", CStr(Now())
			oRegistry.SetStringValue HKEY_CURRENT_USER, sNewKey, "Description", sDescription
			oRegistry.SetStringValue HKEY_CURRENT_USER, sNewKey, "Path", sPath
			If bAllowSubFolders = True Then
				oRegistry.SetDWORDValue HKEY_CURRENT_USER, sNewKey, "AllowSubFolders", 1
			End If
		End If
	End If

	MsgBox "Direct utility has been installed.",,"Success" 

End Sub
	
Function KeyExists(oReg, sKey, sSearchKey)
	oReg.EnumKey HKEY_CURRENT_USER, sKey, aChildKeys
	For Each sChildKey in aChildKeys
		If sChildKey = sSearchKey Then 
			KeyExists = True
			Exit For
		End If
	Next
End Function

Function GetNextKeySequenceNo(oReg, sKey, sSearchKey)
	Dim lKeyCounter
	
	lKeyCounter = 0
	oReg.EnumKey HKEY_CURRENT_USER, sKey, aChildKeys
	For Each sChildKey in aChildKeys
		If Left(sChildKey, Len(sSearchKey)) = sSearchKey AND Len(sChildKey) > Len(sSearchKey) Then
			If CInt(Mid(sChildKey, Len(sSearchKey) + 1)) > lKeyCounter Then
				lKeyCounter = CInt(Mid(sChildKey, Len(sSearchKey) + 1))
			End If
		End If
	Next
	GetNextKeySequenceNo = lKeyCounter + 1
	
End Function

