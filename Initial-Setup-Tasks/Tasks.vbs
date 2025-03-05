Option Explicit

' Constants
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const JOIN_DOMAIN = 1
Const ACCT_CREATE = 2
Const ACCT_DELETE = 4
Const DOMAIN_JOIN_IF_JOINED = 32
Const HKLM = &H80000002

' Declare variables and objects
Dim objFSO, objShell, Shell, FSO, objWMIService, objLogFile, strLogFile
Dim w64, OSArch, Computername, SysLaptop, SchoolDrive, SchoolBackupShare, HomeDrive
Dim strReleaseID, NRHS2079, AdobeInstall, INIPath, RestartSelect, taskname, temp, oEnv, BackupXadmin

' Define log file path and create the log file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set FSO=CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
BackupXadmin = "\\ad.ccrsb.ca\xadmin-NRHS"

' Ensure the log folder exists and is hidden
strLogFile = "C:\CCRCE\Logs\Tasks.log"
If Not objFSO.FolderExists("C:\CCRCE") Then
	objFSO.CreateFolder("C:\CCRCE")
End If
If Not objFSO.FolderExists("C:\CCRCE\Logs") Then
	objFSO.CreateFolder("C:\CCRCE\Logs")
	MakeFolderHidden("C:\CCRCE\Logs")
End If
' Create or open the log file
Set objLogFile = objFSO.OpenTextFile(strLogFile, ForAppending, True)

' Error handling section (before the main code)
On Error Resume Next
If Err.Number <> 0 Then
    LogMessage "Error before the main code: " & Err.Number & " - " & Err.Description
    Err.Clear
End If

If WScript.Arguments.length = 0 Then
	Set objShell = CreateObject("Shell.Application")
	Set Shell = CreateObject ("WSCript.shell" )
	Set FSO=CreateObject("Scripting.FileSystemObject")
	'Copy this script to local user temp folder
	temp = Shell.ExpandEnvironmentStrings("%Temp%")
	FSO.CopyFile WScript.ScriptFullName,  temp & "\" & WScript.ScriptName, True
	objShell.ShellExecute "wscript.exe", Chr(34) & temp & "\" & WScript.ScriptName & Chr(34) & " uac", "", "runas", 1
Else				
	LogMessage "Script started."
	Set Shell = CreateObject ("WSCript.shell" )
	Set FSO = CreateObject("Scripting.FileSystemObject")
	const HKEY_LOCAL_MACHINE = &H80000002 
	const HKEY_CURRENT_USER = &H80000001
	Computername = Shell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
	set oEnv = shell.Environment("PROCESS")
	oEnv("SEE_MASK_NOZONECHECKS") = 1
	If FSO.FolderExists("C:\CCRSB") Then
		FSO.DeleteFolder("C:\CCRSB")
	End If
	If NOT FSO.FolderExists("C:\CCRCE") Then
		FSO.CreateFolder("C:\CCRCE")
	End If
	If NOT FSO.FolderExists("C:\CCRCE\Logs") Then
		FSO.CreateFolder("C:\CCRCE\Logs")
		MakeFolderHidden("C:\CCRCE\Logs")
	Else
		MakeFolderHidden("C:\CCRCE\Logs")
	End If
	'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	'Start doing stuff
	'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	Function BackupDict()
		Set AbbreviationIPsBackupShares = CreateObject("Scripting.Dictionary")
		AbbreviationIPsBackupShares.Add "WDC", "\\10.123.3.70\Teacher\backup.bat"
		AbbreviationIPsBackupShares.Add "NRHS", "\\10.55.2.184\Teacher\backup.bat"
		AbbreviationIPsBackupShares.Add "AGB", "\\10.46.3.215\Storage\Backup\backup.bat"
	End Function
	
	ObtainSysInfo
	dim AbbreviationIPs, AbbreviationIPsBackupShares
    BackupDict()
	dim Namesplit
	Namesplit = Split(ComputerName, "-")
	SchoolDrive = "\\ad.ccrsb.ca\xadmin-" & Namesplit(0)
	MapNetworkDriveIfNotExist "X:", SchoolDrive
	SchoolBackupShare = AbbreviationIPsBackupShares(Namesplit(0))
	If FSO.FileExists(SchoolBackupShare) Then
		Shell.run SchoolBackupShare, 2, False
	End If
	HomeDrive = "\\ad.ccrsb.ca\it-home\IT-SCHOOL-HOME\" & GetCurrentUsername()
	MapNetworkDriveIfNotExist "H:", HomeDrive

	'# Check for rename =====================================================================================================================
	'RenameINI = "H:\Scripts\Rename.ini"
	'NewName = ReadIni(RenameINI, Computername, "Name")
	'If NewName <> "" Then
	'	Shell.Run("powershell.exe -noexit -Command " & Quotes(NewName)),, True
	'End If
	'# Insight =====================================================================================================================
	If NOT FSO.FileExists("C:\CCRCE\Logs\Insight-1") Then
		Shell.run HomeDrive & "\Software\Insight-AD\run.vbs",2, True
		Set objFile = FSO.CreateTextFile("C:\CCRCE\Logs\Insight-1")
		Set objFile = FSO.CreateTextFile("C:\CCRCE\Logs\Insight-1-" & ComputerName)
	End If
	If NOT FSO.FileExists("C:\CCRCE\Logs\Insight-1-" & ComputerName) Then
		InsightChannel
		Set objFile = FSO.CreateTextFile("C:\CCRCE\Logs\Insight-1-" & ComputerName)
	End If
	If SysLaptop = True Then Shell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Insight\ConnectionServerAddress", Namesplit(0) & "-WIN-IT-01.ad.ccrsb.ca:8080", "REG_SZ"
	Shell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Insight\School", Namesplit(0), "REG_SZ"
	
	'# Insight V11=====================================================================================================================
	Dim productCode, versionToCheck, result
	productCode = "{2BFB1EFC-E73E-4723-BC83-029DCB62E593}"
	versionToCheck = "11.40.2100.412"
	result = CheckMSIProductVersion(productCode, versionToCheck)
	If result = "Not Found" Then
		shell.Run("powershell.exe -ExecutionPolicy Bypass -File " & Quotes(HomeDrive & "\Software\Insight\Insight.ps1")), 0, True
	End If
	'# Battery Check =====================================================================================================================		
	INIPath = SchoolDrive & "\Audit\Custom\"
	Shell.Run "cmd /c echo n | powercfg /batteryreport /output " & INIPath & Computername &"-battery-report.html",0, False
	'# Cleanup drive =====================================================================================================================	
	shell.Run("powershell.exe -ExecutionPolicy Bypass -File " & Quotes(HomeDrive & "\Scripts\DeleteStudentUserAccounts.ps1") & " -DaysBeforeDeletion 30"), 0, True
	'Clear SCCM Cache
	Dim oUIResManager, oCache, oCacheElements, oCacheElement
	set oUIResManager = createobject("UIResource.UIResourceMgr")
	set oCache=oUIResManager.GetCacheInfo()
	set oCacheElements=oCache.GetCacheElements
	for each oCacheElement in oCacheElements
		oCache.DeleteCacheElement(oCacheElement.CacheElementID)
	next
	'# Random Software installs =====================================================================================================================		
	'Shell.run "msiexec /i " & Quotes("X:\SOFTWARE\MatchGraph!-2.2.0.2.msi") & " /q", 2, TRUE
	If Namesplit(1) = "CART01" Then
		Shell.run SchoolDrive & "\SOFTWARE\Vcarve\Install.bat", 2, TRUE
	End if
	'# Check for needing update
	strReleaseID = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DisplayVersion"
	If Shell.RegRead(strReleaseID) <> "22H2" Then
		If Shell.RegRead(strReleaseID) = "1909" Then
			Shell.run Quotes("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Endpoint Manager\Configuration Manager\Software Center.lnk"), 2, False
		Else
			Win22H2Package = HomeDrive & "\Software\windows10.0-kb5015684-x64_523c039b86ca98f2d818c4e6706e2cc94b634c4a.msu"
			If FSO.FileExists(Win22H2Package) Then
				FSO.CopyFile Win22H2Package, "C:\CCRCE\Logs\", True
				shell.Run "wusa ""C:\CCRCE\Logs\windows10.0-kb5015684-x64_523c039b86ca98f2d818c4e6706e2cc94b634c4a.msu"" /quiet /norestart", 1, TRUE
			Else
				Shell.run Quotes("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Endpoint Manager\Configuration Manager\Software Center.lnk"), 2, False
			End If
		End If
	End If

	'# Adobe Creative Cloud =====================================================================================================================		
	NRHS2079 = Left(UCase(ComputerName),9)
	AdobeInstall = "No"
	If NRHS2079 = "NRHS-2079" Then
		AdobeInstall = "Yes"
	End If
	If InStr(ComputerName,"VS") <> 0 Then
		AdobeInstall = "Yes"
	End If
	If ComputerName = "NRHS-TCH-LT49" Then
		AdobeInstall = "Yes"
	End If	
	If ComputerName = "NRHS-TCH-LT46" Then
		AdobeInstall = "Yes"
	End If	
	If ComputerName = "NRHS-TCH-LT09" Then
		AdobeInstall = "Yes"
	End If	
	IF AdobeInstall = "Yes" Then
		If NOT FSO.FolderExists("C:\Program Files\Adobe\Adobe Creative Cloud") Then
			'Install Adobe Creative Cloud
			Shell.run SchoolDrive & "\SOFTWARE\CCRCEHSTechED\Build\setup.exe", 2, TRUE
			'Disable autostart
			Shell.Regwrite "HKLM\Software\Policies\Adobe\CCXWelcome\Disabled", 0,"REG_DWORD"
		End If
		'Disable autostart
		Shell.Regwrite "HKLM\Software\Policies\Adobe\CCXWelcome\Disabled", 0,"REG_DWORD"
	End If
	'# Update AD stuff =====================================================================================================================
	FixTrustRelationship()
	Shell.Run "cmd /c echo n | gpupdate /force",0, False
	'Application Deployment Evaluation Cycle
	shell.run("cmd /c WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule '{00000000-0000-0000-0000-000000000121}' /NOINTERACTIVE"),0,true
	'Discovery Data Collection Cycle
	shell.run("cmd /c WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule '{00000000-0000-0000-0000-000000000003}' /NOINTERACTIVE"),0,true
	'Software Inventory Cycle
	shell.run("cmd /c WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule '{00000000-0000-0000-0000-000000000002}' /NOINTERACTIVE"),0,true
	'Software Updates Assignments Evaluation Cycle
	shell.run("cmd /c WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule '{00000000-0000-0000-0000-000000000108}' /NOINTERACTIVE"),0,true
	'Software Update Scan Cycle
	shell.run("cmd /c WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule '{00000000-0000-0000-0000-000000000113}' /NOINTERACTIVE"),0,true
	taskname = "OpenAudit-Startup"
	shell.Run("schtasks /run /tn """ & taskname & """"),0,true
	'# Drivers =====================================================================================================================
	'Check drivers
	Shell.run "devmgmt.msc", 2, False
	For i = 1 To 3
		DriversPath = "C:\CCRCE\Logs\Drivers-" & i
		If FSO.FileExists(DriversPath) Then FSO.DeleteFile(DriversPath)
	Next
	'# WOL =====================================================================================================================
	If NOT FSO.FileExists("C:\CCRCE\Logs\WOL-1") Then
		'Software enable WOL
		WOL()
		Const intWindowStyle = 0
		Set objWmi = GetObject("winmgmts:root\cimv2")
		Set colAdapterConfigs = objWmi.ExecQuery ("Select * From Win32_NetworkAdapterconfiguration Where IPEnabled = True")
		For Each objAdapterConfig In colAdapterConfigs
		  Set colAdapters = objAdapterConfig.Associators_(, "Win32_NetworkAdapter")
		  For Each objAdapter In colAdapters
			Shell.Run "%comspec% /c powercfg /deviceenablewake """ & objAdapter.Name & """", intWindowStyle, True
		  Next
		Next
		Set objFile = FSO.CreateTextFile("C:\CCRCE\Logs\WOL-1")
	End If
	'# NRHS ArtCam =====================================================================================================================
	Dim NRHSCART01
	NRHSCART01 = Left(UCase(ComputerName),11)
	If NRHSCART01 = "NRHS-CART01" Then
		If NOT FSO.FolderExists("C:\Program Files\ArtCAM 2013") Then
			Shell.run SchoolDrive & "\SOFTWARE\ArtCam2013\ArtCamInstall.exe", 2, TRUE
		End If	
	End If
	'# Admin Audit =====================================================================================================================
	admin_audit_script_path = HomeDrive & "\Scripts\Admin-Audit.ps1"
	Shell.Run "powershell.exe -ExecutionPolicy Bypass -File " & Quotes(admin_audit_script_path), 0, True
	'# Printers =====================================================================================================================
	For i = 1 To 3
		PrintersPath = "C:\CCRCE\Logs\Printers-" & i
		If FSO.FileExists(PrintersPath) Then FSO.DeleteFile(PrintersPath)
	Next
	If NOT FSO.FileExists("C:\CCRCE\Logs\Printers-4") Then
		Shell.run HomeDrive & "\Login\Printer-Setup\Printer-Setup.vbs", 0, TRUE
	End If
	'# Complete =====================================================================================================================
	RestartSelect = shell.Popup ("Settings update complete. Do you want to restart the computer?",,"Information",36)
	If RestartSelect <> 7 Then
		Restart
	End If
	oEnv.Remove("SEE_MASK_NOZONECHECKS")
	' Error handling section (after the main code)
	If Err.Number <> 0 Then
		LogMessage "Error after the main code: " & Err.Number & " - " & Err.Description
		Err.Clear
	End If
	objLogFile.Close
	Set objLogFile = Nothing
End If

'# Functions =====================================================================================================================
Function Restart()
	On Error Resume Next
	If Not IsEmpty(oEnv) Then oEnv.Remove("SEE_MASK_NOZONECHECKS")
	If Not objLogFile Is Nothing Then
		objLogFile.Close
		Set objLogFile = Nothing
	End If
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "shutdown /r /t 0", 0, True
    Set objShell = Nothing
End Function

Function GetCurrentUsername()
    Dim objNetwork
    Set objNetwork = CreateObject("WScript.Network")
    GetCurrentUsername = objNetwork.UserName
    GetCurrentUsername = Replace(GetCurrentUsername, "x-", "")
    Set objNetwork = Nothing
End Function

Function MakeFolderHidden(strHideFolder)
   Set folder = fso.GetFolder(strHideFolder)
   folder.attributes = folder.attributes Or 2
End Function

Sub DriveMap(Drive, Path)
    Dim objNetwork, objDrives, objReg, i
    Dim strLocalDrive, strRemoteShare, strShareConnected
    Const HKCU = &H80000001
    strLocalDrive = Drive
    strRemoteShare = Path
    Set objNetwork = WScript.CreateObject("WScript.Network")
    ' Disconnect existing drive mapping if any
    Set objDrives = objNetwork.EnumNetworkDrives
    If objDrives.Count > 0 Then
        For i = 0 To objDrives.Count - 1 Step 2
            If objDrives.Item(i) = strLocalDrive Then
                objNetwork.RemoveNetworkDrive strLocalDrive, True, True
                Exit For
            End If
        Next
    End If
    ' Map the network drive
    objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, False
End Sub

Sub MapNetworkDriveIfNotExist(Drive, Path)
    If Not FSO.FolderExists(Drive) Then
        If FSO.FolderExists(Path) Then
            DriveMap Drive, Path
		Else
			DriveMap Drive, BackupXadmin
        End If
    End If
End Sub

Function ObtainSysInfo()
    Dim objWMIService, colChassis, objChassis, strChassisType, colItems, objItem
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    ' Get the operating system architecture
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each objItem in colItems
        OSArch = objItem.OSArchitecture
    Next
    ' Determine the architecture string
    If OSArch = "64-bit" Then
        w64 = " (x86)"
    Else
        w64 = ""
    End If
    ' Check if the system is a laptop
    Set colChassis = objWMIService.ExecQuery("SELECT * FROM Win32_SystemEnclosure")
    For Each objChassis in colChassis
        For Each strChassisType in objChassis.ChassisTypes
            If strChassisType > 7 And strChassisType < 15 And strChassisType <> 13 Then
                SysLaptop = "Yes"
            End If
        Next
    Next
End Function

Function Quotes(strQuotes)
	Quotes = chr(34) & strQuotes & chr(34)
End Function 

Function InsightChannel()
	dim CartNumber, Number,Nameinfo,SettingsIni,ManualRoom
	CartNumber = ""
	Number = ""
	Nameinfo = Split(ComputerName, "-")
	SettingsIni = HomeDrive & "\Login\Settings.ini"
	ManualRoom = ReadIni(SettingsIni, "Teacher", ComputerName)
	If ManualRoom <> "" Then Nameinfo = ManualRoom
	Number = Nameinfo(1)
	Number = Replace(Number, ".", "")
	If Number = "" Or Not IsNumeric(Number) Then
		Number = 8888
	Else
		If Left(Number, 1) = "0" Then
			Number = CInt(Number)
		Else
			Number = CStr(CInt(Number))
		End If

		If Number >= 1 And Number <= 16000 Then
			' Do nothing, Number is valid
		Else
			Number = 8888
		End If
	End If
	If InStr(Nameinfo(1), "CART") = 1 Then
		CartValue = Mid(Nameinfo(1), 5)
		CartValue = Replace(CartValue, ".", "")
		If IsNumeric(CartValue) Then
			CartNumber = "1" & String(4 - Len(CartValue), "0") & CartValue
		Else
			CartNumber = 8888
		End If
		If CartNumber = "" Then
			If Left(Number, 4) = "1000" Then
				CartNumber = "1000" & Right(Number, 2)
			ElseIf Left(Number, 3) = "100" Then
				CartNumber = "100" & Right(Number, 2)
			End If
		End If
	End If
	If InStr(Nameinfo(1), "CAB") = 1 OR InStr(Nameinfo(1), "TUB") = 1 Then
		CartValue = Mid(Nameinfo(1), 5)
		CartValue = Replace(CartValue, ".", "")
		If IsNumeric(CartValue) Then
			CartNumber = "11" & String(3 - Len(CartValue), "0") & CartValue
		Else
			CartNumber = 8888
		End If
		If CartNumber = "" Then
			If Left(Number, 4) = "1100" Then
				CartNumber = "1100" & Right(Number, 2)
			ElseIf Left(Number, 3) = "110" Then
				CartNumber = "110" & Right(Number, 2)
			End If
		End If
	End If
	If CartNumber <> "" Then
		Channel = CartNumber
	Else
		Channel = Number
	End If
	'Setting the channel
	Shell.Regwrite "HKLM\Software\Insight\Channel", Channel, "REG_DWORD"
	If OSArch = "64-bit" Then
		Shell.Regwrite "HKLM\Software\Wow6432Node\Insight\Channel", Channel, "REG_DWORD"
	End If
	LogMessage("Insight Channel: " & Channel)
End Function

Function ReadIni(myFilePath, mySection, myKey)
    Dim objFSO, objIniFile
    Dim strFilePath, strSection, strKey
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    ReadIni = ""
    strFilePath = Trim(myFilePath)
    strSection = Trim(mySection)
    strKey = Trim(myKey)
    If objFSO.FileExists(strFilePath) Then
        Set objIniFile = objFSO.OpenTextFile(strFilePath, ForReading)
        Do Until objIniFile.AtEndOfStream
            Dim strLine, intEqualPos, strLeftString
            strLine = Trim(objIniFile.ReadLine)
            If LCase(strLine) = "[" & LCase(strSection) & "]" Then
                Do Until objIniFile.AtEndOfStream
                    strLine = Trim(objIniFile.ReadLine)
                    intEqualPos = InStr(strLine, "=")
                    If intEqualPos > 0 Then
                        strLeftString = Trim(Left(strLine, intEqualPos - 1))
                        If LCase(strLeftString) = LCase(strKey) Then
                            ReadIni = Trim(Mid(strLine, intEqualPos + 1))
                            Exit Do
                        End If
                    End If
                Loop
                Exit Do
            End If
        Loop
        objIniFile.Close
    End If
End Function

Function FixTrustRelationship()
	Dim strDomain, strUser,objNetwork, objComputer, ReturnValue, strComputer
	strDomain = "ad.ccrsb.ca"
	strUser = GetCurrentUsername()
	Set objShell = CreateObject("WScript.Shell")
	Set objNetwork = CreateObject("WScript.Network")
	strComputer = objNetwork.ComputerName
	Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" & strComputer & "'")
	ReturnValue = objComputer.JoinDomainOrWorkGroup(strDomain, Null, strDomain & "\x-" & strUser, Null, JOIN_DOMAIN + ACCT_DELETE + ACCT_CREATE + DOMAIN_JOIN_IF_JOINED)
	If Err <> 0 Then
		objShell.Popup "Join failed with error: " & ReturnValue, 1, "Information", 64
	End If
	Select Case ReturnValue
		Case 5
			WScript.Echo "Access was Denied for adding the Computer to the Domain"
		Case 87
			WScript.Echo "The parameter is incorrect"
		Case 110
			WScript.Echo "The system cannot open the specified object"
		Case 1323
			WScript.Echo "Unable to update the password"
		Case 1326
			WScript.Echo "Logon failure: unknown username or bad password"
		Case 1355
			WScript.Echo "The specified domain either does not exist or could not be contacted"
		Case 2224
			WScript.Echo "Account exists in the Domain"
		Case 2691
			objShell.Popup "The machine is already joined to the domain", 1, "Information", 64
		Case 2692
			WScript.Echo "The machine is not currently joined to a domain"
	End Select
	If ReturnValue > 2692 Then
		objShell.Popup "The Computer was successfully added to the domain", 1, "Information", 64
	Else
		If ReturnValue = 0 Then
			objShell.Popup "The Computer was successfully added to the domain", 1, "Information", 64
		Else
			objShell.Popup "The Computer was NOT added to the domain", 1, "Information", 64
		End If
	End If
End Function

Function WOL()
	Dim objReg, objWMIService, arrayNetCards, objNetCard, strNICguid, strDeviceID, strShowNicKeyName, strShowNicKeyName001, strPnPCapabilitesKeyName, strPnPCapabilitesKeyName001, strComputer
	strComputer = "."
	strShowNicKeyName = "SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\"
	strShowNicKeyName001 = "SYSTEM\CurrentControlSet001\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\"
	strPnPCapabilitiesKeyName = "SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\"
	strPnPCapabilitiesKeyName001 = "SYSTEM\CurrentControlSet001\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\"
	ShowNicdwValue = 1
	PnPdwValue = 32
	On Error Resume Next
	Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set arrayNetCards = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	For Each objNetCard in arrayNetCards
		strNICguid = objNetCard.SettingID
		strDeviceID = Mid(objNetCard.Caption, 6, 4)
		objReg.SetDWORDValue HKLM, strShowNicKeyName & strNICguid & "\Connection", "ShowIcon", ShowNicdwValue
		objReg.SetDWORDValue HKLM, strShowNicKeyName001 & strNICguid & "\Connection", "ShowIcon", ShowNicdwValue
		objReg.SetDWORDValue HKLM, strPnPCapabilitiesKeyName & strDeviceID & "\", "PnPCapabilities", PnPdwValue
		objReg.SetDWORDValue HKLM, strPnPCapabilitiesKeyName001 & strDeviceID & "\", "PnPCapabilities", PnPdwValue
	Next
	Set objReg = Nothing
	Set objWMIService = Nothing
End Function

If Err.Number <> 0 Then
    ' Error handling section (after the main code)
	LogMessage "Error after the main code: " & Err.Number & " - " & Err.Description
    Err.Clear
End If

Sub LogMessage(message)
    ' Subroutine to log messages to the log file
	On Error Resume Next
    If Not objLogFile Is Nothing Then
        objLogFile.WriteLine Now & " - " & message
    End If
    On Error GoTo 0
End Sub

Function CheckMSIProductVersion(productCode, versionToCheck)
    Dim installer, products, product, productInfo
    Dim result
    result = "Not Found" ' Default result if product code is not found
    Set installer = CreateObject("WindowsInstaller.Installer")
    Set products = installer.ProductsEx("", "", 7)
    For Each product In products
        If UCase(product.ProductCode) = UCase(productCode) Then
            productInfo = installer.ProductInfo(product.ProductCode, "VersionString")
            If productInfo = versionToCheck Then
                result = "Version Match"
            Else
                result = "Version Mismatch"
            End If
            Exit For
        End If
    Next
    Set installer = Nothing
    CheckMSIProductVersion = result
End Function