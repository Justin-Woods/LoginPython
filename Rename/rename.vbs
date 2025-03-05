On Error Resume Next

If WScript.Arguments.length = 0 Then
    ' Check for administrative privileges and re-run if needed
    Set objShell = CreateObject("Shell.Application")
    Set Shell = CreateObject("WSCript.shell")
    Set FSO = CreateObject("Scripting.FileSystemObject")
	dim temp
    temp = Shell.ExpandEnvironmentStrings("%Temp%")
	cd = FSO.GetParentFolderName(Wscript.ScriptFullName)
    FSO.CopyFile WScript.ScriptFullName, temp & "\rename.vbs", True
	FSO.CopyFile cd & "\output.ini", "C:\CCRCE\Logs\output.ini", True
    objShell.ShellExecute "wscript.exe", Chr(34) & temp & "\rename.vbs" & Chr(34) & " uac", "", "runas", 1
Else
    Set Shell = CreateObject("WSCript.shell")
    Set FSO = CreateObject("Scripting.FileSystemObject")
    tempX = Shell.ExpandEnvironmentStrings("%Temp%")
	dim sOldName
    sOldName = Shell.ExpandEnvironmentStrings("%COMPUTERNAME%")
	OriginalName = sOldName
	Dim BypassRemoveAD
    set oEnv = Shell.Environment("PROCESS")
    oEnv("SEE_MASK_NOZONECHECKS") = 1
	CheckForAssignedName()
	'Prompt for new system name with input validation
	Do
		sNewName = InputBox("Please enter new computer name:" & vbCrLf & vbCrLf & "Computer naming conventions can be found on Teams", "Rename Computer", sOldName)
		' Exit script if blank or closed
		sNewName = Ucase(sNewName)
		If sNewName = "" Then WScript.Quit
		If sNewName = OriginalName Then WScript.Quit
	Loop While Not IsValidComputerName(sNewName)	
	
	IncrementSystemName(sNewName)

    ' Create Temp Wireless profile. My login script will delete this profile after the restart and after the CCRCE cert is updated.
    outFile = tempX & "\wireless-profile-generated.xml"
    Set objFile = FSO.CreateTextFile(outFile, True)
    SSID = "CCRCE-iot"
    HEXSTR = "!!44+Such+think+Gold+21!!"
    objFile.Write "<?xml version=""1.0""?><WLANProfile xmlns=""http://www.microsoft.com/networking/WLAN/profile/v1""><name>" & SSID & "</name><SSIDConfig><SSID><name>" & SSID & "</name></SSID></SSIDConfig><connectionType>ESS</connectionType><connectionMode>auto</connectionMode><MSM><security><authEncryption><authentication>WPA2PSK</authentication><encryption>AES</encryption><useOneX>false</useOneX></authEncryption><sharedKey><keyType>passPhrase</keyType><protected>false</protected><keyMaterial>" & HEXSTR & "</keyMaterial></sharedKey></security></MSM><MacRandomization xmlns=""http://www.microsoft.com/networking/WLAN/profile/v3""><enableRandomization>false</enableRandomization></MacRandomization></WLANProfile>"
    objFile.Close

    ' Add wireless profile
    Shell.Run "cmd /c netsh wlan add profile filename=" & Quotes(outFile), 0, True
    FSO.DeleteFile(outFile)

    ' Trigger auto removal of CCRCE-iot after reboot
    RemoveCCRCEiot()

	If BypassRemoveAD = False Then
		' Remove new system name from AD if you are replacing a dead system then wait 10 seconds for the server to catch up.
		cmdAD = "powershell.exe -command Remove-ADComputer -Identity " & Quotes(sNewName) & " -Confirm:$false"
		output = shell.Exec(cmdAD).StdOut.ReadAll
		
		' Define the strings to check for success and object not found
		successResult = "Success"
		notFoundResult = "Object not found"

		' Check the output for success or object not found
		If InStr(output, successResult) > 0 Then
			' The result is success
			Shell.Popup "The command was successful. Computer object removed.", 1, "Information", 64
			WScript.Sleep 15000
		ElseIf InStr(output, notFoundResult) > 0 Then
			' The object was not found
			Shell.Popup "The computer object was not found in Active Directory.", 1, "Information", 64
		Else
			' Other results (error or different output)
			WScript.Echo "An error occurred or the output was unexpected:" & vbCrLf & output
		End If
    End If

    ' Leave Azure
    Shell.Run "cmd /c dsregcmd /leave", 0, True

    ' Rename the system and pull the results
    cmd = "powershell.exe -Command " & Quotes("Rename-computer -newname " & sNewName)
    Set executor = Shell.Exec(cmd)
    executor.StdIn.Close
    sResults = executor.StdOut.ReadAll

    ' Prompt for Restart and display rename results and delete this script from the temp folder
    RestartSelect = Shell.Popup(sResults & vbCrLf & vbCrLf & " Do you want to restart the computer as " & sNewName & "? ", , "Information", 36)
    If RestartSelect <> "7" Then
        If FSO.FileExists(temp & "\rename.vbs") Then
            FSO.DeleteFile(temp & "\rename.vbs")
        End If
        Restart
    End If
    If FSO.FileExists(temp & "\rename.vbs") Then
        FSO.DeleteFile(temp & "\rename.vbs")
    End If
    oEnv.Remove("SEE_MASK_NOZONECHECKS")
End If

If Err.Number <> 0 Then
    ' Error handling code
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    CleanupAndExit
End If

Function Quotes(strQuotes)
    Quotes = Chr(34) & strQuotes & Chr(34)
End Function

Function Restart()
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\.\root\cimv2")
    Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each objOperatingSystem In colOperatingSystems
        ObjOperatingSystem.Reboot()
    Next
    oEnv.Remove("SEE_MASK_NOZONECHECKS")
    WScript.Quit
End Function

Sub CleanupAndExit()
    ' Clean up and exit the script
    On Error Resume Next
    If FSO.FileExists(temp & "\rename.vbs") Then
        FSO.DeleteFile(temp & "\rename.vbs")
    End If
    oEnv.Remove("SEE_MASK_NOZONECHECKS")
    WScript.Quit
End Sub

Function RemoveCCRCEiot()
	Dim vbsScriptContent
	vbsScriptContent = _
	"On Error Resume Next" & vbCrLf & _
	"" & vbCrLf & _
	"Dim fso, objShell, objExec, Shell, scriptPath" & vbCrLf & _
	"Set objShell = CreateObject(""WScript.Shell"")" & vbCrLf & _
	"" & vbCrLf & _
	"' Check if WScript.Shell was created successfully" & vbCrLf & _
	"If objShell Is Nothing Then" & vbCrLf & _
	"    MsgBox ""Failed to create WScript.Shell object."", vbExclamation, ""Error""" & vbCrLf & _
	"    WScript.Quit 1" & vbCrLf & _
	"End If" & vbCrLf & _
	"" & vbCrLf & _
	"Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
	"" & vbCrLf & _
	"' Check if FileSystemObject was created successfully" & vbCrLf & _
	"If fso Is Nothing Then" & vbCrLf & _
	"    ShowPopup ""Failed to create Scripting.FileSystem Object."", ""Error""" & vbCrLf & _
	"    WScript.Quit 1" & vbCrLf & _
	"End If" & vbCrLf & _
	"" & vbCrLf & _
	"If WScript.Arguments.length = 0 Then" & vbCrLf & _
	"    " & vbCrLf & _
	"    Set Shell = CreateObject(""Shell.Application"")" & vbCrLf & _
	"    " & vbCrLf & _
	"    ' Check if WScript.Shell was created successfully" & vbCrLf & _
	"    If Shell Is Nothing Then" & vbCrLf & _
	"        MsgBox ""Failed to create WScript.Application object."", vbExclamation, ""Error""" & vbCrLf & _
	"        WScript.Quit 1" & vbCrLf & _
	"    End If" & vbCrLf & _
	"    " & vbCrLf & _
	"    'Copy this script to local user temp folder" & vbCrLf & _
	"    temp = objShell.ExpandEnvironmentStrings(""%Temp%"")" & vbCrLf & _
	"    scriptPath = WScript.ScriptFullName" & vbCrLf & _
	"    FSO.CopyFile scriptPath,  temp & ""\"" & WScript.ScriptName, True" & vbCrLf & _
	"    Shell.ShellExecute ""wscript.exe"", Chr(34) & temp & ""\"" & WScript.ScriptName & Chr(34) & "" uac"", """", ""runas"", 1" & vbCrLf & _
	"Else" & vbCrLf & _
	"    Set objExec = objShell.Exec(""certutil -store My"")" & vbCrLf & _
	"    " & vbCrLf & _
	"    ' Check if certutil command executed successfully" & vbCrLf & _
	"    If objExec Is Nothing Then" & vbCrLf & _
	"        ShowPopup ""Failed to execute certutil command."", ""Error""" & vbCrLf & _
	"        WScript.Quit 1" & vbCrLf & _
	"    End If" & vbCrLf & _
	"    " & vbCrLf & _
	"    ' If the scriptPath is not set" & vbCrLf & _
	"    If scriptPath = """" Then" & vbCrLf & _
	"        scriptPath = ""C:\CCRCE\Installers\RemoveCCRCE-iot.vbs""" & vbCrLf & _
	"    End If" & vbCrLf & _
	"    " & vbCrLf & _
	"    Dim systemName, matchFound, maxAttempts, attemptCount" & vbCrLf & _
	"    systemName = objShell.ExpandEnvironmentStrings(""%COMPUTERNAME%"")" & vbCrLf & _
	"    " & vbCrLf & _
	"    If systemName = """" Then" & vbCrLf & _
	"        ShowPopup ""Failed to retrieve system name."", ""Error""" & vbCrLf & _
	"        WScript.Quit 1" & vbCrLf & _
	"    End If" & vbCrLf & _
	"    " & vbCrLf & _
	"    matchFound = False" & vbCrLf & _
	"    maxAttempts = 10" & vbCrLf & _
	"    attemptCount = 0" & vbCrLf & _
	"    " & vbCrLf & _
	"    Do While matchFound = False And attemptCount < maxAttempts" & vbCrLf & _
	"        Do While Not objExec.StdOut.AtEndOfStream" & vbCrLf & _
	"            strLine = objExec.StdOut.ReadLine()" & vbCrLf & _
	"            If InStr(strLine, ""Subject:"") > 0 Then" & vbCrLf & _
	"                subjectLine = Trim(Mid(strLine, InStr(strLine, "":"") + 1))" & vbCrLf & _
	"                subjectLine = Replace(subjectLine, ""CN="", """")" & vbCrLf & _
	"                dotIndex = InStr(subjectLine, ""."")" & vbCrLf & _
	"                If dotIndex > 0 Then" & vbCrLf & _
	"                    subjectLine = Left(subjectLine, dotIndex - 1)" & vbCrLf & _
	"                End If" & vbCrLf & _
	"                If StrComp(subjectLine, systemName, vbTextCompare) = 0 Then" & vbCrLf & _
	"                    matchFound = True" & vbCrLf & _
	"                    Exit Do ' Exit the loop since a match has been found" & vbCrLf & _
	"                End If" & vbCrLf & _
	"            End If" & vbCrLf & _
	"        Loop" & vbCrLf & _
	"        WScript.Sleep 10000" & vbCrLf & _
	"        attemptCount = attemptCount + 1" & vbCrLf & _
	"    Loop" & vbCrLf & _
	"    " & vbCrLf & _
	"    If Not matchFound Then" & vbCrLf & _
	"        ShowPopup ""Wireless profile not found."", ""Profile Not Found""" & vbCrLf & _
	"    Else" & vbCrLf & _
	"        ' Remove wireless profile" & vbCrLf & _
	"        objShell.Run ""cmd /c netsh wlan delete profile name=CCRCE-iot"", 0, True" & vbCrLf & _
	"        ShowPopup ""Wireless profile 'CCRCE-iot' removed successfully."", ""Profile Removed""" & vbCrLf & _
	"    End If" & vbCrLf & _
	"    " & vbCrLf & _
	"    ' Delete the script file itself" & vbCrLf & _
	"    DeleteScriptFile" & vbCrLf & _
	"End If" & vbCrLf & _
	"" & vbCrLf & _
	"Sub DeleteScriptFile()" & vbCrLf & _
	"    If FSO.FileExists(scriptPath) Then" & vbCrLf & _
	"        fso.DeleteFile(scriptPath)" & vbCrLf & _
	"        ShowPopup ""The script file has been deleted."", ""Script File Deleted""" & vbCrLf & _
	"    End If" & vbCrLf & _
	"	taskName = ""\RemoveCCRCE-iot""" & vbCrLf & _
	"	Dim strCommand" & vbCrLf & _
	"	strCommand = ""schtasks.exe /Delete /TN "" & taskName & "" /F""" & vbCrLf & _
	"	objShell.Run strCommand, 0, True" & vbCrLf & _
	"End Sub" & vbCrLf & _
	"" & vbCrLf & _
	"Sub ShowPopup(msg, title)" & vbCrLf & _
	"    objShell.Popup msg, 5, title, vbInformation + vbSystemModal" & vbCrLf & _
	"End Sub"

	Dim wshShell
	Set wshShell = CreateObject("WScript.Shell")

	' Save the script content to a file
	Dim vbsScriptPath
	vbsScriptPath = "C:\CCRCE\Installers\RemoveCCRCE-iot.vbs"

	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")

	If NOT FSO.FolderExists("C:\CCRCE") Then
		FSO.CreateFolder("C:\CCRCE")
	End If
	If NOT FSO.FolderExists("C:\CCRCE\Installers") Then
		FSO.CreateFolder("C:\CCRCE\Installers")
	End If
	Dim vbsScriptFile
	Set vbsScriptFile = fso.CreateTextFile(vbsScriptPath, True)
	vbsScriptFile.Write vbsScriptContent
	vbsScriptFile.Close
	
	' Call the function to add the script to the "RunOnce" registry key
	AddTask()
	
End Function

Function CheckForAssignedName()
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colBIOS = objWMIService.ExecQuery ("Select * from Win32_BIOS")
	For each objBIOS in colBIOS
		serialNumber = CStr(objBIOS.SerialNumber)
	Next	
	AssignedName = ReadIni("C:\CCRCE\Logs\output.ini", serialNumber, "system name")
	If AssignedName <> "" Then
		sOldName = AssignedName
	End If
End Function

Function IsAlphaNumericHyphen(str)
    Dim y
    Dim validChars
    validChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-"

    For y = 1 To Len(str)
        If InStr(validChars, Mid(str, y, 1)) = 0 Then
            ' Character is not valid
			IsAlphaNumericHyphen = False
            Exit Function
        End If
    Next

    ' All characters are valid
    IsAlphaNumericHyphen = True
End Function

Function IsValidComputerName(name)
    ' Check if the computer name follows Windows restrictions
    If Len(name) > 15 Then
        MsgBox "The computer name cannot exceed 15 characters.", vbExclamation, "Invalid Computer Name"
        IsValidComputerName = False
        Exit Function
    ElseIf InStr(name, "--") > 0 Then
        MsgBox "The computer name cannot contain consecutive hyphens.", vbExclamation, "Invalid Computer Name"
        IsValidComputerName = False
        Exit Function
    ElseIf Left(name, 1) = "-" Or Right(name, 1) = "-" Then
        MsgBox "The computer name cannot begin or end with a hyphen.", vbExclamation, "Invalid Computer Name"
        IsValidComputerName = False
        Exit Function
    ElseIf Not IsAlphaNumericHyphen(name) Then
        MsgBox "The computer name can only contain letters (A-Z), numbers (0-9), and hyphens (-).", vbExclamation, "Invalid Computer Name"
        IsValidComputerName = False
        Exit Function
    End If
	
   ' Split the computer name by "-"
    Dim nameParts
    nameParts = Split(name, "-")

    ' Check if the first part nameParts(0) is in the list of abbreviations
    Dim abbreviationList
    abbreviationList = Array("Abbreviation", "CBG", "CHBG", "KBG", "NBG", "TBG", "CO", "NSISP", "CEC", "NRHS", "NNEC", "HERH", "ARH", "REC", "NGA", "WPC", "SCA", "TRES", "TRMS", "DWAM", "AGB", "CCJH", "OREC", "RMS", "SSA", "UDS", "ELD", "WDC", "TRA", "FHM", "GRS", "BHJH", "EBC", "HHE", "CNA", "SJSH", "MRE", "WRCS", "END", "HNRH", "VES", "MEC", "TC", "BHC", "WHE", "NRE", "CCJH", "PDH", "RHDS", "SDE", "PA", "CES", "PRHS", "PRES", "CCE", "JRE", "TE", "TMS", "CEE", "DES", "WEM", "WCC", "HE", "RDE", "SBE", "CDE", "SSE", "GVE", "USE", "ADS", "NCS", "WCS", "NSVS")

    If Not IsEmpty(nameParts) Then
        If UBound(nameParts) > 0 Then
            ' The computer name contains at least one hyphen
            If Not IsNumeric(nameParts(0)) Then
                ' Check if the first part is not numeric
                Dim foundAbbreviation
                foundAbbreviation = False

                ' Loop through the abbreviation list to check if the first part matches any abbreviation
                Dim i
                For i = LBound(abbreviationList) To UBound(abbreviationList)
                    If StrComp(nameParts(0), abbreviationList(i), vbTextCompare) = 0 Then
                        foundAbbreviation = True
                        Exit For
                    End If
                Next

                If Not foundAbbreviation Then
                    MsgBox "The first part of the computer name must match one of the valid abbreviations.", vbExclamation, "Invalid Computer Name"
                    IsValidComputerName = False
                    Exit Function
                End If
            Else
                ' The first part is numeric, which is not allowed
                MsgBox "The first part of the computer name cannot be numeric.", vbExclamation, "Invalid Computer Name"
                IsValidComputerName = False
                Exit Function
            End If
        End If
    End If
    ' The computer name is valid
    IsValidComputerName = True
End Function

Function AddTask()
	Dim xmlContent, xmlFilePath, objFSO

	' Generate the XML content for the task
	xmlContent = "<?xml version=""1.0"" encoding=""UTF-16""?>" & vbCrLf
	xmlContent = xmlContent & "<Task version=""1.2"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">" & vbCrLf
	xmlContent = xmlContent & "  <RegistrationInfo>" & vbCrLf
	xmlContent = xmlContent & "    <Date>2023-07-18T08:47:56.9156699</Date>" & vbCrLf
	xmlContent = xmlContent & "    <Author>CCRSB\" & GetCurrentUsername() & "</Author>" & vbCrLf
	xmlContent = xmlContent & "    <URI>\RemoveCCRCE-iot</URI>" & vbCrLf
	xmlContent = xmlContent & "  </RegistrationInfo>" & vbCrLf
	xmlContent = xmlContent & "  <Triggers>" & vbCrLf
	xmlContent = xmlContent & "    <LogonTrigger>" & vbCrLf
	xmlContent = xmlContent & "      <Enabled>true</Enabled>" & vbCrLf
	xmlContent = xmlContent & "      <Delay>PT1M</Delay>" & vbCrLf
	xmlContent = xmlContent & "    </LogonTrigger>" & vbCrLf
	xmlContent = xmlContent & "  </Triggers>" & vbCrLf
	xmlContent = xmlContent & "  <Principals>" & vbCrLf
	xmlContent = xmlContent & "    <Principal id=""Author"">" & vbCrLf
	xmlContent = xmlContent & "      <UserId>S-1-5-18</UserId>" & vbCrLf
	xmlContent = xmlContent & "      <RunLevel>HighestAvailable</RunLevel>" & vbCrLf
	xmlContent = xmlContent & "    </Principal>" & vbCrLf
	xmlContent = xmlContent & "  </Principals>" & vbCrLf
	xmlContent = xmlContent & "  <Settings>" & vbCrLf
	xmlContent = xmlContent & "    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>" & vbCrLf
	xmlContent = xmlContent & "    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>" & vbCrLf
	xmlContent = xmlContent & "    <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>" & vbCrLf
	xmlContent = xmlContent & "    <AllowHardTerminate>false</AllowHardTerminate>" & vbCrLf
	xmlContent = xmlContent & "    <StartWhenAvailable>false</StartWhenAvailable>" & vbCrLf
	xmlContent = xmlContent & "    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>" & vbCrLf
	xmlContent = xmlContent & "    <IdleSettings>" & vbCrLf
	xmlContent = xmlContent & "      <StopOnIdleEnd>true</StopOnIdleEnd>" & vbCrLf
	xmlContent = xmlContent & "      <RestartOnIdle>false</RestartOnIdle>" & vbCrLf
	xmlContent = xmlContent & "    </IdleSettings>" & vbCrLf
	xmlContent = xmlContent & "    <AllowStartOnDemand>false</AllowStartOnDemand>" & vbCrLf
	xmlContent = xmlContent & "    <Enabled>true</Enabled>" & vbCrLf
	xmlContent = xmlContent & "    <Hidden>false</Hidden>" & vbCrLf
	xmlContent = xmlContent & "    <RunOnlyIfIdle>false</RunOnlyIfIdle>" & vbCrLf
	xmlContent = xmlContent & "    <WakeToRun>false</WakeToRun>" & vbCrLf
	xmlContent = xmlContent & "    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>" & vbCrLf
	xmlContent = xmlContent & "    <Priority>7</Priority>" & vbCrLf
	xmlContent = xmlContent & "  </Settings>" & vbCrLf
	xmlContent = xmlContent & "  <Actions Context=""Author"">" & vbCrLf
	xmlContent = xmlContent & "    <Exec>" & vbCrLf
	xmlContent = xmlContent & "      <Command>C:\Windows\System32\wscript.exe</Command>" & vbCrLf
	xmlContent = xmlContent & "      <Arguments>""C:\CCRCE\Installers\RemoveCCRCE-iot.vbs""</Arguments>" & vbCrLf
	xmlContent = xmlContent & "    </Exec>" & vbCrLf
	xmlContent = xmlContent & "  </Actions>" & vbCrLf
	xmlContent = xmlContent & "</Task>"

	' Specify the path to save the XML file
	xmlFilePath = "C:\CCRCE\Installers\RemoveCCRCE-iot.xml"

	' Create a filesystem object
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	' Write the XML content to the file
	Dim objFile
	Set objFile = objFSO.CreateTextFile(xmlFilePath, True)
	objFile.Write(xmlContent)
	objFile.Close

	' Check if the file was created successfully
	If objFSO.FileExists(xmlFilePath) Then
	  ' Create a shell object
	  Dim objShell
	  Set objShell = CreateObject("WScript.Shell")
	  
	  ' Build the command to import the task using schtasks.exe
	  Dim strCommand
	  strCommand = "schtasks.exe /Create /TN ""\RemoveCCRCE-iot"" /XML """ & xmlFilePath & """ /F"
	  
	  ' Run the command
	  objShell.Run strCommand, 0, True
	  FSO.DeleteFile(xmlFilePath)
	End If
End Function

Function GetCurrentUsername()
    Dim objNetwork
    Set objNetwork = CreateObject("WScript.Network")
    
    GetCurrentUsername = objNetwork.UserName
    
    Set objNetwork = Nothing
End Function

Function ComputerNameExists(computerName, domainController, baseDN)
    On Error Resume Next

    Dim objConnection, objCommand, objRootDSE, objRecordSet

    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"

    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection
    objCommand.Properties("Page Size") = 1000

    objCommand.CommandText = "<" & domainController & "/" & baseDN & ">;" & _
        "(&(objectClass=computer)(name=" & computerName & "));name;subtree"

    Set objRecordSet = objCommand.Execute

    ComputerNameExists = Not objRecordSet.EOF

    objRecordSet.Close
    objConnection.Close
    On Error Goto 0
End Function

Function GetNextFreeName(computerName, domainController, baseDN, maxNumber, Found)
    Dim x, Namesplit, FreeName, LT
	LT = ""
    Namesplit = Split(computerName, "-")
	If Left(Namesplit(2), 2) = "LT" Then
		Namesplit(2) = Mid(Namesplit(2), 3)
		LT = "LT"
	End If
	
    For x = CInt(Namesplit(2)) To maxNumber
        computerName = Namesplit(0) & "-" & Namesplit(1) & "-" & LT & Right("00" & x, 2)

        If Not ComputerNameExists(computerName, domainController, baseDN) Then
            FreeName = computerName
			Found = True
            Exit For
        End If
    Next

    GetNextFreeName = FreeName
End Function

Sub IncrementSystemName(computerName)
    Dim maxNumber
    maxNumber = 900 ' Specify the maximum number to check, e.g., NRHS-1121-01 to NRHS-1121-900
    Dim domainController
    domainController = "LDAP://ad.ccrsb.ca" ' Replace with the actual domain controller name
    Dim baseDN
    baseDN = "DC=ad,DC=ccrsb,DC=ca" ' Replace with the actual base DN of your domain

    Dim Found, FreeName
    Found = False

    FreeName = GetNextFreeName(computerName, domainController, baseDN, maxNumber, Found)

    If Found = True Then
        If FreeName = sNewName Then
			BypassRemoveAD = True
		Else
			Dim IncrementSelect
			IncrementSelect = MsgBox("The system name: " & sNewName & " already exists. Do you want to use the next free name: " & FreeName & "?", vbYesNo + vbQuestion, "Information")

			If IncrementSelect = vbYes Then
				sNewName = FreeName
				BypassRemoveAD = True
			Else
				BypassRemoveAD = False
			End If
		End If
    End If
End Sub

Function ReadIni(myFilePath, mySection, myKey)
    Const ForReading = 1

    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLine, strSection, bInSection
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ReadIni = ""
    strFilePath = Trim(myFilePath)
    strSection = Trim(mySection)
    strKey = Trim(myKey)
    bInSection = False

    If objFSO.FileExists(strFilePath) Then
        Set objIniFile = objFSO.OpenTextFile(strFilePath, ForReading, False)

        Do Until objIniFile.AtEndOfStream
            strLine = Trim(objIniFile.ReadLine)

            ' Check if it's a section header
            If Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then
                ' Update current section
                bInSection = LCase(Mid(strLine, 2, Len(strLine) - 2)) = LCase(strSection)
            Else
                ' Check if line contains an equal sign and belongs to the requested section
                If bInSection And InStr(strLine, "=") > 0 Then
                    Dim arrPair, key, value
                    arrPair = Split(strLine, "=", 2)
                    key = Trim(arrPair(0))
                    value = Trim(arrPair(1))

                    ' Check if the key matches
                    If LCase(key) = LCase(strKey) Then
                        ReadIni = value
                        Exit Do
                    End If
                End If
            End If
        Loop

        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exist. Exiting..."
        Wscript.Quit 1
    End If
End Function