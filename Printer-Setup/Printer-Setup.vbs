On Error Resume Next
Set Shell = CreateObject("WScript.shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
cd = FSO.GetParentFolderName(WScript.ScriptFullName)
Dim ARassignedRooms, numPrintersToInstall, ClearedPrinterNames, temp
' Generate a temporary marker file name
temp = Shell.ExpandEnvironmentStrings("%Temp%")
FSO.CreateTextFile "C:\CCRCE\Logs\PrinterSetup.txt", True
If FSO.FileExists("C:\CCRCE\Logs\Printers-4") Then FSO.DeleteFile("C:\CCRCE\Logs\Printers-4")

If WScript.Arguments.length = 0 Then
    RunElevatedScript
ElseIf WScript.Arguments.length = 1 And WScript.Arguments(0) = "uac" Then
    ' Global
    Computername = Shell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    Set oEnv = Shell.Environment("PROCESS")
    oEnv("SEE_MASK_NOZONECHECKS") = 1

    Dim ARassigned, SecurePrint, SecurePrintEnabled
    If Not FSO.FolderExists("C:\CCRCE") Then
        FSO.CreateFolder("C:\CCRCE")
    End If
    If Not FSO.FolderExists("C:\CCRCE\Printers") Then
        FSO.CreateFolder("C:\CCRCE\Printers")
        MakeFolderHidden("C:\CCRCE\Printers")
    Else
        MakeFolderHidden("C:\CCRCE\Printers")
    End If
    If Not FSO.FolderExists("C:\CCRCE\Logs") Then
        FSO.CreateFolder("C:\CCRCE\Logs")
        MakeFolderHidden("C:\CCRCE\Logs")
    End If
    Call XcopyFiles(cd, "C:\CCRCE\Printers")
    AppendFile "SystemInfo", "C:\CCRCE\Logs\PrinterSetup.txt"

    ' System info School, Room
    Dim w64, OSArch, School, Room
    SystemInfo
	'wscript.echo School
    AppendFile "Room assignments: " & Room, "C:\CCRCE\Logs\PrinterSetup.txt"
    ' Initialize ARassignedRooms and numPrintersToInstall variable
    ARassignedRooms = 0
	numPrintersToInstall = 0
    ' Room assignments
    INIPath = "H:\Login\Printer-Setup\Files\Config\"
	
	For intCount = 0 To 30
		PrinterName = ReadIni(INIPath & School & ".ini", Room, "Printer " & intCount)
		If PrinterName <> "" Then
			' Add printer to the list
			SecurePrint = ReadIni(INIPath & School & ".ini", Room, "Secure Print")
			If SecurePrint = 1 Then
				SecurePrintEnabled = True
			End If
			printerList = printerList & PrinterName & vbCrLf
			DefaultUsed=False
		End If
	Next
	
	If printerList = "" Then
        For intCount = 0 To 30
			PrinterName = ReadIni(INIPath & School & ".ini", "Default", "Printer " & intCount)
			If PrinterName <> "" Then
				' Add printer to the list
				SecurePrint = ReadIni(INIPath & School & ".ini", Room, "Secure Print")
				If SecurePrint = "1" Then
					SecurePrintEnabled = True
				End If
				printerList = printerList & PrinterName & vbCrLf
				DefaultUsed=True
			End If
		Next
	End If
	
	If SecurePrintEnabled Then
		If DefaultUsed Then
			PrinterQuitInstall = Shell.Popup("The room:" & Room & " wasn't found, assigning from default for " & School &". The following printers will be installed with Secure Print Enabled:" & vbCrLf & printerList, 20, "Information", 52)
		Else
			PrinterQuitInstall = Shell.Popup("The following printers will be installed with Secure Print Enabled:" & vbCrLf & printerList, 10, "Information", 36)
		End If
	Else
		If DefaultUsed Then
			PrinterQuitInstall = Shell.Popup("The room:" & Room & " wasn't found, assigning from default for " & School &". The following printers will be installed:" & vbCrLf & printerList, 20, "Information", 52)
		Else
			PrinterQuitInstall = Shell.Popup("The following printers will be installed:" & vbCrLf & printerList, 10, "Information", 36)
		End If
	End If

	If PrinterQuitInstall = 7 Then
		WScript.Quit
	End If

    For intCount = 0 To 30
        PrinterName = ReadIni(INIPath & School & ".ini", Room, "Printer " & intCount)
        If PrinterName <> "" Then
            PrinterConfig(PrinterName)
            ARassigned = ARassigned & "," & PrinterName
            ' Increment the count of printers to install
            numPrintersToInstall = numPrintersToInstall + 1
        End If
    Next

    If ARassigned = "" Then
        For intCount2 = 0 To 30
            PrinterName = ReadIni(INIPath & School & ".ini", "Default", "Printer " & intCount2)
            If PrinterName <> "" Then
                PrinterConfig(PrinterName)
                ARassigned = ARassigned & "," & PrinterName
				' Increment the count of printers to install
				numPrintersToInstall = numPrintersToInstall + 1
            End If
        Next
    End If

    AppendFile "Clear local school printers that aren't in the assignments.", "C:\CCRCE\Logs\PrinterSetup.txt"

    ' Clear local school printers that aren't in the assignments.
    ClearPrinters(School)

    AppendFile "Set Printer config like Secure Print Only", "C:\CCRCE\Logs\PrinterSetup.txt"

    ' Set Printer config like Secure Print Only
    SecurePrint = ReadIni(INIPath & School & ".ini", Room, "Secure Print")
    If SecurePrint = "1" Then
        AppendFile "Enable secure print", "C:\CCRCE\Logs\PrinterSetup.txt"
        ' Enable secure print
        Shell.RegWrite "HKEY_LOCAL_MACHINE\Software\Xerox\PrinterDriver\V5.0\Configuration\RepositoryUNCPath", "C:\CCRCE\Printers\Files\Config", "REG_SZ"
        Shell.RegWrite "HKEY_LOCAL_MACHINE\Software\Xerox\PrinterDriver\V5.0\Configuration\CacheExpirationInMinutes", 0, "REG_DWORD"
    Else
        AppendFile "Break the reg path so secure print only enforcement doesn't work.", "C:\CCRCE\Logs\PrinterSetup.txt"
        ' Break the reg path so secure print only enforcement doesn't work.
        Shell.RegWrite "HKEY_LOCAL_MACHINE\Software\Xerox\PrinterDriver\V5.0\Configuration\RepositoryUNCPath", "C:\CCRCE\Printers\Files", "REG_SZ"
    End If
	'Force SecurePrint on everyone...
	Shell.RegWrite "HKEY_LOCAL_MACHINE\Software\Xerox\PrinterDriver\V5.0\Configuration\RepositoryUNCPath", "C:\CCRCE\Printers\Files\Config", "REG_SZ"
	Shell.RegWrite "HKEY_LOCAL_MACHINE\Software\Xerox\PrinterDriver\V5.0\Configuration\CacheExpirationInMinutes", 0, "REG_DWORD"

    AppendFile "Say it's done.", "C:\CCRCE\Logs\PrinterSetup.txt"
    Set objFile = FSO.CreateTextFile("C:\CCRCE\Logs\Printers-4")
	If ClearedPrinterNames <> vbCrLf Then
	    Shell.Popup "The following Printers were installed: " & vbCrLf & printerList, 30, "Information", 64
	Else
		Shell.Popup "The following Printers were removed: "  & vbCrLf & ClearedPrinterNames & vbCrLf & " The following Printers were installed: " & vbCrLf & printerList, 30, "Information", 64
	End If

    oEnv.Remove("SEE_MASK_NOZONECHECKS")
End If

' Send individual printer configuration to the install function
Function PrinterConfig(Name)
    AppendFile "Send individual printer configuration to the install function =" & Name, "C:\CCRCE\Logs\PrinterSetup.txt"
    PrinterIP = ReadIni(INIPath & "config.ini", Name, "IP Address")
    If PrinterIP <> "" Then
        PrinterLocation = ReadIni(INIPath & "config.ini", Name, "Location")
        InstallPrinter Name, PrinterIP, PrinterLocation
        ARassigned = ARassigned & "," & Name
    End If
End Function

' Install the passed printer
Function InstallPrinter(Name, IP, Location)
    AppendFile "Install the passed printer = " & Name, "C:\CCRCE\Logs\PrinterSetup.txt"
    Dim strComputer, objWMIService, objNewPort, Driver, INF
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    AppendFile "Port configuration = " & IP, "C:\CCRCE\Logs\PrinterSetup.txt"

    ' Port configuration
    Set objNewPort = objWMIService.Get("Win32_TCPIPPrinterPort").SpawnInstance_
    objNewPort.Name = "IP_" & IP
    objNewPort.Protocol = 1
    objNewPort.HostAddress = IP
    objNewPort.PortNumber = "9100"
    objNewPort.SNMPEnabled = False
    objNewPort.Put_

    AppendFile "Driver", "C:\CCRCE\Logs\PrinterSetup.txt"

    ' Driver
    Driver = "Xerox Global Print Driver PS"
    INF = "C:\CCRCE\Printers\Files\Drivers\UNIV_5.979.3.0_PS_x64_Driver.inf\x3UNIVP.inf"

    Shell.run "%comspec% /c RUNDLL32 PRINTUI.DLL,PrintUIEntry /ia /m " & Quotes(Driver) & " /h " & w64 & " /v " & Quotes("Type 3 - User Mode") & " /f " & Quotes(INF) & " /u", 0, True

    AppendFile "Printer = " & Name, "C:\CCRCE\Logs\PrinterSetup.txt"

    ' Printer
    Set objPrinter = objWMIService.Get("Win32_Printer").SpawnInstance_
    objPrinter.DriverName = Driver
    objPrinter.PortName = "IP_" & IP
    objPrinter.DeviceID = Name
    objPrinter.Location = Location
    objPrinter.Network = True
    objPrinter.Put_
	
    ' Update the count of assigned printers
    ARassignedRooms = ARassignedRooms + 1
	
	shell.Popup "The following printer has been installed: " & Name ,2,"Information",64
End Function

Function Quotes(strQuotes)
    Quotes = Chr(34) & strQuotes & Chr(34)
End Function

Sub RunElevatedScript()
    Set objShell = CreateObject("Shell.Application")
    FSO.CopyFile WScript.ScriptFullName, temp & "\" & WScript.ScriptName, True
    objShell.ShellExecute "wscript.exe", Chr(34) & temp & "\" & WScript.ScriptName & Chr(34) & " uac", "", "runas", 1

    ' Wait for the elevated script to create a marker file
    Do While Not FSO.FileExists("C:\CCRCE\Logs\Printers-4")
        WScript.Sleep 1000
    Loop
End Sub

Function SystemInfo()
    Computername = Shell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    Dim Nameinfo, colItems, objItem
    Nameinfo = Split(ComputerName, "-")
	SettingsIni = "H:\Login\settings.ini"
	ManualRoom = ReadIni(SettingsIni, "Teacher", ComputerName)
	If ManualRoom <> "" Then 
		Room = Replace(ManualRoom, """", "")
	Else
		'wscript.echo Nameinfo(1)
		Room = Nameinfo(1)
	End If
	Room = Replace(Room, ".", "")
    School = Nameinfo(0)
	'wscript.echo School
    Set colItems = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
    For Each objItem in colItems
        OSArch = objItem.OSArchitecture
    Next

    If OSArch = "64-bit" Then
        w64 = "x64"
    Else
        w64 = "x86"
    End If
End Function

Sub XcopyFiles(strSource, strDestination)
    AppendFile "XcopyFiles =" & strSource, "C:\CCRCE\Logs\PrinterSetup.txt"
    Shell.Run "xcopy.exe """ & strSource & """ """ & strDestination & """ /f /d /s /e /y /h /r", 2, True
End Sub

Function ClearPrinters(School)
    AppendFile "ClearPrinters " & School, "C:\CCRCE\Logs\PrinterSetup.txt"
    Dim STprinter, objItem, colItems
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_Printer", , 48)
    For Each objItem in colItems
        STprinter = (InStr(objItem.DeviceID, School))
        If STprinter = 1 Then
            ARInstalled = ARInstalled & "," & objItem.DeviceID
            STassigned = InStr(ARassigned, objItem.DeviceID)
            If STassigned = 0 Then
                ' Actually deletes the printer
                Shell.Run "%comspec% /c rundll32 printui.dll,PrintUIEntry /dl /q /n " & Chr(34) & objItem.DeviceID & Chr(34), 0
				ClearedPrinterNames = objItem.DeviceID & vbCrLf
            Else
                ' Display printer properties
                Shell.Run "%comspec% /c rundll32 printui.dll,PrintUIEntry /p /n " & Chr(34) & objItem.DeviceID & Chr(34) & " /t7", 0
            End If
        End If
    Next
End Function

Function MakeFolderHidden(strHideFolder)
    Set folder = FSO.GetFolder(strHideFolder)
    folder.attributes = folder.attributes Or 2
End Function

Sub AppendFile(line, outFilepath)
    Const ForAppending = 8
    Dim output
    Set output = FSO.OpenTextFile(outFilepath, ForAppending, True)
    output.WriteLine line
    output.Close
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