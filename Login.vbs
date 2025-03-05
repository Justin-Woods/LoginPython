Option Explicit

Const ForAppending = 8
Const ForReading = 1
Const WshHide = 0

' Declare variables
Dim Shell, FSO, oEnv
Dim Computername, SchoolDrive
Dim LogsFolderPath
Dim objFolder, objFile
Dim strReleaseID
Dim Drive, Path, objNetwork, objDrives, objReg, i
Dim BackupXadmin

LogsFolderPath = "C:\CCRCE\Logs"
BackupXadmin = "\\ad.ccrsb.ca\xadmin-NRHS"

' Main script logic
On Error Resume Next
Set Shell = CreateObject("WSCript.shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set oEnv = Shell.Environment("PROCESS")
oEnv("SEE_MASK_NOZONECHECKS") = 1
ErrCheck

' Drive Mapping
Computername = Shell.ExpandEnvironmentStrings("%COMPUTERNAME%")
dim Namesplit
Namesplit = Split(ComputerName, "-")
SchoolDrive = "\\ad.ccrsb.ca\xadmin-" & Namesplit(0)
MapNetworkDriveIfNotExist "X:", SchoolDrive
ErrCheck

'Create folders
If FSO.FolderExists("C:\CCRSB") Then
    FSO.DeleteFolder("C:\CCRSB")
End If
If Not FSO.FolderExists("C:\CCRCE") Then
    FSO.CreateFolder("C:\CCRCE")
End If
If Not FSO.FolderExists(LogsFolderPath) Then
    FSO.CreateFolder(LogsFolderPath)
	MakeFolderHidden(LogsFolderPath)
Else
	MakeFolderHidden(LogsFolderPath)
End If
If Not FSO.FolderExists(SchoolDrive & "\Audit") Then
    FSO.CreateFolder(SchoolDrive & "\Audit")
End If
If Not FSO.FolderExists(SchoolDrive & "\Audit\Custom") Then
    FSO.CreateFolder(SchoolDrive & "\Audit\Custom")
End If
If Not FSO.FolderExists("C:\CCRCE\Printers") Then
    FSO.CreateFolder("C:\CCRCE\Printers")
    MakeFolderHidden("C:\CCRCE\Printers")
Else
	MakeFolderHidden("C:\CCRCE\Printers")
End If
ErrCheck

Dim loginAppPath
loginAppPath = "\\ad.ccrsb.ca\it-home\IT-SCHOOL-HOME\" & GetCurrentUsername() & "\Login\Login.py"

' Run login App
If FSO.FileExists(loginAppPath) Then
    Shell.Run loginAppPath, 1, False
End If

strReleaseID = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DisplayVersion"
If Shell.RegRead(strReleaseID) <> "22H2" Then
    Shell.Run Quotes("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Endpoint Manager\Configuration Manager\Software Center.lnk"), 2, False
End If

Dim taskname
taskname = "OpenAudit-Startup"
Shell.Run ("schtasks /run /tn """ & taskname & """"), 2, True
ErrCheck

' Prep Printers
If FSO.FolderExists("\\ad.ccrsb.ca\it-home\IT-SCHOOL-HOME\" & GetCurrentUsername() & "\Login\Printer-Setup") Then
    XcopyFiles "\\ad.ccrsb.ca\it-home\IT-SCHOOL-HOME\" & GetCurrentUsername() & "\Login\Printer-Setup", "C:\CCRCE\Printers"
End If

Shell.Run "cmd /c echo n | gpupdate /force", 0, False

' Launch Initial System Setup Tasks as admin
' OR Ucase(Namesplit(1)) = "CART26"
If Ucase(Namesplit(1)) = "Y24" Then
	Shell.Run "runas /user:CCRSB\x-" & GetCurrentUsername() & " ""wscript.exe " & "\\ad.ccrsb.ca\it-home\IT-SCHOOL-HOME\" & GetCurrentUsername() & "\Login\Rename\rename.vbs" & """", 1, False
Else
	Dim tasksVbsPath
	tasksVbsPath = "\\ad.ccrsb.ca\it-home\IT-SCHOOL-HOME\" & GetCurrentUsername() & "\Login\Initial-Setup-Tasks\Tasks.vbs"
	If FSO.FileExists(tasksVbsPath) Then
		Shell.Run "runas /user:CCRSB\x-" & GetCurrentUsername() & " ""wscript.exe " & tasksVbsPath & """", 2, False
	End If
End if
ErrCheck
On Error GoTo 0

' Clean up and close the log file
oEnv.Remove("SEE_MASK_NOZONECHECKS")
Set Shell = Nothing
Set FSO = Nothing

' Functions
Function Quotes(strQuotes)
    Quotes = Chr(34) & strQuotes & Chr(34)
End Function

Sub XcopyFiles(strSource, strDestination)
    Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	' Copies files and folders from strSource to strDestination
    Shell.Run "xcopy.exe """ & strSource & """ """ & strDestination & """ /f /d /s /e /y /h /r", 2, True
    ' Remove files in strDestination that are not in strSource
    Dim objSourceFolder, objDestinationFolder
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objSourceFolder = objFSO.GetFolder(strSource)
    Set objDestinationFolder = objFSO.GetFolder(strDestination)

    For Each objFile In objDestinationFolder.Files
        If Not objFSO.FileExists(objSourceFolder.Path & "\" & objFile.Name) Then
            objFile.Delete
        End If
    Next
    ' Remove empty subfolders in strDestination
    RemoveEmptyFolders objFSO, objDestinationFolder
    Set objDestinationFolder = Nothing
    Set objSourceFolder = Nothing
    Set objFSO = Nothing
    ErrCheck
End Sub

Sub RemoveEmptyFolders(objFSO, objFolder)
	' Removes empty subfolders in objFolder
    Dim objSubFolder, objSubFolders
    Set objSubFolders = objFolder.SubFolders
    On Error Resume Next
    For Each objSubFolder In objSubFolders
        RemoveEmptyFolders objFSO, objSubFolder
        If objSubFolder.Files.Count = 0 And objSubFolder.SubFolders.Count = 0 Then
            objFSO.DeleteFolder objSubFolder.Path, True
        End If
    Next
	Set objFSO = Nothing
    ErrCheck
    On Error GoTo 0
End Sub

Function MakeFolderHidden(strHideFolder)
    On Error Resume Next
    Dim Shell
    Set Shell = CreateObject("WScript.Shell")
    Shell.Run "cmd.exe /c attrib +h """ & strHideFolder & """", 0, True
    If Err.Number <> 0 Then
        LogMessage "Error: " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Function

Sub DriveMap(Drive, Path)
    ' Maps a network drive with the given drive letter and network path
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
    ErrCheck
End Sub

Sub MapNetworkDriveIfNotExist(Drive, Path)
    ' Maps a network drive with the given drive letter and network path if it doesn't exist
    If Not FSO.FolderExists(Drive) Then
        If FSO.FolderExists(Path) Then
            DriveMap Drive, Path
		Else
			DriveMap Drive, BackupXadmin
        End If
    End If
End Sub

Sub LogMessage(strMessage)
    ' Logs a message with the current timestamp to the log file
    objLogFile.WriteLine Now & " - " & strMessage
End Sub

Sub ErrCheck()
    ' Checks for errors and logs the error message
    If Err.Number <> 0 Then
        LogMessage "Error: " & Err.Number & " - " & Err.Description & " - " & Err.Source
        Err.Clear
    End If
End Sub

Function GetCurrentUsername()
    Dim objNetwork
    Set objNetwork = CreateObject("WScript.Network")
    GetCurrentUsername = objNetwork.UserName
    Set objNetwork = Nothing
End Function