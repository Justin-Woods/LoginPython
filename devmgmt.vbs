On Error Resume Next
'--------------
If WScript.Arguments.length = 0 Then
	Set objShell = CreateObject("Shell.Application")
	Dim scriptPath
	scriptPath = WScript.ScriptFullName
	If Left(scriptPath, 2) = "H:" Then
		scriptPath = "\\ad.ccrsb.ca\it-home\IT-SCHOOL-HOME\" & GetCurrentUsername() & Mid(scriptPath, 3)
	End If
	objShell.ShellExecute "wscript.exe", Chr(34) & scriptPath & Chr(34) & " uac", "", "runas", 1
	While Not WScript.Arguments(0) = "uac"
		WScript.Sleep 100
	Wend
Else
	Set Shell = CreateObject ("WSCript.shell" )
	Set objShell = CreateObject ("WSCript.shell" )
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Computername = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
	Namesplit = Split(ComputerName, "-")
	SchoolDrive = "\\ad.ccrsb.ca\xadmin-" & Namesplit(0)
	BackupXadmin = "\\ad.ccrsb.ca\xadmin-NRHS"
	MapNetworkDriveIfNotExist "X:", SchoolDrive
	MapNetworkDriveIfNotExist "H:", "\\ad.ccrsb.ca\it-home\IT-SCHOOL-HOME\" & GetCurrentUsername()
	Shell.run "devmgmt.msc", 2, False
End If

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

Function GetCurrentUsername()
	Dim objNetwork
	Set objNetwork = CreateObject("WScript.Network")
	GetCurrentUsername = objNetwork.UserName
End Function