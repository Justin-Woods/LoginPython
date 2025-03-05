On Error Resume Next
'--------------
If WScript.Arguments.length =0 Then
	Set objShell = CreateObject("Shell.Application")
	objShell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
Else
	dim IP
	Set Shell = CreateObject ("WSCript.shell" )
	ObtainIP()
	arrIP = split(IP,".")
	Select Case arrIP(1)
		Case 123
			SchoolDrive="\\ad.ccrsb.ca\xadmin-wdc"
		Case 55
			SchoolDrive="\\ad.ccrsb.ca\xadmin-nrhs"
		Case 46
			SchoolDrive="\\ad.ccrsb.ca\xadmin-agb"
	End Select	
	SchoolDrive="\\ad.ccrsb.ca\xadmin-nrhs"
	DriveMap "X:", SchoolDrive
	Shell.run "devmgmt.msc", 2, False
	
	Function ObtainIP()
		set objNetwork = CreateObject("WScript.Network")
		strComputer = objNetwork.ComputerName
		set objExec = Shell.Exec("%comspec% /c ping.exe " & strComputer & " -n 1 -w 100 -4")
		do until objExec.Stdout.AtEndOfStream
			strLine = objExec.StdOut.ReadLine
			if (inStr(strLine, "Reply")) then
				strIP = mid(strLine, 11, inStr(strLine, ":") - 11)
				exit do
			end if
		loop
		IP = strIP
	End Function
	
	Function DriveMap(Drive, Path)
		' Map a network drive 
		Dim objNetwork, objDrives, objReg, i
		Dim strLocalDrive, strRemoteShare, strShareConnected, strMessage
		Dim bolFoundExisting, bolFoundRemembered
		Const HKCU = &H80000001
		strLocalDrive = Drive
		strRemoteShare = Path
		bolFoundExisting = False
		Set objNetwork = WScript.CreateObject("WScript.Network")
		' Loop through the network drive connections and disconnect any that match strLocalDrive
		Set objDrives = objNetwork.EnumNetworkDrives
		If objDrives.Count > 0 Then
		  For i = 0 To objDrives.Count-1 Step 2
			If objDrives.Item(i) = strLocalDrive Then
			  strShareConnected = objDrives.Item(i+1)
			  objNetwork.RemoveNetworkDrive strLocalDrive, True, True
			  i=objDrives.Count-1
			  bolFoundExisting = True
			End If
		  Next
		End If
		' If there's a remembered location (persistent mapping) delete the associated HKCU registry key
		If bolFoundExisting <> True Then
		  Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		  objReg.GetStringValue HKCU, "Network\" & Left(strLocalDrive, 1), "RemotePath", strShareConnected
		  If strShareConnected <> "" Then
			objReg.DeleteKey HKCU, "Network\" & Left(strLocalDrive, 1)
			bolFoundRemembered = True
		  End If
		End If
		'Now actually do the drive map
		objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, False
	End Function
End If