On Error Resume Next
'--------------
If WScript.Arguments.length =0 Then
	Set objShell = CreateObject("Shell.Application")
	Set Shell = CreateObject ("WSCript.shell" )
	Set FSO=CreateObject("Scripting.FileSystemObject")
	cd = FSO.GetParentFolderName(WScript.ScriptFullName)
	temp = Shell.ExpandEnvironmentStrings("%Temp%")
	FSO.CopyFile WScript.ScriptFullName,  temp & "\" & WScript.ScriptName, True
	If FSO.FileExists(cd & "\backup_gui.py") Then
		FSO.CopyFile cd & "\backup_gui.py",  temp & "\backup_gui.py", True
	End If
	objShell.ShellExecute "wscript.exe", Chr(34) & temp & "\" & WScript.ScriptName & Chr(34) & " uac", "", "runas", 1
Else
	Set Shell = CreateObject ("WSCript.shell" )
	Set FSO=CreateObject("Scripting.FileSystemObject")
	cd = FSO.GetParentFolderName(WScript.ScriptFullName)
	If FSO.FileExists(cd & "\backup_gui.py") Then
		Shell.Run cd & "\backup_gui.py", 2, False
	End If	
End If
