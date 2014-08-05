RunWithCscript(WScript.Arguments.Count)
If WScript.Arguments.Count = 4 Then 
	strComputer = wscript.arguments(0)
	strUser = wscript.arguments(1)
	strPwd = wscript.arguments(2)
	strFile = wscript.arguments(3)
Else 
	Call Usage
	WScript.Quit
End If 



set olct=createobject("wbemscripting.swbemlocator")
set wbemServices=olct.connectserver(strComputer,"root\cimv2",strUser,strPwd)
Set colMonitoredProcesses = wbemServices.ExecNotificationQuery("select * from __instancecreationevent " & " within 1 where TargetInstance isa 'Win32_Process'")
i = 0
Do While i = 0
	Set objLatestProcess = colMonitoredProcesses.NextEvent
	Wscript.Echo now & " " & objLatestProcess.TargetInstance.CommandLine
	Set objFS = CreateObject("Scripting.FileSystemObject")
	Set objNewFile = objFS.OpenTextFile(strFile,8,true)
	objNewFile.WriteLine Now() & " " & objLatestProcess.TargetInstance.CommandLine
	objNewFile.Close
Loop





Sub RunWithCscript(ArgCount)
	If (LCase(Right(WScript.FullName,11))="wscript.exe") Then 
		Set objShell=WScript.CreateObject("wscript.shell")
		If ArgCount=0 Then 
			objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&WScript.ScriptFullName&chr(34))
		Else
			Dim argTmp
			For Each arg In WScript.Arguments
				argTmp=argTmp&arg&" "
			Next 
			objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&WScript.ScriptFullName&chr(34)&" "&argTmp)
		End If
		WScript.Quit
	End If
End Sub

Sub Usage()
	WScript.Echo "USAGE:" & vbCrLf & "cscript "&chr(34)&WScript.ScriptFullName&chr(34)&" Computer User Password files"
	WScript.Quit 
End Sub 