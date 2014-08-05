RunWithCscript(WScript.Arguments.Count)
Do
	WScript.StdOut.Write "PID:"
	pid=WScript.StdIn.ReadLine
	ret=ExeCmd("wmic process where handle="&pid&" get Caption,CommandLine,Description,ExecutablePath")
	WScript.Echo ret
Loop


'--------------------------------------------------------------------------------
'只能在cscript中使用，在wscript中会弹黑框,CmdStr中绝不可以加 cmd /c 后果自负
Function ExeCmd(CmdStr)
	Dim WS
	Set WS = WScript.CreateObject("Wscript.Shell")
	Set CMD=WS.Exec("%comspec%")
	cmd.StdIn.WriteLine CmdStr
	cmd.StdIn.Close
	cmdERR=cmd.StdErr.ReadAll
	If cmdERR <> "" Then 
		ExeCmd=cmdERR
	Else
		For i=0 To 3 
			cmd.StdOut.SkipLine
		Next
		Do Until cmd.StdOut.AtEndOfStream
			tmp=tmp&vbCrLf&cmd.StdOut.ReadLine
			If Not cmd.StdOut.AtEndOfStream Then ExeCmd=tmp
		Loop
	End If
	
	Set CMD=Nothing
	Set WS=Nothing
End Function








'-------------------------------------------------------------------------------------------------------
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