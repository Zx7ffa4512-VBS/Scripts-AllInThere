RunAsAdmin()

Dim WS
Set WS = WScript.CreateObject("Wscript.Shell")

'WS.RegWrite "HKEY_CLASSES_ROOT\*\shell\cmd.exe","cmd.exe","REG_SZ"
WS.RegWrite "HKEY_CLASSES_ROOT\*\shell\cmd.exe\command\","cmd.exe","REG_SZ"

'WS.RegWrite "HKEY_CLASSES_ROOT\*\shell\notepad.exe","notepad.exe %1","REG_SZ"
WS.RegWrite "HKEY_CLASSES_ROOT\*\shell\notepad.exe\command\","notepad.exe %1","REG_SZ"

WS.RegWrite "HKEY_CLASSES_ROOT\*\shell\发到共享(&Z)\command\","cmd /c ""@echo 正在发送...&copy ""%1"" \\192.168.0.100\Temp\""","REG_SZ"

'cmd /c "@echo 正在发送...&copy /y \\192.168.0.100\Temp\"
WS.RegWrite "HKEY_CLASSES_ROOT\*\shell\提取共享(&X)\command\","cmd /c ""@echo 正在发送...&xcopy /y /E \\192.168.0.100\Temp\""","REG_SZ"


'REG ADD HKLM\Software\MyCo /v Data /t REG_BINARY /d fe340ead
'win7
'WS.Run "cmd /c reg add " & chr(34) & "HKEY_CLASSES_ROOT\Local Settings\MuiCache\63\AAF68885" & chr(34) & " /v " & chr(34) & "@C:\Windows\system32\notepad.exe,-469" & chr(34) & " /t REG_SZ /d g文本文档 /f",0,True

'winxp
WS.Run "cmd /c reg add " & chr(34) & "HKEY_CURRENT_USER\Software\Microsoft\Windows\ShellNoRoam\MUICache" & chr(34) & " /v " & chr(34) & "@C:\Windows\system32\notepad.exe,-469" & chr(34) & " /t REG_SZ /d g文本文档 /f",0,True


Set WS=Nothing
'WS.RegWrite "HKEY_CLASSES_ROOT\Local Settings\MuiCache\63\AAF68885\@C:" & chr(35) & "Windows" & chr(35) & "system32" & chr(35) & "notepad.exe,-469","h文本文档","REG_SZ"
'------------------------------------------------------------------------
'管理员权限运行，win7下uac
'------------------------------------------------------------------------
Sub RunAsAdmin()
	Dim objShell
	Set objShell = CreateObject("Shell.Application")
	If WScript.Arguments.Count=0 Then 
		objShell.ShellExecute LCase(Right(WScript.FullName,11)), Chr(34) & WScript.ScriptFullName & Chr(34) &" RunAsAdmin",Left(Wscript.ScriptFullName,Len(Wscript.ScriptFullName)-Len(WScript.ScriptName)), "runas", 1
		WScript.Quit
	ElseIf WScript.Arguments(WScript.Arguments.Count-1)<>"RunAsAdmin" Then
		Dim argTmp
		For Each arg In WScript.Arguments
			argTmp=argTmp&arg&" "
		Next 
		objShell.ShellExecute LCase(Right(WScript.FullName,11)), Chr(34) & WScript.ScriptFullName & Chr(34)&" "&argTmp&" RunAsAdmin",Left(Wscript.ScriptFullName,Len(Wscript.ScriptFullName)-Len(WScript.ScriptName)),"runas",1
		WScript.Quit
	End If
	Set objShell=Nothing
End Sub