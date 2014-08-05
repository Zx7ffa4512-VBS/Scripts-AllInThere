RunWithCscript()

On Error Resume Next
Dim url,target
If WScript.Arguments.Count = 2 Then 
	url=WScript.Arguments(0)
	target=WScript.Arguments(1)
ElseIf WScript.Arguments.Count = 1 Then 
	url=WScript.Arguments(0)
Else
	Call Usage()
	WScript.StdOut.Write "Enter URL:"
	url=WScript.StdIn.Readline()
	WScript.StdOut.Write "Enter Target:"
	target=WScript.StdIn.Readline()
End If


If url<>"" And target<>"" Then
	Call Download(url,target)
ElseIf url<>"" And target="" Then 
	target=Split(url,"/")(UBound(Split(url,"/")))
	Call Download(url,target)
Else
	WScript.Echo "No Url"
	WScript.Quit
End If

WScript.Echo vbcrlf & "Download Succeed!"



Sub Usage()
	WScript.Echo String(79,"*")
	WScript.Echo "Usage:"
	WScript.Echo "cscript "&Chr(34)&WScript.ScriptFullName&Chr(34)&" URL Target"
	WScript.Echo String(79,"*")&vbCrLf 
End Sub

'------------------------------------------------------------------------
'强制用cscript运行
'------------------------------------------------------------------------
Sub RunWithCscript()
	If (LCase(Right(WScript.FullName,11))="wscript.exe") Then 
		Set objShell=WScript.CreateObject("wscript.shell")
		If WScript.Arguments.Count=0 Then 
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


'------------------------------------------------------------------------
'下载文件，包括网页文件，不考虑编码问题
'------------------------------------------------------------------------
Sub Download(url,target)
	WScript.Echo "Downloading...."
	Const adTypeBinary = 1
	Const adSaveCreateOverWrite = 2
	Dim http,ado
	Set http = CreateObject("Msxml2.ServerXMLHTTP")
	http.SetOption 2,13056 '忽略https错误
	http.open "GET",url,False
	http.send
	Set ado = CreateObject("Adodb.Stream")
	ado.Type = adTypeBinary
	ado.Open
	ado.Write http.responseBody
	ado.SaveToFile target,2
	ado.Close
End Sub