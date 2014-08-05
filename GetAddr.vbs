RunWithCscript(WScript.Arguments.Count)
Do 
	WScript.StdOut.Write "IP:"
	ip=WScript.StdIn.ReadLine
	ret=HttpGet("http://ip138.com/ips138.asp?ip="&ip&"&action=2")
	WScript.Echo RPCFindString(ret,"<h1>(.+?(?=</h1>))")
	WScript.Echo RPCFindString(ret,"<li>(.+?(?=</li>))")
Loop 


'---------------------------------------------------------------------------------------
'
Function HttpGet(url)
	Dim http
	Set http=CreateObject("Msxml2.ServerXMLHTTP")
	http.setOption 2,13056	'忽略https错误
	http.open "GET",url,False
	http.send
	http.waitForResponse 30
	HttpGet = GB2312ToUnicode(http.responseBody)
End Function 

Function GB2312ToUnicode(str)
	With CreateObject("adodb.stream")
		.Type = 1 : .Open
		.Write str : .Position = 0
		.Type = 2 : .Charset = "gb2312"
		GB2312ToUnicode = .ReadText : .Close
	End With
End Function

'-----------------------------------------------------------------------------
'将sSource用sPartn匹配，返回匹配出的值，每个一行
Function RPCFindString(sSource,sPartn)
	Dim RegEx,Match,Matches,SubMatch,ret,ret2
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = sPartn
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(sSource)
	For Each Match In Matches 
		ret=ret&Match.submatches(0)&vbCrLf 
	Next
	RPCFindString=ret	
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