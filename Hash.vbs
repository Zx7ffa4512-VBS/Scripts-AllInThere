'在线破解hash，可以界面下运行，也可以cmd下
Dim RunWith,fso
RunWith=JudgeRunWith()
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
If RunWith="w" Then 
	Call RunWithWscript
Else
	Call RunWithCscript
End If
'******************************************************************************
'处理参数
Sub RunWithWscript()
	If WScript.Arguments.Count = 1 Then 
		arg=WScript.Arguments(0)
	Else
		arg=InputBox("Enter hash:","Enter",String(32,"0") & ":" & String(32,"0"))
	End If
	If arg<>"" Then Call main(arg)
End Sub

Sub RunWithCscript()
	If WScript.Arguments.Count = 1 Then 
		arg=WScript.Arguments(0)
	Else 
		WScript.StdOut.Write "Enter hash or FileName:"
		arg=WScript.StdIn.Readline()
	End If
	If arg<>"" Then Call main(arg)
End Sub
'******************************************************************************

Sub main(arg)
	Dim HashStr,HashFileStr,IsHash,Ret,Password,PasswordFileStr,BusyFileStr
	PasswordFileStr="Password.txt"
	BusyFileStr="Busy.txt"
	If Len(arg)=65 And InStr(arg,":")=33 Then IsHash=True
	If IsHash Then 
		HashStr=arg
		Password=CreakHash(HashStr)
		OutPutMsg(Password)
		Do 
			WScript.StdOut.Write "Hash:"
			HashStr=WScript.StdIn.ReadLine
			Password=CreakHash(HashStr)
			OutPutMsg(Password)
		Loop 
	Else
		Dim HashFile,Line,ValueArray
		HashFileStr=arg
		Set HashFile=fso.OpenTextFile(HashFileStr,1,True)
		Do Until HashFile.AtEndOfStream
			Line=HashFile.ReadLine
			ValueArray=Split(Line,":")
			If UBound(ValueArray)=3 Then 
				ValueArray(0)=Replace(ValueArray(0),"(current)","")
				If ValueArray(2)=String(32,"*") Or ValueArray(2)=String(32,"0") Then
					WriteFile PasswordFileStr,StringN(40,ValueArray(0))&ValueArray(2)&":"&ValueArray(3)
				Else
					HashStr=ValueArray(2)&":"&ValueArray(3)
					Password=CreakHash(HashStr)
					If InStr(Password,"You are blacklisted")=1 Then 
						WScript.StdOut.Write "You are blacklisted. Contact info@objectif-securite.ch ,Change your IP,Press [Enter] to continue:"
						WScript.StdIn.ReadLine
						Password=CreakHash(HashStr)
						WriteFile PasswordFileStr,StringN(40,ValueArray(0))&Password					
					ElseIf InStr(Password,"Busy! Try again")=1 Then 
						WriteFile BusyFileStr,Line
					ElseIf Password="" Then
						Password="Your Network is too busy!"
						WriteFile BusyFileStr,Line
					Else
						WriteFile PasswordFileStr,StringN(40,ValueArray(0))&Password
					End If 
					If RunWith="c" Then WScript.Echo HashFile.Line&"."&ValueArray(0)&":"&Password
				End If 
			End If 
		Loop
		WScript.Echo "All Done!"
	End If 
End Sub


'ABA\dbmail(current):1279:001C5A0B859F56BFB01B23448CB7DCDF:4787145233A20D77776E9103971CB4BC
'You are blacklisted. Contact info@objectif-securite.ch



Function CreakHash(HashStr)
	Dim Ret,match
	Ret=HttpPost("http://www.objectif-securite.ch/ophcrack.php","hash=" & HashStr)
	If InStr(Ret,"Busy! Try again in 30 seconds...") Then 
		CreakHash="Busy! Try again in 30 seconds..."
	Else 
		match=RPCFindString(Ret,"<b>Password\:</b></td><td><b>(.+?(?=</b>))")
		If InStr(match,"You are blacklisted")<>0 Then
			CreakHash="You are blacklisted. Contact info@objectif-securite.ch"
		Else
			CreakHash=match
		End If
	End If
End Function 


Function OutPutMsg(MsgStr)
	If RunWith="w" Then 
		InputBox "Password:","Password:",MsgStr
	Else 
		WScript.Echo "Password:"&MsgStr
	End If 
End Function

'---------------------------------------------------------------------------------------
'返回w或者c
Function JudgeRunWith()	
	JudgeRunWith=left(LCase(Right(WScript.FullName,11)),1)
End Function

Function WriteFile(inFileStr,inDataStr)
	Set inFile=fso.OpenTextFile(inFileStr,8,True)
	inFile.WriteLine inDataStr
	inFile.Close
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
	On Error Resume Next 
	Dim http
	Set http=CreateObject("Msxml2.ServerXMLHTTP")
	http.setOption 2,13056	'忽略https错误
	http.open "POST",url,False
	http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	http.send data
	http.waitForResponse 30
	HttpPost = GB2312ToUnicode(http.responseBody)
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
		ret=ret&Match.submatches(0) 
	Next
	RPCFindString=ret	
End Function


'---------------------------------------------------------------------------------
'返回指定长度，超出长度返回原字符串
Function StringN(Num_IWKD,Str_ECBJ)
	If Len(Str_ECBJ)<Num_IWKD Then
		StringN=Str_ECBJ & String(Num_IWKD-Len(Str_ECBJ)," ")
	Else
		StringN=Str_ECBJ
	End If 
End Function