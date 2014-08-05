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
	Dim HashStr,HashFileStr,IsHash,Ret,Password
	If Len(arg)=65 And InStr(arg,":")=33 Then IsHash=True
	If IsHash Then 
		HashStr=arg
		Password=CreakHash(HashStr)
		OutPutMsg(Password)
		If RunWith="c" Then 	
			Do 
				WScript.StdOut.Write "Hash:"
				HashStr=WScript.StdIn.ReadLine
				Password=CreakHash(HashStr)
				OutPutMsg(Password)
			Loop 
		End If 
	Else
		HashFileStr=arg
		Dim n:n=0
		Dim AllHash
		Set HashFile=fso.OpenTextFile(HashFileStr,1,True)
		Do Until HashFile.AtEndOfStream
			Line=HashFile.ReadLine
			LineArray=Split(Line,":")
			If UBound(LineArray)>=2 Then 
				AllHash=AllHash & StringN(45,LineArray(0),-1) & LineArray(UBound(LineArray)-1) & ":" & LineArray(UBound(LineArray)) & vbCrLf
				n=n+1
			End If 
		Loop
		HashFile.Close
		Set HashFile=fso.OpenTextFile(HashFileStr,2,True)
		HashFile.Write AllHash
		HashFile.Close
		MsgBox "OK" & vbCrLf & "共" & n & "条"
	End If 
End Sub

'
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




Function WriteFile(FileStr,DataStr)
	On Error Resume Next 
	Dim File
	Set File=fso.OpenTextFile(FileStr,8,True)
	File.WriteLine DataStr
	File.Close
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


'------------------------------------------------------------------------
'返回指定长度，超出长度返回原字符串,LeftCenterRight,-1左对齐,0居中,1右对齐
'------------------------------------------------------------------------
Function StringN(Num,Str,LeftCenterRight)
	If LenEx(Str)<Num Then
		Select Case LeftCenterRight
			Case -1 
				StringN=Str & String(Num-Len(Str)," ")
			Case 0
				Dim nYushu,nShang
				nShang=(Num-Len(Str))/2
				nYushu=(Num-Len(Str)) Mod 2
				If nYushu = 0 Then 
					StringN=String(nShang," ") & Str & String(nShang," ")
				Else
					StringN=String(nShang-0.5," ") & Str & String(nShang+0.5," ")
				End If 
			Case 1
				StringN=String(Num-Len(Str)," ") & Str
		End Select
	Else
		StringN=Str
	End If 
End Function

'------------------------------------------------------------------------
'获取字符串长度,中文有效,len("测试")=2,LenEx("测试")=4
'------------------------------------------------------------------------
Function LenEx(Str)
    Dim singleStr,i,iCount
    iCount = 0
    For i = 1 To Len(Str)
        singleStr = Mid(Str,i,1)
        If Asc(singleStr) < 0 Then
            iCount = iCount + 2
        Else 
            iCount = iCount + 1
        End If   
    Next
    LenEx = iCount
End Function