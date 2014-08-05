'------------------------------------------------------------------------
'����
'"[A-Za-z0-9_\x22]+([-+.][A-Za-z0-9_\x22]+)*@((\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\x5B\d+\.\d+\.\d+\.\d+\x5D))"
'
'
'
'
'gUrl="http://tieba.baidu.com/f?kw=%B5%E7%D3%B0%D6%D6%D7%D3"
'gUsefulInformationPattern="[A-Za-z0-9_\x22]+([-+.][A-Za-z0-9_\x22]+)*@((\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\x5B\d+\.\d+\.\d+\.\d+\x5D))"
'gValueFile="tiebaEmail.txt"
'
'
'
'
'
'------------------------------------------------------------------------
Dim gUrl,gUsefulInformationPattern,fso,gValueFile
'********************************************************************************************
gUrl="http://www.jtg-inc.com/"
gUsefulInformationPattern="[A-Za-z0-9_\x22]+([-+.][A-Za-z0-9_\x22]+)*@((\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\x5B\d+\.\d+\.\d+\.\d+\x5D))"
gValueFile="jtg-inc_email.txt"
'********************************************************************************************














Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set page = New PageAnalyze
page.Analyze(gUrl) 

Function WriteFile(FileStr,DataStr)
	Dim File
	Set File=fso.OpenTextFile(FileStr,8,True)
	File.WriteLine DataStr
	File.Close
End Function 

'------------------------------------------------------------------------
'��sSource��sPartnƥ�䣬����ƥ�����ֵ��ÿ��һ��
'------------------------------------------------------------------------
Function FindString(sSource,sPartn)
	Dim RegEx,Match,Matches,ret
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = sPartn
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(sSource)
	For Each Match In Matches 
		ret = ret & Match.Value
	Next
	FindString = ret
End Function

'------------------------------------------------------------------------
'��ƥ�䷴��Ԥ�飬\d(sPartn)��������ƥ��sPartnǰΪ����ֵ
'------------------------------------------------------------------------
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
'��ƥ�䷴��Ԥ�飬\d(sPartn)��������ƥ��sPartnǰΪ����ֵ
'------------------------------------------------------------------------
Function RPCFindStringVbcrlf(sSource,sPartn)
	Dim RegEx,Match,Matches,SubMatch,ret,ret2
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = sPartn
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(sSource)
	For Each Match In Matches 
		ret=ret & Match.submatches(0) & vbCrLf
	Next
	RPCFindStringVbcrlf=ret	
End Function


'------------------------------------------------------------------------
'��sSource��sPartnƥ�䣬����ƥ�����ֵ��ÿ��һ��
'------------------------------------------------------------------------
Function FindStringVbcrlf(sSource,sPartn)
	Dim RegEx,Match,Matches,ret
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = sPartn
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(sSource)
	For Each Match In Matches 
		ret = ret & Match.Value & vbCrLf
	Next
	FindStringVbcrlf = ret
End Function




'********************************************************************************************





'------------------------------------------------------------------------
'url������
'------------------------------------------------------------------------
Class URLAnalyze
	Dim Protocol
	Dim Host
	Dim Port
	Dim Path
	Dim Arg
	Dim Link
	Private Sub Class_Initialize()
        
    End Sub

    Private Sub Class_Terminate()
        
    End Sub
	Public Function Update()
		Protocol=FindString(Link,"http.+//")
		Host=RPCFindString(Link,"//(.+?)(?=[:'\x22 /])")
		Port=FindString(Link,":\d+(?=['\x22 /])"):If Port="" Then Port=":80"
		Path=RPCFindString(Link,Host & Port & "(/.+?)(?=['\x22 \?])")
		Arg=FindString(Link,"\?.+?$")
	End Function
End Class 





'********************************************************************************************





'------------------------------------------------------------------------
'ҳ�������
'------------------------------------------------------------------------
Class PageAnalyze
	Dim Address 
	Dim PageContent
	Dim AllUrl
	
	Dim Pattern
	Dim Value
	
	
	'------------------------------------------------------------------------
	'��ȡ��ҳ����Ч��Ϣ������
	'------------------------------------------------------------------------
	Private Sub Class_Initialize()
        Set AllUrl = New UrlDictionary
        Pattern=gUsefulInformationPattern
    End Sub	
    
	Private Sub Class_Terminate()
        Set UrlArray=Nothing
    End Sub
    
	Public Sub Analyze(url)
		'������ǰurl
		Set Address = New URLAnalyze
		Address.Link = url:Address.Update
		
		'�õ���Χ
		Dim tmp:tmp=Split(Address.Host,".")
		If tmp(UBound(tmp)-1)<>"com" Then 
			AllUrl.FilterStr=tmp(UBound(tmp)-1) & "." & tmp(UBound(tmp))
		Else
			AllUrl.FilterStr=tmp(UBound(tmp)-2) & "." & tmp(UBound(tmp)-1) & "." & tmp(UBound(tmp))
		End If 
		'��ȡҳ��
		WScript.Echo "���ڻ�ȡ... >> " & url
		PageContent=HttpGet(url)
		'�����ɨ���¼
		AllUrl.AddHadBeenScaned url
		'��Ӵ�ɨ���¼
		GetUrl(url)
		'��ȡ��ǰҳ��������Ϣ�����
		Value=FindStringVbcrlf(PageContent,Pattern)
		WScript.Echo Value
		WriteFile gValueFile,Value 
		For i=0 To AllUrl.GoingToScan.Count-1
			DoLoop(AllUrl.GoingToScan.item(i))
			AllUrl.GoingToScan.Remove(i)
		Next 
	End Sub 
	
	Private Sub DoLoop(url)
		'������ǰurl
		Set Address = New URLAnalyze
		Address.Link = url:Address.Update
		
		WScript.Echo "���ڻ�ȡ... >> " & url
		PageContent=HttpGet(url)
		'�����ɨ���¼
		AllUrl.AddHadBeenScaned url
		'��Ӵ�ɨ���¼
		GetUrl(url)
		'��ȡ��ǰҳ��������Ϣ�����
		Value=FindStringVbcrlf(PageContent,Pattern)
		WScript.Echo Value
		WriteFile gValueFile,Value 
	End Sub 
	
		
	'------------------------------------------------------------------------
	'��ȡҳ���е�����,��������   '(?<=href=['\x22])((?![#;]).*?)(?=['\x22])
	'------------------------------------------------------------------------
	Private Sub GetUrl(url)
		'��ȡҳ����url
		Dim Ret : Ret=Split(RPCFindStringVbcrlf(PageContent,"href=['\x22 ]((?![#;]).*?)(?=['\x22 #])"),vbCrLf)
		'�ų�#��javascript: void(0) ����������
		For Each u In Ret
			If Left(u,1)="/" And Left(u,2)<>"//" Then
				AllUrl.AddGoingToScan(Address.Protocol & Address.Host & Address.Port & u)
			ElseIf Left(LCase(u),4)="http" Then
				AllUrl.AddGoingToScan(u)
			ElseIf InStr(LCase(u),"javascript:")=0 And InStr(LCase(u),"void(0)")=0 Then 
				AllUrl.AddGoingToScan(Address.Protocol & Address.Host & Address.Port & "/" & u)
			End If
		Next 	
	End Sub 
	
	'------------------------------------------------------------------------
	'post get���ð�
	'------------------------------------------------------------------------
	Function HttpGet(url)
		On Error Resume Next 
		Dim http,Cs,responseStr
		Set http=CreateObject("Msxml2.ServerXMLHTTP")
		http.setOption 2,13056	'����https����
		http.open "GET",url,False
		http.send
		http.waitForResponse 50
		Cs=JudgeCharset(http.responseBody)
		HttpGet = BytesToStr(http.responseBody,Cs)
	End Function 
	
	Function HttpPost(url,data)
		Dim http
		Set http=CreateObject("Msxml2.ServerXMLHTTP")
		http.setOption 2,13056	'����https����
		http.open "POST",url,False
		http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		http.send data
		http.waitForResponse 50
		Cs=JudgeCharset(http.responseBody)
		HttpPost = BytesToStr(http.responseBody,Cs)
	End Function 
	
	
	'------------------------------------------------------------------------
	'�ж��ַ�����
	'------------------------------------------------------------------------
	Function JudgeCharset(sSource)
		Dim Str
		With CreateObject("adodb.stream")
			.Type = 1 : .Open
			.Write sSource : .Position = 0
			.Type = 2 : .Charset = "gb2312"
			Str = .ReadText : .Close
		End With
		
		Dim RegEx,Match,Matches,SubMatch,ret,ret2
		Set RegEx=New RegExp
		RegEx.MultiLine = True
		RegEx.Pattern = "Charset=[\x22' ]?(.+?)(?=[\x22' ])"
		RegEx.IgnoreCase=1
		RegEx.Global=1
		Set Matches=RegEx.Execute(Str)
		If Matches.Count<>0 Then 
			JudgeCharset=Replace(Matches(0).Submatches(0),"'","")
			JudgeCharset=Replace(Matches(0).Submatches(0),Chr(34),"")
			JudgeCharset=Replace(Matches(0).Submatches(0)," ","")
		End If
	End Function
	
	'------------------------------------------------------------------------
	'ת���õ� 
	'------------------------------------------------------------------------
	Function BytesToStr(Str,charset)
		If charset="" Then charset=InputBox("δ�������,�ֶ�����:","����","GB2312")
		With CreateObject("adodb.stream")
			.Type = 1 : .Open
			.Write Str : .Position = 0
			.Type = 2 : .Charset = charset
			BytesToStr = .ReadText : .Close
		End With
	End Function
End Class 
'********************************************************************************************


Class UrlDictionary
	Dim GoingToScan
	Dim HadBeenScaned
	Dim GTSNum
	Dim HBSNum
	Dim FilterStr
	Private Sub Class_Initialize()
		GTSNum=0
		HBSNum=0
        Set GoingToScan = CreateObject("scripting.dictionary")
        Set HadBeenScaned = CreateObject("scripting.dictionary")
    End Sub

    Private Sub Class_Terminate()
        Set GoingToScan = Nothing 
        Set HadBeenScaned = Nothing 
    End Sub
	
	Public Function AddGoingToScan(UrlStr)
		Dim Scaned:Scaned=False
		For Each j In HadBeenScaned
			If UrlStr=HadBeenScaned.item(j) Then Scaned=True
		Next
		If RangeScan(UrlStr) And Not Scaned  Then 
			GoingToScan.add GTSNum,UrlStr
			GTSNum=GTSNum+1
		End If 
	End Function 
	
	Public Function AddHadBeenScaned(UrlStr)
		HadBeenScaned.Add HBSNum,UrlStr
		HBSNum=HBSNum+1
	End Function 
	
	Public Function RangeScan(Str)
		RangeScan = InStr(Str,FilterStr)
	End Function 
End Class
