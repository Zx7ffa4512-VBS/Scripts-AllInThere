Set ie=WScript.CreateObject("internetexplorer.application","event_") 	'����ie����'
ie.menubar=0 				'ȡ���˵���'
ie.AddressBar=0 			'ȡ����ַ��'
ie.toolbar=0 				'ȡ��������'
ie.statusbar=0 				'ȡ��״̬��'
ie.width=900 				'��400'
ie.height=600 				'��400'
ie.resizable=0 				'�������û��ı䴰�ڴ�С'
ie.navigate "about:blank" 	'�򿪿հ�ҳ��'
ie.left=Fix((ie.document.parentwindow.screen.availwidth-ie.width)/2) 	'ˮƽ����'
ie.top=Fix((ie.document.parentwindow.screen.availheight-ie.height)/2) 	'��ֱ����'
ie.visible=1 				'���ڿɼ�'
'------------------------------------------------------------------------
'�༭����
'------------------------------------------------------------------------
With ie.document 			'���µ���document.write����
	.write "<html><body bgcolor=#999999 scroll=no>"
	.write "<h2 align=center>Post Get����</h2><br>"
	.write "<p>URL��<input id=URL type=text size=100 value=http://dwz.cn/create.php></p>"
	.write "<form><input type=radio checked=checked id=radpost name=fangfa value=ppp >Post "
	.write "<input type=radio name=fangfa id=radget value=ggg > Get </form>"
	.write "<p>Data��<textarea id=textdata align=center cols=80 rows=20 placeholder='url=www.hao123.com'></textarea></p>"
	.write "<p align=center><br>"
	.write "<input id=confirm type=button value=�ύ> "
	.write "<input id=cancel type=button value=ȡ��>"
	.write "</p></body></html>"
End With
'************************************************************************

dim wmi 									'��ʽ����һ��ȫ�ֱ���'
Set wnd=ie.document.parentwindow 			'����wndΪ���ڶ���'
Set IdOrName=ie.document.all 				'����idΪdocument��ȫ������ļ���'


'------------------------------------------------------------------------
'�¼�ӳ��
'------------------------------------------------------------------------
IdOrName.confirm.onclick=GetRef("confirm2") '���õ��"ȷ��"��ťʱ�Ĵ�����'
IdOrName.cancel.onclick=GetRef("cancel") 	'���õ��"ȡ��"��ťʱ�Ĵ�����'

'------------------------------------------------------------------------
'��ѭ��
'------------------------------------------------------------------------
Do While True 								'����ie����֧���¼���������Ӧ�ģ�
	WScript.Sleep 200 						'�ű�������ѭ�����ȴ������¼�.
Loop


'------------------------------------------------------------------------
'�¼��������
'------------------------------------------------------------------------
Sub event_onquit 'ie�˳��¼��������'
	wscript.quit '��ie�˳�ʱ���ű�Ҳ�˳�
End Sub 

Sub cancel '"ȡ��"�¼��������'
	ie.quit '����ie��quit�������ر�IE����'
End Sub '���ᴥ��event_onquit�����ǽű�Ҳ�˳���'


Sub confirm2
	Dim IERet
	Set IERet=CreateObject("Internetexplorer.Application")
	IERet.Navigate "about:blank"
	IERet.visible=1
	If IdOrName.radpost.checked Then 
		IERet.Document.Write HttpPost(IdOrName.URL.value,IdOrName.textdata.value)
	ElseIf IdOrName.radget.checked Then 
		IERet.Document.Write HttpGet(IdOrName.URL.value)
	End If
	Set IERet=Nothing
End Sub 


'------------------------------------------------------------------------
'post get���ð�
'------------------------------------------------------------------------
Function HttpGet(url)
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
		.Type = 2 : .Charset = "utf-8"
		Str = .ReadText : .Close
	End With
	
	Dim RegEx,Match,Matches,SubMatch,ret,ret2
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = "Charset=\x22?(utf-8|unicode|gb2312|gbk)\x22?"
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(Str)
	If Matches.Count<>0 Then JudgeCharset=Matches(0).Submatches(0)
End Function

'------------------------------------------------------------------------
'ת���õ� 
'------------------------------------------------------------------------
Function BytesToStr(Str,charset)
	If charset="" Then charset="utf-8"
	With CreateObject("adodb.stream")
		.Type = 1 : .Open
		.Write Str : .Position = 0
		.Type = 2 : .Charset = charset
		BytesToStr = .ReadText : .Close
	End With
End Function
