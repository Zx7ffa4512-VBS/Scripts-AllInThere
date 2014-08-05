Set ie=WScript.CreateObject("internetexplorer.application","event_") 	'创建ie对象'
ie.menubar=0 				'取消菜单栏'
ie.AddressBar=0 			'取消地址栏'
ie.toolbar=0 				'取消工具栏'
ie.statusbar=0 				'取消状态栏'
ie.width=900 				'宽400'
ie.height=600 				'高400'
ie.resizable=0 				'不允许用户改变窗口大小'
ie.navigate "about:blank" 	'打开空白页面'
ie.left=Fix((ie.document.parentwindow.screen.availwidth-ie.width)/2) 	'水平居中'
ie.top=Fix((ie.document.parentwindow.screen.availheight-ie.height)/2) 	'垂直居中'
ie.visible=1 				'窗口可见'
'------------------------------------------------------------------------
'编辑界面
'------------------------------------------------------------------------
With ie.document 			'以下调用document.write方法
	.write "<html><body bgcolor=#999999 scroll=no>"
	.write "<h2 align=center>Post Get工具</h2><br>"
	.write "<p>URL：<input id=URL type=text size=100 value=http://dwz.cn/create.php></p>"
	.write "<form><input type=radio checked=checked id=radpost name=fangfa value=ppp >Post "
	.write "<input type=radio name=fangfa id=radget value=ggg > Get </form>"
	.write "<p>Data：<textarea id=textdata align=center cols=80 rows=20 placeholder='url=www.hao123.com'></textarea></p>"
	.write "<p align=center><br>"
	.write "<input id=confirm type=button value=提交> "
	.write "<input id=cancel type=button value=取消>"
	.write "</p></body></html>"
End With
'************************************************************************

dim wmi 									'显式定义一个全局变量'
Set wnd=ie.document.parentwindow 			'设置wnd为窗口对象'
Set IdOrName=ie.document.all 				'设置id为document中全部对象的集合'


'------------------------------------------------------------------------
'事件映射
'------------------------------------------------------------------------
IdOrName.confirm.onclick=GetRef("confirm2") '设置点击"确定"按钮时的处理函数'
IdOrName.cancel.onclick=GetRef("cancel") 	'设置点击"取消"按钮时的处理函数'

'------------------------------------------------------------------------
'死循环
'------------------------------------------------------------------------
Do While True 								'由于ie对象支持事件，所以相应的，
	WScript.Sleep 200 						'脚本以无限循环来等待各种事件.
Loop


'------------------------------------------------------------------------
'事件处理过程
'------------------------------------------------------------------------
Sub event_onquit 'ie退出事件处理过程'
	wscript.quit '当ie退出时，脚本也退出
End Sub 

Sub cancel '"取消"事件处理过程'
	ie.quit '调用ie的quit方法，关闭IE窗口'
End Sub '随后会触发event_onquit，于是脚本也退出了'


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
'post get常用版
'------------------------------------------------------------------------
Function HttpGet(url)
	Dim http,Cs,responseStr
	Set http=CreateObject("Msxml2.ServerXMLHTTP")
	http.setOption 2,13056	'忽略https错误
	http.open "GET",url,False
	http.send
	http.waitForResponse 50
	Cs=JudgeCharset(http.responseBody)
	HttpGet = BytesToStr(http.responseBody,Cs)
End Function 

Function HttpPost(url,data)
	Dim http
	Set http=CreateObject("Msxml2.ServerXMLHTTP")
	http.setOption 2,13056	'忽略https错误
	http.open "POST",url,False
	http.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	http.send data
	http.waitForResponse 50
	Cs=JudgeCharset(http.responseBody)
	HttpPost = BytesToStr(http.responseBody,Cs)
End Function 


'------------------------------------------------------------------------
'判断字符编码
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
'转码用的 
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
