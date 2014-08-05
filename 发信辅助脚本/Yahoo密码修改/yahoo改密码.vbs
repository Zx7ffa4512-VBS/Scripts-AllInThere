Const BTN=0
Const EDT=1
Const SEL=2
Const NEWPASS="LL%*d23@!+t54ZZ"
Dim WS

Set WS = WScript.CreateObject("Wscript.Shell")

Dim fso,File,File2,Line,LineArray
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set File=fso.OpenTextFile("yahoo邮箱.txt",1,False)
Set File2=fso.OpenTextFile("Yahoo修改.txt",8,True)
Do Until File.AtEndOfStream
	Line=File.ReadLine
	LineArray=Split(Line,";")
	Dim IE
	Set IE=CreateObject("Internetexplorer.Application")
	IE.Left=0
	IE.Top=0
	IE.Width=1000
	IE.Height=870
	IE.Navigate "https://edit.yahoo.com/config/change_pw"
	AutoComplete IE,"username",EDT,LineArray(0)
	AutoComplete IE,"passwd",EDT,LineArray(1)
	AutoCompleteById IE,".save",BTN,1
	

	WaitForReady IE,"Enter your current password and then choose your new password"
	AutoComplete IE,"opw",EDT,LineArray(1)
	AutoComplete IE,"newpw1",EDT,NEWPASS
	AutoComplete IE,"newpw2",EDT,NEWPASS
	AutoComplete IE,"saveBtn",BTN,1
	

	WaitForReady IE,"You have successfully chosen a new password"
	IE.Navigate "http://login.yahoo.com/config/login?logout=1&.direct=2&.done=http://www.yahoo.com&.src=ymbr&.intl=us&.lang=en-US"
	

	WaitForReady IE,"Shopping"
	File2.WriteLine LineArray(0)&";"&NEWPASS
	IE.Quit
	Set IE=Nothing
Loop 
File.Close
File2.Close
MsgBox "完成"

Function AutoComplete(IE,sField,nType,sValue)
	IE.Visible=True 
	While IE.Busy Or IE.ReadyState<>4
		WScript.Sleep 300
	Wend
	Eval("IE.Document.All."&sField).Focus
	Select Case nType
	Case BTN
		Eval("IE.Document.All."&sField).Click
	Case EDT
		Eval("IE.Document.All."&sField).value=sValue
	Case SEL
		Eval("IE.Document.All."&sField).value=sValue
	Case Else
		MsgBox "Error!"
	End Select
End Function

Function AutoCompleteById(IE,sFieldID,nType,sValue)
	IE.Visible=True 
	While IE.Busy Or IE.ReadyState<>4
		WScript.Sleep 300
	Wend
	If 1 Then 
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
	Else 
		MsgBox "网页出错,脚本将退出!"
		WScript.Quit
	End If 
End Function 


Function WaitForReady(IE,text)
	On Error Resume Next  
	Dim n
	Do Until InStr(IE.Document.Body.OuterHtml,text)
		WScript.Sleep 500
		n=n+1
		If n=40 Then Exit Do
	Loop
End Function

Function WriteLog(str)
	Dim fso,File
	Set fso = WScript.CreateObject("Scripting.Filesystemobject")
	Set File=fso.OpenTextFile("log.txt",8,True)
	File.WriteLine str
	File.Close
	Set fso=Nothing
End Function 