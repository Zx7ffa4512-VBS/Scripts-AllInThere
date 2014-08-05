Const BTN=0
Const EDT=1
Const SEL=2

Dim fso,File,File2,Line
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set File=fso.OpenTextFile("num.txt",1,True)
Line = File.ReadLine
File.Close
Set File2=fso.OpenTextFile("num.txt",2,True)
File2.Write Line+1
File2.Close

Dim IE
Set IE=CreateObject("Internetexplorer.Application")
IE.Left=0
IE.Top=0
IE.Width=1000
IE.Height=870

IE.Navigate "https://signup.live.com/signup.aspx?wa=wsignin1.0&rpsnv=11&ct=1380018686&rver=6.1.6206.0&wp=MBI&wreply=http%3a%2f%2fmail.live.com%2fdefault.aspx&id=64855&cbcxt=mai&snsc=1&bk=1380018687&uiflavor=web&mkt=ZH-CN&lc=2052&lic=1"
AutoComplete IE,"iLastName",EDT,"kerri"
AutoComplete IE,"iFirstName",EDT,"Green"
AutoComplete IE,"iBirthYear",EDT,"1975"
AutoComplete IE,"iBirthMonth",EDT,"9"
AutoComplete IE,"iBirthDay",EDT,"15"
AutoComplete IE,"iGender",EDT,"f"
AutoComplete IE,"imembernamelive",EDT,"kerriGreen" & Line
AutoComplete IE,"iPwd",EDT,"Zz123456!@#"
AutoComplete IE,"iRetypePwd",EDT,"Zz123456!@#"
'AutoComplete IE,"iPhone",EDT,"13569852365"
AutoComplete IE,"iAltEmail",EDT,"vincercbk@yahoo.com"
AutoComplete IE,"iqsaswitch",BTN,"Susan"
AutoComplete IE,"iSQ",EDT,"第一个宠物的名字"
AutoComplete IE,"iSA",EDT,"Susan"

AutoComplete IE,"iZipCode",EDT,"110101"


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