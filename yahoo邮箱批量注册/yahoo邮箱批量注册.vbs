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
IE.Navigate "https://edit.yahoo.com/registration?.intl=us&.lang=en-US&.pd=ym_ver%253D0%2526c%253D%2526ivt%253D%2526sg%253D&new=1&.done=http%3A//mail.yahoo.com&.src=ym&.v=0&.u=9qu17cp8uar9r&partner=&.partner=&pkg=&stepid=&.p=&promo=&.last="
AutoComplete IE,"firstname",EDT,"Smile"
AutoComplete IE,"secondname",EDT,"Open"
AutoComplete IE,"yahooid",EDT,"SmileOpen" + Line 
AutoComplete IE,"password",EDT,"Zz123456!@#"
AutoComplete IE,"passwordconfirm",EDT,"Zz123456!@#"
AutoComplete IE,"mm",SEL,"3"
AutoComplete IE,"dd",EDT,"12"
AutoComplete IE,"yyyy",EDT,"1967"
AutoComplete IE,"gender",SEL,"f"
AutoComplete IE,"language",SEL,"en-US"
AutoComplete IE,"country",SEL,"us"
AutoComplete IE,"postalcode",EDT,"98101"
'AutoComplete IE,"IAgreeBtn",BTN,"1"
MsgBox "等待页面刷新..." + vbCrLf + "成功请点[确定]",4096+64,"IE.ReadyState"

AutoComplete IE,"secquestion",SEL,"In which city did you study abroad?"
AutoComplete IE,"secquestionanswer",EDT,"asdf"
AutoComplete IE,"secquestion2",SEL,"What was the last name of your favorite teacher?"
AutoComplete IE,"secquestionanswer2",EDT,"asdfasdf"





Function AutoComplete(IE,sFieldID,nType,sValue)
	IE.Visible=True 
	While IE.Busy Or IE.ReadyState<>4
		WScript.Sleep 50
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