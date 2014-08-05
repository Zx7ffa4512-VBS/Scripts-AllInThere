'------------------------------------------------------------------------
'vbs加密工具
'用法1：拖动vbs文件到工具上
'用法2：cscript vbs最强加密_v2.0.vbs 【文件或文件夹】
'------------------------------------------------------------------------

RunWithCscript(WScript.Arguments.Count)
Dim fso,Files
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count = 0 Then
    arg=fso.GetParentFolderName(WScript.ScriptFullName)
ElseIf WScript.Arguments.Count=1 Then 
	arg=FindString(WScript.Arguments(0),"([a-zA-Z]\:\\)?([^\\\/\:\*\?\<\>\|\x0d\x0a]+\\?)+")
End If


If fso.FileExists(arg) Then
	Randomize
	pass = Int(Rnd*12)+20 '异或加密有效范围20-31，所以随机生成好了。
	main arg
ElseIf fso.FolderExists(arg) Then 
	Files=GetAllFiles(arg)
	For Each F In Files
		Randomize
		pass = Int(Rnd*12)+20 '异或加密有效范围20-31，所以随机生成好了。
		If f<>WScript.ScriptFullName Then main F
	Next 
Else
	WScript.Echo "文件或文件夹不存在!"
	Call Usage
	WScript.Quit
End If 



Function main(PathStr)
	Dim SouFile,DesFile,data
	Set SouFile=fso.OpenTextFile(PathStr,1)
	data = SouFile.ReadAll
	SouFile.Close
	data = "d=" & Chr(34) & ASCdata(data) & Chr(34)
	data = data & vbCrLf & ":M=Split(D):For each O in M:N=N&chr(O):Next:execute N"
	data = Replace(data, " ", ",")
	Set DesFile=fso.OpenTextFile(PathStr & "_加密.vbe", 2, True)
	DesFile.Write Encoder(EncHexXorData(data))
	DesFile.Close
	WScript.Echo "加密完毕,文件生成到：" & PathStr & "_加密.vbe"& vbCrLf
End Function 



Sub RunWithCscript(ArgCount)
	If (LCase(Right(WScript.FullName,11))="wscript.exe") Then 
		Set objShell=WScript.CreateObject("wscript.shell")
		If ArgCount=0 Then objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&WScript.ScriptFullName&chr(34))
		If ArgCount=1 Then objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&WScript.ScriptFullName&chr(34)&" "&WScript.Arguments(0))
		WScript.Quit
	End If
End Sub 





Function EncHexXorData(data)
    EncHexXorData = "x=""" & EncHexXor(data) & """:For i=1 to Len(x) Step 2:s=s&Chr(CLng(""&H""&Mid(x,i,2)) Xor " & pass & "):Next:Execute Replace(s,"","","" "")"
End Function

Function Encoder(data) '加密3
    Encoder = CreateObject("Scripting.Encoder").EncodeScriptFile(".vbs", data, 0, "VBScript")
End Function

Function EncHexXor(x) '加密2
    For i = 1 To Len(x)
        EncHexXor = EncHexXor & Hex(Asc(Mid(x, i, 1)) Xor pass)
    Next
End Function

Function ASCdata(Data) '加密1
    num = Len(data)
    newdata = ""
    For j = 1 To num
        If j = num Then
            newdata = newdata&Asc(Mid(data, j, 1))
        Else
            newdata = newdata&Asc(Mid(data, j, 1)) & " "
        End If
    Next
    ASCdata = newdata
End Function


Function GetAllFiles(FolderStr)
	Set folderX=fso.GetFolder(FolderStr)
    Set subFiles=FolderX.Files
    For Each subFile In subFiles
    	If fso.GetExtensionName(subfile)="vbs" Then 
	    	If Tmp="" Then 
	    		Tmp=subFile.Path
	    	Else 
	    		Tmp=Tmp&vbCrLf&subFile.Path
	    	End If
	    End If 
    Next
    GetAllFiles=Split(Tmp,vbCrLf)
End Function


Sub Usage()
	WScript.Echo String(79,"*")
	WScript.Echo "Usage:"
	WScript.Echo "cscript "&chr(34)&WScript.ScriptFullName&chr(34)&" [File OR Folder]"
	WScript.Echo String(79,"*")
End Sub 

'-----------------------------------------------------------------------------
'将sSource用sPartn匹配，返回匹配出的值，每个一行
Function FindString(sSource,sPartn)
	Dim RegEx,Match,Matches,ret
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = sPartn
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(sSource)
	For Each Match In Matches 
		FindString = Match.Value
	Next
End Function


