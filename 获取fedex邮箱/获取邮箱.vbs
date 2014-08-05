Option Explicit
Const GETDATE=1
Const GETTITLE=0
Dim WS
Set WS = WScript.CreateObject("Wscript.Shell")
Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
		

Dim oHtml
Set oHtml=CreateObject("htmlfile")
Dim oldClipboardText,Changed,ClipboardText
ClipboardText=oHtml.ParentWindow.ClipboardData.GetData("text")
oldClipboardText=ClipboardText
Do
	ClipboardText=oHtml.ParentWindow.ClipboardData.GetData("text")
	Changed=(ClipboardText<>oldClipboardText)
	oldClipboardText=ClipboardText
	If Changed Then
		DealWith(ClipboardText)		'剪贴板发生变化时的处理
	Else
		wscript.sleep 100
	End If
	ClipboardText=""
Loop
Set ws=Nothing
Set oHtml=Nothing
Set fso=Nothing


Function DealWith(ClipboardTextCP)
	'----------------------------------------------------------------------------
	'写tmp.txt，将剪贴板的内容写入
	Dim File,ret
	Set File=fso.OpenTextFile("tmp.txt",2,True,-1)
	File.Write(ClipboardTextCP)
	File.Close
	
	'----------------------------------------------------------------------------
	'读FilePath.txt，将路径用\分割，改变日期文件夹
	Dim Path,PathArray,i
	Set File=fso.OpenTextFile("FilePath.txt",1,True)
	Path=File.Readline()
	PathArray=Split(Path,"\")
	File.Close
	Path = ""
	For i=0 To UBound(PathArray)-1
		Path=Path & PathArray(i) & "\"
		If Not fso.FolderExists(Path) Then
			fso.CreateFolder(Path)
		End If
	Next
	Path=Path+PathArray(UBound(PathArray))
	If Not fso.FileExists(Path) Then 
		Set File=fso.CreateTextFile(Path,True)
		File.Close
	End If 
	
	'读tmp.txt
	'----------------------------------------------------------------------------
	'写邮件地址
	Dim EmailAddr,File2,Line,n
	n=0
	Set File=fso.OpenTextFile("tmp.txt",1,True,-1)
	Set File2=fso.OpenTextFile(Path,8,True)
	Do Until File.AtEndOfStream
		Line = Trim(File.ReadLine())
		EmailAddr=GetEmail(Line)
		If EmailAddr <> "" Then
			File2.Write EmailAddr
			n=n+1
		End If 
	Loop
	File.Close
	File2.Close
	'----------------------------------------------------------------------------
	'弹出信息提示个数
	'MsgBox "成功!" & vbCrLf & "获取邮箱" & n & "个!",64+4096,"提示:"
	WS.Popup "成功!" & vbCrLf & "获取邮箱" & n & "个!",2,"提示:",64+4096
End Function



'********************************************************************************
'获取每一行中的邮箱
Function GetEmail(Line)
		Dim regEx,Match,Matches,RetStr 
		Set regEx=New RegExp 
		regEx.Pattern= "[A-Za-z0-9_]+([-+.][A-Za-z0-9_]+)*@fedex\.com" 
		regEx.IgnoreCase=True 
		regEx.Global=True 
		Set Matches = regEx.Execute(Line) 
		For Each Match In Matches 
		RetStr = RetStr & Match.Value & vbCrLf 
		Next 
		GetEmail = RetStr
End Function
'********************************************************************************

