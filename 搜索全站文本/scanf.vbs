'On Error Resume Next 
RunWithCscript(WScript.Arguments.Count)
If WScript.Arguments.Count = 2 Then 
	Path=WScript.Arguments(0)
	ExportFile=WScript.Arguments(1)
Else 
	WScript.StdOut.Write "Enter Start Path( Like:c:\ ):"
	Path=WScript.StdIn.Readline()
	WScript.StdOut.Write "Enter ExportFile( Like:Export.TXT ):"
	ExportFile=WScript.StdIn.Readline()
End If
If Path="" Then Call Usage : WScript.Quit
'--------------------------------------------------------------------------------
Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set fod=fso.GetFolder(Path)
fnTraversalFolder(fod)      '遍历文件夹
'-------------------------------------------------------------------------------------------------------
Sub RunWithCscript(ArgCount)
	If (LCase(Right(WScript.FullName,11))="wscript.exe") Then 
		Set objShell=WScript.CreateObject("wscript.shell")
		If ArgCount=0 Then 
			objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&WScript.ScriptFullName&chr(34))
		Else
			Dim argTmp
			For Each arg In WScript.Arguments
				argTmp=argTmp&arg&" "
			Next 
			objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&WScript.ScriptFullName&chr(34)&" "&argTmp)
		End If
		WScript.Quit
	End If
End Sub

Sub Usage()
	WScript.Echo String(79,"*")
	WScript.Echo "Usage:"
	WScript.Echo "cscript "&Chr(34)&WScript.ScriptFullName&Chr(34)&" FolderPath ExportFile"
	WScript.Echo String(79,"*")&vbCrLf 
End Sub


Function WriteFile(FileStr,DataStr)
	Set File=fso.OpenTextFile(FileStr,8,True)
	File.WriteLine DataStr
	File.Close
End Function 
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
		ret = ret & Match.Value
	Next
	FindString = ret
End Function

'---------------------------------------------------------------------
'文件夹对象
Function fnTraversalFolder(FolderObj)
	Dim SubFolders,Folder,Files,File,Exten
    Set SubFolders=FolderObj.SubFolders 
    For Each Folder In SubFolders 
        fnTraversalFolder(Folder) 
    Next
    Set Files=FolderObj.Files 
    For Each File In Files 
    	Exten=fso.GetExtensionName(File)
    	If FindString(Exten,"txt|bat|ini|inf|htm|html|asp|aspx|jsp|php|vbs|js|hta|config")<>"" Then
			fnReadFile(File.Path)
		End If
    Next
End Function 


Function fnReadFile(FileStr)
	Dim File,Line,ArrayLine(4)   '     (此数+2)/2=3
	Set File=fso.OpenTextFile(FileStr,1,True)
	Do Until File.AtEndOfStream
		Line = File.ReadLine
		For i = 3 To 0 Step -1 
			ArrayLine(i+1)=ArrayLine(i)	
		Next
		ArrayLine(0)=Line
		If fnWhetherWrite(ArrayLine(2)) Then
			If Len(FileStr)<80 Then WriteFile ExportFile,String((80-Len(FileStr))/2,"*")&FileStr&String((80-Len(FileStr))/2,"*")
			If Len(FileStr)>=80 Then WriteFile ExportFile,FileStr
			For k=4 To 0 Step -1
				WriteFile ExportFile,ArrayLine(k)
			Next
			WriteFile ExportFile,String(79,"*")&vbCrLf&vbCrLf&vbCrLf&vbCrLf&vbCrLf
		End If 
	Loop 
End Function


Function fnWhetherWrite(Str)
	If FindString(str,"(mysql_connect\s{0,}\(""\w+?"",""\w+?"",""\w+?""\)\s{0,};)|(\$cfg\['Servers'\]\[\$i\]\['(password|user)'\]\s+?=.+?;)")<>"" Then 
		WScript.Echo Str 
		fnWhetherWrite=True
	Else 
		fnWhetherWrite=False
	End If 
End Function 

'conn     		"mysql_connect\s{0,}\(""\w+?"",""\w+?"",""\w+?""\)\s{0,};"
'phpmyadmin		"\$cfg\['Servers'\]\[\$i\]\['(password|user)'\]\s+?=.+?;"
'
'