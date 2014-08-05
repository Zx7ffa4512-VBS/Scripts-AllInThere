Function FindCSVFile(sPath)
	Set oFso = CreateObject("Scripting.FileSystemObject")
	Set oFolder = oFso.GetFolder(sPath)
	Set oSubFolders = oFolder.Subfolders

	Set oFiles = oFolder.Files
	For Each oFile In oFiles
		If oFile.Type="Microsoft Office Excel 逗号分隔值文件" Then 
			GetCSVEmail oFile.Path
		End If 
	Next
	
	For Each oSubFolder In oSubFolders
		FindCSVFile(oSubFolder.Path)'递归
	Next
	
	Set oFolder = Nothing
	Set oSubFolders = Nothing
	Set oFso = Nothing
End Function

Function GetCSVEmail(FullPath)
	Dim fso,File,Content
	Set fso = WScript.CreateObject("Scripting.Filesystemobject")
	fso.CopyFile FullPath,Left(FullPath,Len(fullpath)-3)&"txt",True 
	Set File=fso.OpenTextFile(Left(FullPath,Len(fullpath)-3)&"txt",1,True)
	File.ReadLine
	Content=File.ReadAll
	File.Close
	fso.DeleteFile Left(FullPath,Len(fullpath)-3)&"txt",True 
	Set File=fso.OpenTextFile("AllUser.txt",8,True)
	File.Write Content 
	File.Close
End Function 
Function DealAllUser()
	Dim fso,File,Line,File2
	Set fso = WScript.CreateObject("Scripting.Filesystemobject")
	Set File=fso.OpenTextFile("AllUser.txt",1,False)
	Set File2=fso.OpenTextFile("tmp.txt",8,True )
	Do Until File.AtEndOfStream
		Line=Split(File.ReadLine,chr(34)&","&Chr(34))
		For Each l In Line 
			If l<>"" And l<>Chr(34) Then File2.Write l&"|"
		Next
		File2.Write vbCrLf 
	Loop 
	File.Close
	File2.Close
	Set File=fso.OpenTextFile("AllUser.txt",2,False)
	Set File2=fso.OpenTextFile("tmp.txt",1,True )
	File.Write File2.ReadAll
	File.Close
	File2.Close
	fso.DeleteFile "tmp.txt",True 
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
		ret = ret & Match.Value & vbcrlf 
	Next
	FindString = ret
End Function
'-----------------------------------------------------------------------------
'将sSource用sPartn匹配，返回匹配出的值，每个一行
Function RPCFindString(sSource,sPartn)
	Dim RegEx,Match,Matches,SubMatch,ret,ret2
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = sPartn
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(sSource)
	For Each Match In Matches 
		ret=ret&Match.submatches(0)&vbCrLf 
	Next
	RPCFindString=ret	
End Function



Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
MySelfPath=fso.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
FindCSVFile(MySelfPath) '遍历  
Set fso=Nothing
DealAllUser()
MsgBox "ok!"