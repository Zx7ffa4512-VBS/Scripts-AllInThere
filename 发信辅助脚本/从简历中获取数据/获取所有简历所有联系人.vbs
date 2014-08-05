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
	Dim oExcel,n,Email
	Set oExcel = WScript.CreateObject("Excel.Application")
	oExcel.Workbooks.Open(FullPath)
	n=GetRowNum(oExcel)
	For i=2 To n
		Email=oExcel.Cells(i,6).value
		If Email="" Then
			WriteEmail oExcel.Cells(i,4).value 
		Else
			WriteEmail Email
		End If 	
	Next 
	oExcel.WorkBooks.Close

	oExcel.Quit

	Set oExcel = Nothing 
End Function 


Function GetRowNum(oExcel)
	On Error Resume Next 
	For n=1 To 65535
		If oExcel.Cells(n,2).value="" And oExcel.Cells(n,4).value="" And oExcel.Cells(n,6).value="" Then 
			Exit For 
		End If 
	Next 
	GetRowNum=n
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

Function TestEmail(Email)
	Dim ret
	ret=FindString(Email,"@.*?((\Wgov(\W|$))|(\Wmil(\W|$)))")
	If ret="" Then
		TestEmail=True 
	Else
		TestEmail=False
	End If 
End Function 


Function WriteEmail(Email)
	Dim fso,File
	If TestEmail(Email) Then 
		Set fso = WScript.CreateObject("Scripting.Filesystemobject")
		Set File=fso.OpenTextFile("Email.txt",8,True)
		File.WriteLine Email
		File.Close
	End If 
End Function 



Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
MySelfPath=fso.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
FindCSVFile(MySelfPath) '遍历  
Set fso=Nothing
MsgBox "ok!"