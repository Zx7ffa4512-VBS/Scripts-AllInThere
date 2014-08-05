Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Dim WS
Set WS = WScript.CreateObject("Wscript.Shell")
Dim FileStr
If WScript.Arguments.Count=1 Then 
	FileStr=WScript.Arguments(0)
Else
	FileStr="password.txt"
End If 

Dim File,OldLine,NewLine(65535),OldContent,AllNum,LeaveNum
Set File=fso.OpenTextFile(FileStr,1,False)
Do Until File.AtEndOfStream
	OldContent=File.ReadAll
Loop
File.Close
OldLine=Split(OldContent,vbCrLf)
AllNum=UBound(OldLine)
Dim bWrite,nWrite
bWrite=True 
nWrite=0 
For i=0 To UBound(OldLine)-1
	For j=0 To i
		If NewLine(j)=OldLine(i) Then 
			bWrite=False
			Exit For	
		End If 
	Next 
	If bWrite Then
		NewLine(nWrite)=OldLine(i)
		nWrite=nWrite+1
	End If 
	bWrite=True
Next
File.close

n=0
Set File=fso.OpenTextFile(FileStr,2,False)
Do Until Trim(NewLine(n))=""
	File.WriteLine NewLine(n)
	n=n+1
Loop 
LeaveNum=n
File.Close
MsgBox "去重完成!"&vbCrLf&"共有条目:"&AllNum&vbCrLf&"去除重复条目:"&AllNum-LeaveNum&vbCrLf&"剩余条目:"&LeaveNum,4096,"完成"