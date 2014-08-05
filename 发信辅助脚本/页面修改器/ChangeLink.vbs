Dim WS
Set WS = WScript.CreateObject("Wscript.Shell")
Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Dim File,Line,Param(4),n,AllContent
Set File=fso.OpenTextFile("param.txt",1,False)
n=0
Do Until File.AtEndOfStream
	Line=File.ReadLine
	If Trim(Line)<>"" Then 
		Param(n)=Trim(Line)
		n=n+1
	End If
Loop
For i=0 To 3
	If Param(i)="" Then MsgBox "Quit!"
Next
File.Close
Set File=fso.OpenTextFile("1.html",1,False)
AllContent=File.ReadAll
File.Close
AllContent=Replace(AllContent,"Var1Var1Var1",Param(0))
AllContent=Replace(AllContent,"Var2Var2Var2",Param(1))
AllContent=Replace(AllContent,"Var3Var3Var3",Param(2))
Set File=fso.OpenTextFile("mail.html",2,True)
File.Write AllContent
File.Close

Set File=fso.OpenTextFile("2.htm",1,False)
AllContent=File.ReadAll
File.Close
AllContent=Replace(AllContent,"Var4Var4Var4",Param(3))
Set File=fso.OpenTextFile("linkedin.htm",2,True)
File.Write AllContent
File.Close
