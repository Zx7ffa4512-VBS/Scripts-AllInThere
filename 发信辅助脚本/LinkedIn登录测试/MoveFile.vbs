Dim fso,VarFile,UserName,PassWord,User
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set VarFile=fso.OpenTextFile("var.txt",1,False)
UserName=VarFile.ReadLine
PassWord=VarFile.ReadLine
User=VarFile.ReadLine
VarFile.Close
If Not fso.FolderExists(User) Then 
	fso.CreateFolder(User)
	Dim File
	Set File=fso.OpenTextFile(User&"\�ʺ�����.txt",2,True)
	File.Write UserName & vbCrLf & PassWord
	File.Close
	fso.MoveFile "*.pdf",User
	fso.MoveFile "*.csv",User
	WScript.Quit
Else 
	MsgBox "����ͬ���ļ��У�"
	WScript.Quit
End If