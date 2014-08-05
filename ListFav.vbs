'Option Explicit
On Error Resume Next
Dim sPath
Const FAVORITES = &H6&

Function FilesTree(sPath)
	Dim objShell,objFso,objFolder,objItem,urlFile,urlLine
	Set objShell	= CreateObject("Shell.Application")
	set objFso		= CreateObject("Scripting.FileSystemObject")
	Set objFolder	= objShell.Namespace(sPath)
	For Each objItem in objFolder.Items
		If objItem.IsFolder	Then
			'Check Folder
			'WScript.Echo objItem.Path
			FilesTree(objItem.Path)
		ElseIf objItem.IsLink Then
			'Check File Lnk
			WScript.Echo objItem.Getlink.Path
		ElseIf  objItem.IsFileSystem And GetExtensionName(objItem.Name) = "url" Then
			'Check File Url
			'WScript.Echo objItem.Path
			set urlFile = objFso.OpenTextFile(objItem.Path ,1)
			Do While urlFile.AtEndOfStream <> True
				urlLine = urlFile.ReadLine
				If Lcase(Left(urlLine,4))="url=" Then
					urlLine = Right(urlLine,Len(urlLine)-4)
					WScript.Echo urlLine
					Exit Do
				End If 
			Loop
		End If
	Next
End Function


if (lcase(right(wscript.fullname,11))="wscript.exe") Then
   Dim objShell
   set objShell = WScript.createObject("wscript.shell")
   objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&wscript.scriptfullname&chr(34))
   WScript.quit
end if

if wscript.arguments.count<1 then
	FilesTree(FAVORITES)
Else
	sPath = Wscript.arguments(0)
	FilesTree(sPath)
End If