'******************************************************************************************************************************
'������־
'v1.0
'1.ʵ�ֻ������
'2.������������

'v1.1
'1.���Ƹ����������ʽ
'2.������������

'v1.2
'1.�Ż�����������ʽ
'2.����������

'v1.3
'1.�Ż�����
'2.����������������bug

'v1.4
'1.���ְλ��

'v2.0
'1.��ı䣬���ܿɷ�����ѡ��
'2.�Ż������������ʽ
'3.�������ܶ����ظ���������ȥ�ع���

'v2.1
'1.�Ż����������Ĵ�������ֱ���õ�����
'	1->2  2->4  3-30  4->32
'2.�Ż�����
'3.�޸�û����������ı���bug

'v2.2
'1.�������ܸ�Ϊ����txt
'


'******************************************************************************************************************************
'����   
'	̽��.txt
'	AllUser.txt 
'	��ȡ̽����Ϣv2.2.vbs
'----------------------------
'	AllUser.txt����Ϊ FirstName|LastName|��˾|ְλ|
'		
'			
'
'******************************************************************************************************************************
'����ʼ

Call main()

Sub Main()
'**********************************************************************************************************************
'��ʼ��Ԥ����Ĺ���	
	Dim [����ѡ��]
	[����ѡ��]=InputBox("��ѡ����"&vbCrLf&vbCrLf&"0:ȫ����"&vbCrLf&vbCrLf&"1:����ȡ̽��"&vbCrLf&vbCrLf&"2:����������","����ѡ��","0")
	Select Case [����ѡ��]
	    Case 0	
	    	ChooseTZ
	    	ChooseSS
	    Case 1	
	    	ChooseTZ
	    Case 2	
	    	ChooseSS
	    Case Else	
	    	MsgBox "�������󣬽ű����˳�",16+4096,"����"
	    	WScript.Quit
	End Select
MsgBox "Successed!"&vbCrLf&"All The Work Has Been Done.",4096,"Congratulations!"
End Sub 

Sub ChooseTZ()
	Const Row1="��	��	��˾	ְλ	����	HostID	NetIP	Referer	Cookies	�豸	ϵͳ	�������ʶ	�����	ɱ�����	Flash	Java	Office	����̽��"
	Dim WS,fso,oHZExcel,TanZhenFile,TableHZ,MySelfPath,Row1Array
	Set WS = WScript.CreateObject("Wscript.Shell")
	Set fso = WScript.CreateObject("Scripting.Filesystemobject") 
	Set oHZExcel = WScript.CreateObject("Excel.Application")
	MySelfPath=fso.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
	
	
	If Not fso.FileExists("̽��.txt") Then 
		WS.Popup "û�ҵ���̽��.txt���ļ�" & vbCrLf & "�ű����˳�!",30,"�ļ�δ�ҵ�",4096+16
		WScript.Quit
	Else
		Set TanZhenFile=fso.OpenTextFile("̽��.txt",1,False)
		oHZExcel.Workbooks.Add
		oHZExcel.ActiveWorkbook.SaveAs MySelfPath & "��Ϣ����"
		Set TableHZ=oHZExcel.Workbooks
		oHZExcel.Cells.Columns.ColumnWidth=50
		oHZExcel.Cells.Columns.RowHeight=30
		oHZExcel.Cells.Columns.HorizontalAlignment = 3
		Row1Array=Split(Row1,"	")
		For i= 0 To UBound(Row1Array)
			oHZExcel.Cells(1,i+1).value=Row1Array(i)
		Next
	End If 
'��ѭ��	
	Dim TmpFile,EachMessage
	Do Until TanZhenFile.AtEndOfStream
		Set TmpFile=fso.OpenTextFile("tmp.txt",8,True)
		DealTanZhen TanZhenFile,TmpFile	'��ȡ̽����Ϣ��tmpfile
		TmpFile.Close
		Set EachMessage=New cMessage
		Set EachMessage.File=fso.OpenTextFile("tmp.txt",1,True)
		EachMessage.GetAllInformation()
		WtTanzhen EachMessage,oHZExcel   '����д��̽������ֹ�˾��Ϣ
		EachMessage.File.Close
		Set EachMessage=Nothing 
		Set TmpFile=fso.OpenTextFile("tmp.txt",2,True)
		TmpFile.Write ""
		TmpFile.Close
	Loop 
	TanZhenFile.Close
	oHZExcel.ActiveWorkBook.Save
	oHZExcel.WorkBooks.Close
	oHZExcel.Quit
	fso.DeleteFile "tmp.txt",True  
	Set WS=Nothing
	Set fso=Nothing
End Sub

Sub ChooseSS()
	Dim WS,fso,oHZExcel,MySelfPath,n,File,Content,Email,DataArray
	Set WS = WScript.CreateObject("Wscript.Shell")
	Set fso = WScript.CreateObject("Scripting.Filesystemobject")
	Set oHZExcel = WScript.CreateObject("Excel.Application")
	MySelfPath=fso.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
	
	If Not fso.FileExists("AllUser.txt") Then
		WS.Popup "û�ҵ���AllUser.txt���ļ�" & vbCrLf & "�ű����˳�!",30,"�ļ�δ�ҵ�",4096+16
		WScript.Quit
	ElseIf Not fso.FileExists("��Ϣ����.xlsx") Then 
		WS.Popup "û�ҵ�����Ϣ����.xlsx���ļ�" & vbCrLf & "�ű����˳�!",30,"�ļ�δ�ҵ�",4096+16
		WScript.Quit
	Else
		Set File=fso.OpenTextFile("AllUser.txt",1,False)
		Content=File.ReadAll
		File.Close 
		oHZExcel.Workbooks.Open(MySelfPath &"��Ϣ����.xlsx")
		n=1
		Do 
			Email=oHZExcel.Cells(n,5).value
			If Trim(Email)="" Then Exit Do 
			If InStr(Content,Email) Then 
					pt="\r\n(.+?"&Email&".+?)(?=\r\n)"
					DataArray=Split(RPCFindString(Content,pt),"|")
					WtCompName DataArray,n,oHZExcel 
			End If
			n=n+1
		Loop
	End If 
	
	Set fso=Nothing 
	Set WS=Nothing 
	oHZExcel.ActiveWorkBook.Save
	oHZExcel.WorkBooks.Close
	oHZExcel.Quit
	Set oHZExcel=Nothing 
End Sub 

Function DealTanZhen(TanZhenFile,TmpFile)
	Dim Line,ret,st,ed,Line2,Return
	Return=False 
	Do Until TanZhenFile.AtEndOfStream
		Line=TanZhenFile.ReadLine
		st=FindString(Line,"^\d+-+$")
		If st<>"" Then
			tmpfile.WriteLine st
			Do Until TanZhenFile.AtEndOfStream
				Line2=TanZhenFile.ReadLine
				ed=FindString(Line2,"^-+$")
				If ed="" Then
					tmpfile.WriteLine Line2
				Else 
					Return=True
					Exit Do
				End If 
			Loop 	
		End If
		If Return Then Exit Do 
	Loop 
End Function

Function WtTanzhen(EachMessage,oHZExcel)
	If EachMessage.Email<>"" Then 
		Dim StartRow,StartColumn
		StartColumn=5
		For i=1 To 65535
			If oHZExcel.Cells(i,StartColumn).value="" Then 
				StartRow=i
				Exit For 
			End If
		Next
		For Each [������] In EachMessage.[Ҫ��д�����Լ���]
			oHZExcel.Cells(StartRow,StartColumn).value=Eval("EachMessage."&[������])
			StartColumn=StartColumn+1
		Next
		oHZExcel.Cells(StartRow,StartColumn).Value=EachMessage.[����̽����Ϣ]
	End If
End Function 

Function WtCompName(DataArray,n,oHZExcel)
		On Error Resume Next 
		oHZExcel.Cells(n,1).value=DataArray(0)
		oHZExcel.Cells(n,2).value=DataArray(1)
		oHZExcel.Cells(n,3).value=DataArray(3)
		oHZExcel.Cells(n,4).value=DataArray(4)
End Function

'-----------------------------------------------------------------------------
'��sSource��sPartnƥ�䣬����ƥ�����ֵ��ÿ��һ��
Function FindString(sSource,sPartn)
	Dim RegEx,Match,Matches,ret
	Set RegEx=New RegExp
	RegEx.MultiLine = True
	RegEx.Pattern = sPartn
	RegEx.IgnoreCase=1
	RegEx.Global=1
	Set Matches=RegEx.Execute(sSource)
	For Each Match In Matches 
		ret = ret + Match.Value
	Next
	FindString = ret
End Function
'-----------------------------------------------------------------------------
'��sSource��sPartnƥ�䣬����ƥ�����ֵ��ÿ��һ��
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



Class cMessage
	Public No,HostID,NetIP,HostInfo,Email,Referer,Cookies,ScanInfo,[���в��]
	Public WS,fso,File
	Public [�������ʶ],[�豸],[ϵͳ],[�����],[ɱ��],Flash,java,Office
	Public [Ҫ��д�����Լ���],[����̽����Ϣ]
	Private Sub Class_Initialize()
		Set WS = WScript.CreateObject("Wscript.Shell")
		Set fso = WScript.CreateObject("Scripting.Filesystemobject")
		No=HostID=NetIP=HostInfo=Email=Referer=Cookies=ScanInfo=""
		[Ҫ��д�����Լ���]=Split("Email,HostID,NetIP,Referer,Cookies,[�豸],[ϵͳ],[�������ʶ],[�����],[ɱ��],Flash,java,Office",",")
	End Sub 
	Private Sub Class_Terminate()
    	No=HostID=NetIP=HostInfo=Email=Referer=Cookies=ScanInfo=[����̽����Ϣ]=""
    	Set WS=Nothing
    	Set fso=Nothing 
    	
	End Sub
	Public Function GetAllInformation()
		Dim Line
		Do Until File.AtEndOfStream
			Line=File.ReadLine
			[����̽����Ϣ]=[����̽����Ϣ] & Line & vbCrLf 
			If InStr(Line,"---") Then 
				No=GetNo(Line)
			ElseIf InStr(Line,"Hostid") Then
				HostID=GetHostID(Line)
			ElseIf InStr(Line,"Netip") Then
				NetIP=GetNetIP(Line)
			ElseIf InStr(Line,"HostInfo") Then
				HostInfo=GetHostInfo(Line)
			ElseIf InStr(Line,"Url") Then
				Email=GetEmail(Line)
			ElseIf InStr(Line,"Referer") Then
				Referer=GetReferer(Line)
			ElseIf InStr(Line,"Cookies") Then
				Cookies=GetCookies(Line)
			ElseIf InStr(Line,"ScanInfo") Then
				ScanInfo=GetScanInfo(Line)
			End If
		Loop
		[���в��]=[��ȡ���в��](ScanInfo)
		[�������ʶ]=[��ȡ�������ʶ](HostInfo)
		[�豸]=[��ȡ�豸](HostInfo)
		[ϵͳ]=[��ȡϵͳ](HostInfo)
		[�����]=[��ȡ�����](HostInfo)
		[ɱ��]=[��ȡɱ��]([���в��])
		Flash=GetFlash([���в��])
		java=GetJava([���в��])
		Office=GetOffice([���в��])
	End Function 
	private Function GetNo(Line)
		GetNo=FindString(Line,"^\d+(?=-)")
	End Function 
	
	private Function GetHostID(Line)
		GetHostID=FindString(Line,"\d+$")
	End Function 
	
	private Function GetNetIP(Line)
		GetNetIP=FindString(Line,"\d+\.\d+\.\d+\.\d+$")
	End Function 
	
	private Function GetHostInfo(Line)
		GetHostInfo=FindString(Line,"mozilla[^\r\n]+?$")
	End Function 
	
	private Function GetEmail(Line)
		GetEmail=FindString(Line,"((\w+([-+.]\w+)*)|(\x22\w+([-+.]\w)*\x22))@((\w+([-.]\w+)*\.\w+([-.]\w+)*)|(\x5B\d+\.\d+\.\d+\.\d+\x5D))")
	End Function 
	
	private Function GetReferer(Line)	
		GetReferer=RPCFindString(Line,":(.+?$)") 
	End Function
	
	private Function GetCookies(Line)
		GetCookies=RPCFindString(Line,":(.+?$)")
	End Function 
	
	private Function GetScanInfo(Line)
		GetScanInfo=RPCFindString(Line,"ScanInfo\s+(.+?$)") 
	End Function
	
	Public Function test()
	End Function
	
	
	
	private Function [��ȡ�������ʶ](HostInfo)
		[��ȡ�������ʶ]=FindString(HostInfo,"mozilla/\d\.\d")
	End Function 
	private Function [��ȡ�豸](HostInfo)
		Dim FunctionRet,tmp(1)
		tmp(0)="\(([^,]+?(?=;|\)))"
		tmp(1)=";\s(\w+?\s\d+?(?=;))"
		For i=0 To UBound(tmp)
			FunctionRet=FunctionRet&RPCFindString(HostInfo,tmp(i))
		Next 
		[��ȡ�豸]=FunctionRet
	End Function 
	private Function [��ȡϵͳ](HostInfo)
		Dim FindRet,FunctionRet
		Dim tmp1(7) 
		tmp1(0)="windows.+?(?=\)|;)"
		tmp1(1)="wow.+?(?=;|\))"
		tmp1(2)="win.+?(?=;|\))"
		tmp1(3)="x\d+?(?=;|\))"
		tmp1(4)="linux(?=;|\))"
		tmp1(5)="android.+?(?=;|\))"
		tmp1(6)="blackberry(?=;|\))"
		tmp1(7)="ubuntu(?=;|\))"
		For i = 0 To UBound(tmp1) 
			 FindRet=FindString(HostInfo,tmp1(i))
			 If FindRet<>"" Then 
			 	FunctionRet = FunctionRet & FindRet & vbCrLf
			 End If 
		Next
		Dim tmp2(0)
		tmp2(0)=";(.+?mac os x.*?(?=;|\)))"
		For i = 0 To UBound(tmp2) 
			 FindRet=RPCFindString(HostInfo,tmp2(i))
			 If FindRet<>"" Then 
			 	FunctionRet = FunctionRet & FindRet & vbCrLf
			 End If 
		Next
		[��ȡϵͳ]=Trim(FunctionRet)
	End Function 
	private Function [��ȡ�����](HostInfo)
		Dim FindRet,tmp2
		Dim tmp(1) 
		tmp(0)="msie.+?;"
		tmp(1)="\s\w+?/\d+(\.\d+)+"
		For i = 0 To UBound(tmp) 
			 FindRet=FindString(HostInfo,tmp(i))
			 If FindRet<>"" Then 
			 	tmp2 = tmp2 & FindRet & vbCrLf
			 End If 
		Next 
		[��ȡ�����]=tmp2
	End Function 
	private Function [��ȡɱ��]([���в��])
		Dim tmp
		For Each item In [���в��]
			If FindString(item,"AVG|McAfee|Avast|Norton|Avira|Bitdefender|Symantec|Kaspersky|Trendmicro|NOD32|MSN|Sophos|Panda|Comodo")<>"" Then
				[��ȡɱ��]=[��ȡɱ��] & item & vbCrLf 
			End If 
		Next 
	End Function 
	private Function GetFlash([���в��])
		For Each item In [���в��]
			If FindString(item,"flash")<>"" Then 
				GetFlash=GetFlash & item & vbCrLf
			End If 
		Next 
	End Function 
	private Function GetJava([���в��])
		For Each item In [���в��]
			If FindString(item,"java")<>"" Then 
				GetJava=GetJava & item & vbCrLf
			End If 
		Next 
	End Function 
	private Function GetOffice([���в��])
		For Each item In [���в��]
			If FindString(item,"office")<>"" Then 
				GetOffice=GetOffice & item & vbCrLf
			End If 
		Next 
	End Function 
	private Function [��ȡ���в��](ScanInfo)
		[��ȡ���в��]=Split(ScanInfo,",")
	End Function 
End Class 
