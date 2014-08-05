'******************************************************************************************************************************
'更新日志
'v1.0
'1.实现基本框架
'2.不可搜索名单

'v1.1
'1.完善各项的搜索方式
'2.不可搜索名单

'v1.2
'1.优化搜索正则表达式
'2.可搜索名单

'v1.3
'1.优化搜索
'2.解决搜索不到报错的bug

'v1.4
'1.添加职位列

'v2.0
'1.大改变，功能可分两项选择
'2.优化邮箱的正则表达式
'3.搜索功能对于重复的邮箱有去重功能

'v2.1
'1.优化对名单表格的处理，可以直接用导出表
'	1->2  2->4  3-30  4->32
'2.优化界面
'3.修改没有总名单表的报错bug

'v2.2
'1.搜索功能改为搜索txt
'


'******************************************************************************************************************************
'需求：   
'	探针.txt
'	AllUser.txt 
'	提取探针信息v2.2.vbs
'----------------------------
'	AllUser.txt必须为 FirstName|LastName|公司|职位|
'		
'			
'
'******************************************************************************************************************************
'程序开始

Call main()

Sub Main()
'**********************************************************************************************************************
'初始化预处理的过程	
	Dim [功能选择]
	[功能选择]=InputBox("请选择功能"&vbCrLf&vbCrLf&"0:全功能"&vbCrLf&vbCrLf&"1:仅提取探针"&vbCrLf&vbCrLf&"2:仅搜索名单","功能选择","0")
	Select Case [功能选择]
	    Case 0	
	    	ChooseTZ
	    	ChooseSS
	    Case 1	
	    	ChooseTZ
	    Case 2	
	    	ChooseSS
	    Case Else	
	    	MsgBox "输入有误，脚本将退出",16+4096,"错误"
	    	WScript.Quit
	End Select
MsgBox "Successed!"&vbCrLf&"All The Work Has Been Done.",4096,"Congratulations!"
End Sub 

Sub ChooseTZ()
	Const Row1="名	姓	公司	职位	邮箱	HostID	NetIP	Referer	Cookies	设备	系统	浏览器标识	浏览器	杀毒软件	Flash	Java	Office	完整探针"
	Dim WS,fso,oHZExcel,TanZhenFile,TableHZ,MySelfPath,Row1Array
	Set WS = WScript.CreateObject("Wscript.Shell")
	Set fso = WScript.CreateObject("Scripting.Filesystemobject") 
	Set oHZExcel = WScript.CreateObject("Excel.Application")
	MySelfPath=fso.GetFile(Wscript.ScriptFullName).ParentFolder & "\"
	
	
	If Not fso.FileExists("探针.txt") Then 
		WS.Popup "没找到【探针.txt】文件" & vbCrLf & "脚本将退出!",30,"文件未找到",4096+16
		WScript.Quit
	Else
		Set TanZhenFile=fso.OpenTextFile("探针.txt",1,False)
		oHZExcel.Workbooks.Add
		oHZExcel.ActiveWorkbook.SaveAs MySelfPath & "信息汇总"
		Set TableHZ=oHZExcel.Workbooks
		oHZExcel.Cells.Columns.ColumnWidth=50
		oHZExcel.Cells.Columns.RowHeight=30
		oHZExcel.Cells.Columns.HorizontalAlignment = 3
		Row1Array=Split(Row1,"	")
		For i= 0 To UBound(Row1Array)
			oHZExcel.Cells(1,i+1).value=Row1Array(i)
		Next
	End If 
'主循环	
	Dim TmpFile,EachMessage
	Do Until TanZhenFile.AtEndOfStream
		Set TmpFile=fso.OpenTextFile("tmp.txt",8,True)
		DealTanZhen TanZhenFile,TmpFile	'提取探针信息到tmpfile
		TmpFile.Close
		Set EachMessage=New cMessage
		Set EachMessage.File=fso.OpenTextFile("tmp.txt",1,True)
		EachMessage.GetAllInformation()
		WtTanzhen EachMessage,oHZExcel   '搜索写入探针和名字公司信息
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
		WS.Popup "没找到【AllUser.txt】文件" & vbCrLf & "脚本将退出!",30,"文件未找到",4096+16
		WScript.Quit
	ElseIf Not fso.FileExists("信息汇总.xlsx") Then 
		WS.Popup "没找到【信息汇总.xlsx】文件" & vbCrLf & "脚本将退出!",30,"文件未找到",4096+16
		WScript.Quit
	Else
		Set File=fso.OpenTextFile("AllUser.txt",1,False)
		Content=File.ReadAll
		File.Close 
		oHZExcel.Workbooks.Open(MySelfPath &"信息汇总.xlsx")
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
		For Each [属性名] In EachMessage.[要填写的属性集合]
			oHZExcel.Cells(StartRow,StartColumn).value=Eval("EachMessage."&[属性名])
			StartColumn=StartColumn+1
		Next
		oHZExcel.Cells(StartRow,StartColumn).Value=EachMessage.[完整探针信息]
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
		ret = ret + Match.Value
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



Class cMessage
	Public No,HostID,NetIP,HostInfo,Email,Referer,Cookies,ScanInfo,[所有插件]
	Public WS,fso,File
	Public [浏览器标识],[设备],[系统],[浏览器],[杀软],Flash,java,Office
	Public [要填写的属性集合],[完整探针信息]
	Private Sub Class_Initialize()
		Set WS = WScript.CreateObject("Wscript.Shell")
		Set fso = WScript.CreateObject("Scripting.Filesystemobject")
		No=HostID=NetIP=HostInfo=Email=Referer=Cookies=ScanInfo=""
		[要填写的属性集合]=Split("Email,HostID,NetIP,Referer,Cookies,[设备],[系统],[浏览器标识],[浏览器],[杀软],Flash,java,Office",",")
	End Sub 
	Private Sub Class_Terminate()
    	No=HostID=NetIP=HostInfo=Email=Referer=Cookies=ScanInfo=[完整探针信息]=""
    	Set WS=Nothing
    	Set fso=Nothing 
    	
	End Sub
	Public Function GetAllInformation()
		Dim Line
		Do Until File.AtEndOfStream
			Line=File.ReadLine
			[完整探针信息]=[完整探针信息] & Line & vbCrLf 
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
		[所有插件]=[获取所有插件](ScanInfo)
		[浏览器标识]=[获取浏览器标识](HostInfo)
		[设备]=[获取设备](HostInfo)
		[系统]=[获取系统](HostInfo)
		[浏览器]=[获取浏览器](HostInfo)
		[杀软]=[获取杀软]([所有插件])
		Flash=GetFlash([所有插件])
		java=GetJava([所有插件])
		Office=GetOffice([所有插件])
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
	
	
	
	private Function [获取浏览器标识](HostInfo)
		[获取浏览器标识]=FindString(HostInfo,"mozilla/\d\.\d")
	End Function 
	private Function [获取设备](HostInfo)
		Dim FunctionRet,tmp(1)
		tmp(0)="\(([^,]+?(?=;|\)))"
		tmp(1)=";\s(\w+?\s\d+?(?=;))"
		For i=0 To UBound(tmp)
			FunctionRet=FunctionRet&RPCFindString(HostInfo,tmp(i))
		Next 
		[获取设备]=FunctionRet
	End Function 
	private Function [获取系统](HostInfo)
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
		[获取系统]=Trim(FunctionRet)
	End Function 
	private Function [获取浏览器](HostInfo)
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
		[获取浏览器]=tmp2
	End Function 
	private Function [获取杀软]([所有插件])
		Dim tmp
		For Each item In [所有插件]
			If FindString(item,"AVG|McAfee|Avast|Norton|Avira|Bitdefender|Symantec|Kaspersky|Trendmicro|NOD32|MSN|Sophos|Panda|Comodo")<>"" Then
				[获取杀软]=[获取杀软] & item & vbCrLf 
			End If 
		Next 
	End Function 
	private Function GetFlash([所有插件])
		For Each item In [所有插件]
			If FindString(item,"flash")<>"" Then 
				GetFlash=GetFlash & item & vbCrLf
			End If 
		Next 
	End Function 
	private Function GetJava([所有插件])
		For Each item In [所有插件]
			If FindString(item,"java")<>"" Then 
				GetJava=GetJava & item & vbCrLf
			End If 
		Next 
	End Function 
	private Function GetOffice([所有插件])
		For Each item In [所有插件]
			If FindString(item,"office")<>"" Then 
				GetOffice=GetOffice & item & vbCrLf
			End If 
		Next 
	End Function 
	private Function [获取所有插件](ScanInfo)
		[获取所有插件]=Split(ScanInfo,",")
	End Function 
End Class 
