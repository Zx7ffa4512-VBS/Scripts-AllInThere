On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=true", , 48)
For Each objItem in colItems
     msg = msg & "编号：" & objItem.Index & vbCrLf & "MAC：" & objItem.MACAddress & vbCrLf & "网卡：" & objItem.Description & vbCrLf & vbCrLf
     ix = ix & objItem.Index & ","
     MAC = objItem.MACAddress
Next
If msg = "" Then
     MsgBox "未找到活动网卡,单击确定退出程序",64,"MAC修改程序"
     wscript.Quit
End If
Do
     idx = InputBox( msg , "请输入您要修改MAC的网卡编号", Left(ix,InStr(ix,",") - 1))
     if idx = False Then Wscript.Quit
Loop Until IsNumeric(idx) And InStr(ix,idx)
Do
     MAC = InputBox( "输入您指定的MAC地址值(注意应该是12位的连续数字或字母(A~F)，其间没有-、：等分隔符)" , "输入新的MAC地址", MAC)
     if MAC = False Then Wscript.Quit
     MAC = Replace(Replace(Replace(MAC, ":", ""), "-", ""), " ", "")
loop until rt("^[\da-fA-F]{12}$",MAC)

idx = Right("00000" & idx, 4)
reg = "HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}\" & idx
Set WSHShell = CreateObject("WScript.Shell")
WshShell.RegWrite reg & "\NetworkAddress", MAC , "REG_SZ"
WshShell.RegWrite reg & "\Ndi\params\NetworkAddress\default" , MAC , "REG_SZ"
WshShell.RegWrite reg & "\Ndi\params\NetworkAddress\ParamDesc" , "NetworkAddress" , "REG_SZ"
WshShell.RegWrite reg & "\Ndi\params\NetworkAddress\optional" , "1" , "REG_SZ"
'得到网卡的名称
NetWorkName = WshShell.RegRead("HKLM\SYSTEM\ControlSet001\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & WshShell.RegRead(reg & "\NetCfgInstanceId") & "\Connection\Name")
WSHShell.Popup "程序将重启您的网卡,请稍后...",2,"MAC修改程序",64
restartNetWork NetWorkName
If Not Err Then WSHShell.Popup "修改成功!" & vbcrlf & vbcrlf & "如果出现异常,请打开设备管理器,展开网络适配器," & vbcrlf & "在相应的网卡上右击,选择属性" & _
"切换到高级选项卡," & vbcrlf & "将属性NetworkAddress值修改为不存在." & vbcrlf & "(不同网卡略有不同,请据实际情况修改)",8,"MAC地址修改程序",64 _
Else WSHShell.Popup "修改失败",3,"MAC地址修改程序",64

Function restartNetWork(sConnectionName)
Set shellApp = CreateObject("shell.application")
Set oControlPanel = shellApp.Namespace(3)
For Each folderitem in oControlPanel.Items
     If folderitem.Name = "网络连接" Then
         Set oNetConnections = folderitem.GetFolder
         Exit For
     End If
Next  
For Each folderitem in oNetConnections.Items
     If LCase(folderitem.Name) = LCase(sConnectionName) Then
         Set oLanConnection = folderitem
         Exit For
     End If
Next  
For i = 1 To 2
     For Each verb in oLanConnection.verbs
         If verb.Name = "启用(&A)" Then verb.DoIt
         If verb.Name = "停用(&B)" Then verb.DoIt
     Next
     WScript.Sleep 5000
Next
End Function

Function rt(patrn,str)
Set re = New Regexp
re.Pattern = patrn
re.IgnoreCase = True
re.Global = True
rt = re.Test(str)
End Function 





Dim reval
Set objnet = CreateObject ("WScript.Network")
Set R = CreateObject("WScript.Shell")
reval = InputBox("当前的计算机名是：" & objnet.ComputerName,"输入新的计算机名",objnet.ComputerName)
On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")
Set colComputers = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
   
For Each objComputer in colComputers
    errReturn = ObjComputer.Rename (reval)
    If reval <> "" Then
                return=MsgBox ("你确定要重启计算机吗?",vbokcancel+vbexclamation,"注意！")
                If return=vbok Then
                        R.run("Shutdown.exe -r -t 0")
                End if
    End If
Next