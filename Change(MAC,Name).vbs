On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=true", , 48)
For Each objItem in colItems
     msg = msg & "��ţ�" & objItem.Index & vbCrLf & "MAC��" & objItem.MACAddress & vbCrLf & "������" & objItem.Description & vbCrLf & vbCrLf
     ix = ix & objItem.Index & ","
     MAC = objItem.MACAddress
Next
If msg = "" Then
     MsgBox "δ�ҵ������,����ȷ���˳�����",64,"MAC�޸ĳ���"
     wscript.Quit
End If
Do
     idx = InputBox( msg , "��������Ҫ�޸�MAC���������", Left(ix,InStr(ix,",") - 1))
     if idx = False Then Wscript.Quit
Loop Until IsNumeric(idx) And InStr(ix,idx)
Do
     MAC = InputBox( "������ָ����MAC��ֵַ(ע��Ӧ����12λ���������ֻ���ĸ(A~F)�����û��-�����ȷָ���)" , "�����µ�MAC��ַ", MAC)
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
'�õ�����������
NetWorkName = WshShell.RegRead("HKLM\SYSTEM\ControlSet001\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\" & WshShell.RegRead(reg & "\NetCfgInstanceId") & "\Connection\Name")
WSHShell.Popup "����������������,���Ժ�...",2,"MAC�޸ĳ���",64
restartNetWork NetWorkName
If Not Err Then WSHShell.Popup "�޸ĳɹ�!" & vbcrlf & vbcrlf & "��������쳣,����豸������,չ������������," & vbcrlf & "����Ӧ���������һ�,ѡ������" & _
"�л����߼�ѡ�," & vbcrlf & "������NetworkAddressֵ�޸�Ϊ������." & vbcrlf & "(��ͬ�������в�ͬ,���ʵ������޸�)",8,"MAC��ַ�޸ĳ���",64 _
Else WSHShell.Popup "�޸�ʧ��",3,"MAC��ַ�޸ĳ���",64

Function restartNetWork(sConnectionName)
Set shellApp = CreateObject("shell.application")
Set oControlPanel = shellApp.Namespace(3)
For Each folderitem in oControlPanel.Items
     If folderitem.Name = "��������" Then
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
         If verb.Name = "����(&A)" Then verb.DoIt
         If verb.Name = "ͣ��(&B)" Then verb.DoIt
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
reval = InputBox("��ǰ�ļ�������ǣ�" & objnet.ComputerName,"�����µļ������",objnet.ComputerName)
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
                return=MsgBox ("��ȷ��Ҫ�����������?",vbokcancel+vbexclamation,"ע�⣡")
                If return=vbok Then
                        R.run("Shutdown.exe -r -t 0")
                End if
    End If
Next