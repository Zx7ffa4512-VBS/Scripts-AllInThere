'----------------------------------------------------------
' CODE BY: EAGLE
' Email    : eagle846@163.com
' QQ    : 369029696
'
' ������ �� IEmsg
' ��    �� �� ����Ļ���½�����һ����ҳ������������ʾ��Ϣ
' ��    �� �� ���Խ���ҳ���Զ���ʱ����Զ��ر�
'
' ��    �� �� IEmsg(title,msg,time)
'    title -     ��ʾ��Ϣ�ı���
'    msg -     ��ʾ��Ϣ�����ݣ����з�Ϊ"<br>"
'    time -     �趨��ҳ�رյ�ʱ�䣬����Ϊ��λ
'
' ��    �� �� Call IEmsg("����-VBS","����-hello word",10)
'----------------------------------------------------------
Call IEmsg("�����Ǳ���","����������",3)
Function IEmsg(title,msg,time)
       On Error Resume Next

       set Oie = createobject("internetexplorer.application")
       screenw = createobject("htmlfile").parentWindow.screen.availWidth
       screenh = createobject("htmlfile").parentWindow.screen.availHeight

       With OIE
            .left    = screenw -300
            .top    = screenh
            .height    = 200
            .width    = 300
            .menubar = 0
            .toolbar = 0
            .statusBar = 0
            .visible = 1
            .navigate    "About:"
       End With

       Do while OIE.busy

       Loop

       With OIE.document
            .Open
            .WriteLn "<HTML><HEAD>"
            .WriteLn "<style type="    & chr(34) &    "text/css"    & chr(34) &    ">"
            .WriteLn " html { background:#e1f4ff;} .titlefont {font-size:19px;color:#ef0eef;}    .msgfont {font-size:14px;color:#000304;}"
            .WriteLn "</style>"
            .WriteLn "<TITLE>" & title & "</TITLE></HEAD>"
            .WriteLn "<BODY>"
            .WriteLn "<span class=" & chr(34) & "titlefont" & chr(34) & ">" & title & "</span><br><span class=" & chr(34) & "msgfont" & chr(34) & ">" & msg & "</font>"
            .WriteLn "</BODY>"
            .WriteLn "</HTML>"
            .Close
       End With

       Do while Oie.top>screenh - Oie.height
            Oie.top = Oie.top - 4
       Loop

       Wscript.sleep CDbl(time * 1000)

       If Oie.Top = "" Then
            Else
            Do while Oie.top < screenh + 50
                 Oie.top = Oie.top + 4
            Loop
            Oie.Quit
       End If
End Function 