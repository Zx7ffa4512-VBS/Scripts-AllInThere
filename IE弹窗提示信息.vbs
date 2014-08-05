'----------------------------------------------------------
' 函数名 ： IEmsg
' 功    能 ： 在屏幕右下角升起一个网页，可以用于提示信息
' 特    点 ： 可以将网页在自定义时间后自动关闭
'
' 参    数 ： IEmsg(title,msg,time)
'    title -     提示信息的标题
'    msg -     提示信息的内容，换行符为"<br>"
'    time -     设定网页关闭的时间，以秒为单位
'
' 例    子 ： Call IEmsg("标题-VBS","内容-hello word",10)
'----------------------------------------------------------
Call IEmsg("这里是标题","这里是内容",3)
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
