Set argv = WScript.Arguments
If argv.Count = 0 Then
    WScript.Echo "Tips：请把要加密的文件拖到我身上！"
    WScript.Quit
End If
Set fso = CreateObject("Scripting.FileSystemObject")
Randomize
pass = Int(Rnd*12)+20 '异或加密有效范围20-31，所以随机生成好了。
data = fso.OpenTextFile(argv(0), 1).ReadAll
data = "d=" & Chr(34) & ASCdata(data) & Chr(34)
data = data & vbCrLf & ":M=Split(D):For each O in M:N=N&chr(O):Next:execute N"
data = Replace(data, " ", ",")
fso.OpenTextFile(argv(0) & "_加密.vbe", 2, True).Write Encoder(EncHexXorData(data))
WScript.Echo "加密完毕,文件生成到：" & vbCrLf & vbCrLf & argv(0) & "_加密.vbe"

Function EncHexXorData(data)
    EncHexXorData = "x=""" & EncHexXor(data) & """:For i=1 to Len(x) Step 2:s=s&Chr(CLng(""&H""&Mid(x,i,2)) Xor " & pass & "):Next:Execute Replace(s,"","","" "")"
End Function

Function Encoder(data) '加密3
    Encoder = CreateObject("Scripting.Encoder").EncodeScriptFile(".vbs", data, 0, "VBScript")
End Function

Function EncHexXor(x) '加密2
    For i = 1 To Len(x)
        EncHexXor = EncHexXor & Hex(Asc(Mid(x, i, 1)) Xor pass)
    Next
End Function

Function ASCdata(Data) '加密1
    num = Len(data)
    newdata = ""
    For j = 1 To num
        If j = num Then
            newdata = newdata&Asc(Mid(data, j, 1))
        Else
            newdata = newdata&Asc(Mid(data, j, 1)) & " "
        End If
    Next
    ASCdata = newdata
End Function