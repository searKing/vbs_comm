Private Sub Form_Send (MSComm1, send_buffer)
	MSComm1.Output = send_buffer
	MSComm1.OutBufferCount=0
End Sub
Private Sub Form_Send_crlf (MSComm1)
	MSComm1.Output = ""&vbCr
	MSComm1.OutBufferCount=0
End Sub
Private Sub Form_Recv (MSComm1)
	WScript.sleep(1000)
	Do
		recv_buffer = recv_buffer & MSComm1.Input
		'MsgBox Cstr(i)&Cstr(recv_buffer)
		i=i+1
	Loop Until MSComm1.InBufferCount=0'InStr(recv_buffer, "OK")
	
	' 从串行端口读 "OK" 响应。
	MsgBox Cstr(recv_buffer)
End Sub

Set MSComm1 = createObject("MSCOMMLib.MSComm.1")
' 使用 COM1。
MSComm1.CommPort = 1
' 9600 波特，无奇偶校验，8 位数据，一个停止位。
MSComm1.Settings = "115200,N,8,1"
' 当输入占用时，
' 告诉控件读入整个缓冲区。
MSComm1.InputLen = 0
' 打开端口。
MSComm1.PortOpen = True
MSComm1.InBufferCount=0
MSComm1.OutBufferCount=0

'test1
Form_Send MSComm1 , "ls"
Form_Send_crlf MSComm1
Form_Recv MSComm1

'test2
send_String = "busybox awk -F '>' 'NR==4{print$2}' /data/data/com.hikvision.iezviz/shared_prefs/login.xml |busybox awk -F'<' '{print$1}'"
Form_Send MSComm1 , send_String
Form_Send_crlf MSComm1
Form_Recv MSComm1

'test3
Form_Send MSComm1 , "busybox "
Form_Send MSComm1 , "awk -F"
Form_Send MSComm1 , "'>' "
Form_Send MSComm1 , "'NR==4{print$2}' "
Form_Send MSComm1 , "/data/data/"
Form_Send MSComm1 , "com.hikvision.iezviz/"
Form_Send MSComm1 , "shared_prefs/"
Form_Send MSComm1 , "login.xml"
Form_Send MSComm1 , "|busybox "
Form_Send MSComm1 , "awk -F'<' "
Form_Send MSComm1 , "'{print$1}' "
Form_Send_crlf MSComm1


Form_Recv MSComm1
