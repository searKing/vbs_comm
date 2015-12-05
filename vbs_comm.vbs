Private Sub MSComm_Send (MSComm1, send_buffer)
	MSComm1.Output = send_buffer
	MSComm1.OutBufferCount=0
End Sub
Private Sub MSComm_Send_crlf (MSComm1)
	MSComm1.Output = ""&vbCr
	MSComm1.OutBufferCount=0
End Sub
Private Function MSComm_Recv (MSComm1)
	WScript.sleep(1000)
	Do
		recv_buffer = recv_buffer & MSComm1.Input
		'MsgBox Cstr(i)&Cstr(recv_buffer)
		i=i+1
	Loop Until MSComm1.InBufferCount=0'InStr(recv_buffer, "OK")' 从串行端口读 "OK" 响应。
	MSComm_Recv = recv_buffer
End Function
Private Function MSComm_Open (commPort, settings)
	comm_name = "MSCOMMLib.MSComm." + Cstr(commPort)
	Set MSComm = createObject(comm_name)
	' 使用 COM1。
	MSComm.CommPort = commPort
	' 9600 波特，无奇偶校验，8 位数据，一个停止位。
	MSComm.Settings = settings
	' 当输入占用时，
	' 告诉控件读入整个缓冲区。
	MSComm.InputLen = 0
	'...打开串口
	If Comm.PortOpen = False Then
		Comm.PortOpen = True 
	End If
	MSComm.InBufferCount=0
	MSComm.OutBufferCount=0
	MSComm_Open = MSComm
End Function
CommPort = 1
Settings = "115200,N,8,1"

MSComm = MSComm_Open(CommPort, Settings) 

'test
send_String = "busybox awk -F '>' 'NR==4{print$2}' /data/data/com.hikvision.iezviz/shared_prefs/login.xml |busybox awk -F'<' '{print$1}'"
MSComm_Send MSComm , send_String
MSComm_Send_crlf MSComm

recv_String = MSComm_Recv(MSComm) 
MsgBox Cstr(recv_String)
