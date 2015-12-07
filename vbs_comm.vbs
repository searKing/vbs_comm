Private Sub MSComm_Send (ms_comm, send_buffer)
	ms_comm.Output = send_buffer
	ms_comm.OutBufferCount=0
End Sub
Private Sub MSComm_Send_crlf (ms_comm)
	ms_comm.Output = ""&vbCr
	ms_comm.OutBufferCount=0
End Sub
Private Function MSComm_Recv (ms_comm)
	WScript.sleep(1000)
	Do
		recv_buffer = recv_buffer & ms_comm.Input
		'MsgBox Cstr(i)&Cstr(recv_buffer)
		i=i+1
	Loop Until ms_comm.InBufferCount=0'InStr(recv_buffer, "OK")' 从串行端口读 "OK" 响应。
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
'获取最后一行字符串
recv_StringArray = split(recv_String,"\n")
last_idx=ubound(recv_StringArray)-1
last_recv_String = recv_StringArray(last_idx)
MsgBox Cstr(last_recv_String)
