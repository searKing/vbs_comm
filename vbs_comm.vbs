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
	Loop Until ms_comm.InBufferCount=0'InStr(recv_buffer, "OK")' �Ӵ��ж˿ڶ� "OK" ��Ӧ��
	MSComm_Recv = recv_buffer
End Function
Private Function MSComm_Open (commPort, settings)
	comm_name = "MSCOMMLib.MSComm." + Cstr(commPort)
	Set ms_comm = createObject(comm_name)
	' ʹ�� COM1��
	ms_comm.CommPort = commPort
	' 9600 ���أ�����żУ�飬8 λ���ݣ�һ��ֹͣλ��
	ms_comm.Settings = settings
	' ������ռ��ʱ��
	' ���߿ؼ�����������������
	ms_comm.InputLen = 0
	'...�򿪴���
	If ms_comm.PortOpen = False Then
		ms_comm.PortOpen = True
	End If
	ms_comm.InBufferCount=0
	ms_comm.OutBufferCount=0
	set MSComm_Open = ms_comm
End Function
CommPort = 1
Settings = "115200,N,8,1"
set MSComm = MSComm_Open(CommPort, Settings)

'test
send_String = "busybox awk -F '>' 'NR==4{print$2}' /data/data/com.hikvision.iezviz/shared_prefs/login.xml |busybox awk -F'<' '{print$1}'"
MSComm_Send MSComm , send_String
MSComm_Send_crlf MSComm

recv_String = MSComm_Recv(MSComm)
'��ȡ���һ���ַ���
recv_StringArray = split(recv_String,vbLf)
last_idx=ubound(recv_StringArray)-1
last_recv_String = recv_StringArray(last_idx)
MsgBox Cstr(last_recv_String)
