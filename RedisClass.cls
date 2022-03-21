VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RedisClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------
'VB6.0��дRedis������ࡣ
'��Ҫ��װ��get/set/lpush/rpush/lpop/rpop�ȳ��õ�Redis�������ܡ�
'���ڳ��õ�Key=>Value����еĶ�д�Ѿ����������ˡ�
'���๦�ܿ��Ե���RedisCommand()����ָ��Ͳ�����ʵ�֣������Լ���װ�ɡ�
'Ϊ����socket����TCP���ӵ�Redis��������������Ҫ����mswinsck.ocx�ؼ���
'ע�⣺���ڲ˵��������̡�>>���� mswinsck.ocx�������ڡ�������>>���ӡ�
'���ߣ���������������ҳ��www.fj543.com ��QQȺ��80784252
'---------------------------------------------------------------------

Option Explicit

Dim WithEvents sckRedis As Winsock '��winsock�ؼ�����һ���࣬�����ɱ���ʹ�á�
Attribute sckRedis.VB_VarHelpID = -1
Private Flag As String '����״̬��־
Private Cache As String '����socket��������
Private ErrMsg As String '��Ŵ�����ʾ��Ϣ
Public Timeout As Long '�����ȴ���ʱ����

Private Sub Class_Initialize()  '�������ʵ��ʱ��ִ��
    On Error Resume Next
    Set sckRedis = New Winsock '����socket����ʵ�����������������ڰ�winsock�ؼ��Ϸŵ�form1�����ϡ�ֻ���������ǡ����á���ʽ���������ڿ��ӻ����塣
    sckRedis.RemotePort = 6379
    sckRedis.RemoteHost = "127.0.0.1"
    Timeout = 5
    'sckRedis.Connect
End Sub
Private Sub Class_Terminate()
    sckRedis.Close
    Set sckRedis = Nothing
End Sub

'����Redis������������/IP
Public Property Let Host(ByVal strHost As String)
    If strHost <> "" Then
        sckRedis.RemoteHost = strHost
        If sckRedis.State <> 0 Then sckRedis.Close
    End If
End Property
Public Property Get Host() As String
    Host = sckRedis.RemoteHost
End Property

'����Redis����˿�
Public Property Let Port(ByVal intPort As Long)
    If intPort > 0 And intPort <= 65535 Then
        sckRedis.RemotePort = intPort
        If sckRedis.State <> 0 Then sckRedis.Close
    End If
End Property
Public Property Get Port() As Long
    Port = sckRedis.RemotePort
End Property

'��ȡ���������Ϣ
Public Property Get lastError() As String
    lastError = ErrMsg
End Property

'��ʱ�ȴ�Redis��socket����״̬���
Private Sub RedisWait(Optional ByVal Seconds As Long)
    Dim t As Single, job As String
    If Seconds < 1 Then Seconds = Timeout
    job = Flag
    t = Timer()
    While (Timer() - t) < Seconds And Flag = job '�ȴ�����ʱ������tag���
        DoEvents
    Wend
End Sub
'����Redis����������״̬
Private Function RedisConn(Optional ByVal Host As String, Optional ByVal Port As Long) As Long
    If Host <> "" Or Port > 0 Then
        If sckRedis.RemoteHost <> Host Or sckRedis.RemotePort <> Port Then 'IP��˿ڲ�ͬ��Ҫ������������
            sckRedis.RemoteHost = Host
            sckRedis.RemotePort = Port
            sckRedis.Close
        End If
    End If
    If sckRedis.State = 0 Then
        sckRedis.Connect
        Flag = "con"
        RedisWait Timeout
    End If
    RedisConn = sckRedis.State
End Function

'ת����ͨ��Redisָ�������ִ��
Public Function RedisCommand(CMD As String, Key As String, Optional Value As String, Optional ByVal ParaCount As Long) As String
    If CMD = "" Then Exit Function
    Dim tmp As String
    Cache = ""
    If RedisConn() = 7 Then
        Flag = CMD
        If ParaCount < 1 Then ' �Զ��жϲ����м�����
            If Value <> "" Then ParaCount = 3 Else ParaCount = 2
        End If
        'sckRedis.SendData "set " & Key & " " & Value & vbCrLf '�������ʽ������Ҫת��ո񡢻��з������ŵȡ�
        'ͨ���﷨����Ҫת��������ţ�Ҫ��ȷͳ���ֽ�����
        tmp = "*" & ParaCount & vbCrLf '���������ܸ���
        tmp = tmp & "$" & Len(CMD) & vbCrLf & CMD & vbCrLf '��һ��������������ĳ��ȼ�����
        tmp = tmp & "$" & LenB(StrConv(Key, vbFromUnicode)) & vbCrLf & Key & vbCrLf '�ڶ��������������������ĳ��ȼ�����
        If ParaCount > 2 Then tmp = tmp & "$" & LenB(StrConv(Value, vbFromUnicode)) & vbCrLf & Value & vbCrLf '������������������ֵ���ĳ��ȼ�����
        'MsgBox tmp
        sckRedis.SendData tmp
        RedisWait Timeout '�ȴ�������Ӧ����
        If Cache = "" Then
            ErrMsg = "��ʱδ�õ���Ӧ"
        Else
            RedisCommand = Cache
        End If
    Else
        ErrMsg = "����δ����"
    End If
End Function


'��Redis�ж�ȡָ��Key�����ݡ�������Get��Set��VB�����﷨��Ϊ�����ͻ��ֻ����GetS��SetS����������
Public Function GetS(ByVal Key As String) As String
    GetS = RedisCommand("GET", Key)
End Function

'��Redisָ��Key�ı���д������
Public Function SetS(ByVal Key As String, ByVal Value As String) As String
    SetS = RedisCommand("SET", Key, Value, 3)
End Function

'ȡ��Redis����Key[]ĩβһ������
Public Function rPop(ByVal Key As String) As String
    rPop = RedisCommand("RPOP", Key)
End Function

'ȡ��Redis����Key[]��ͷһ������
Public Function lPop(ByVal Key As String) As String
    lPop = RedisCommand("LPOP", Key)
End Function

'��Redis����Key[]��ͷ����һ������
Public Function lPush(ByVal Key As String, ByVal Value As String) As String
    lPush = RedisCommand("LPUSH", Key, Value, 3)
End Function

'��Redis����Key[]ĩβ����һ������
Public Function rPush(ByVal Key As String, ByVal Value As String) As String
    rPush = RedisCommand("RPUSH", Key, Value, 3)
End Function


'Redis��Ӧ����
Private Sub sckRedis_DataArrival(ByVal bytesTotal As Long)
    sckRedis.GetData Cache
    Flag = "" '����Ӧ�ˣ���ʾ��ǰ������ɣ��ѱ�־��ա�
End Sub

Private Sub sckRedis_Connect()
    If Flag = "con" Then Flag = "" '���Ӳ�����ɣ���ձ�־
End Sub

Private Sub sckRedis_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ErrMsg = Now() & " ���Ӵ���"
    sckRedis.Close
    If Flag <> "" Then Flag = ""
End Sub
