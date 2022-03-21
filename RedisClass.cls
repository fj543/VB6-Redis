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
'VB6.0读写Redis缓存的类。
'主要封装了get/set/lpush/rpush/lpop/rpop等常用的Redis操作功能。
'对于常用的Key=>Value或队列的读写已经基本够用了。
'更多功能可以调用RedisCommand()传入指令和参数来实现，或者自己封装吧。
'为了用socket创建TCP连接到Redis服务器，工程需要引用mswinsck.ocx控件。
'注意：是在菜单栏“工程”>>引用 mswinsck.ocx，不是在“部件”>>添加。
'作者：福建吴世昌，主页：www.fj543.com ，QQ群：80784252
'---------------------------------------------------------------------

Option Explicit

Dim WithEvents sckRedis As Winsock '将winsock控件当作一个类，声明成变量使用。
Attribute sckRedis.VB_VarHelpID = -1
Private Flag As String '工作状态标志
Private Cache As String '缓存socket返回数据
Private ErrMsg As String '存放错误提示信息
Public Timeout As Long '操作等待超时秒数

Private Sub Class_Initialize()  '创建类的实例时会执行
    On Error Resume Next
    Set sckRedis = New Winsock '创建socket对象实例。这句的作用类似于把winsock控件拖放到form1窗体上。只不过现在是“引用”形式，不依赖于可视化窗体。
    sckRedis.RemotePort = 6379
    sckRedis.RemoteHost = "127.0.0.1"
    Timeout = 5
    'sckRedis.Connect
End Sub
Private Sub Class_Terminate()
    sckRedis.Close
    Set sckRedis = Nothing
End Sub

'设置Redis服务器主机名/IP
Public Property Let Host(ByVal strHost As String)
    If strHost <> "" Then
        sckRedis.RemoteHost = strHost
        If sckRedis.State <> 0 Then sckRedis.Close
    End If
End Property
Public Property Get Host() As String
    Host = sckRedis.RemoteHost
End Property

'设置Redis服务端口
Public Property Let Port(ByVal intPort As Long)
    If intPort > 0 And intPort <= 65535 Then
        sckRedis.RemotePort = intPort
        If sckRedis.State <> 0 Then sckRedis.Close
    End If
End Property
Public Property Get Port() As Long
    Port = sckRedis.RemotePort
End Property

'读取最近错误消息
Public Property Get lastError() As String
    lastError = ErrMsg
End Property

'延时等待Redis的socket任务状态变更
Private Sub RedisWait(Optional ByVal Seconds As Long)
    Dim t As Single, job As String
    If Seconds < 1 Then Seconds = Timeout
    job = Flag
    t = Timer()
    While (Timer() - t) < Seconds And Flag = job '等待到超时，或者tag变更
        DoEvents
    Wend
End Sub
'连接Redis并返回连接状态
Private Function RedisConn(Optional ByVal Host As String, Optional ByVal Port As Long) As Long
    If Host <> "" Or Port > 0 Then
        If sckRedis.RemoteHost <> Host Or sckRedis.RemotePort <> Port Then 'IP或端口不同，要重新设置连接
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

'转换成通用Redis指令参数并执行
Public Function RedisCommand(CMD As String, Key As String, Optional Value As String, Optional ByVal ParaCount As Long) As String
    If CMD = "" Then Exit Function
    Dim tmp As String
    Cache = ""
    If RedisConn() = 7 Then
        Flag = CMD
        If ParaCount < 1 Then ' 自动判断参数有几个。
            If Value <> "" Then ParaCount = 3 Else ParaCount = 2
        End If
        'sckRedis.SendData "set " & Key & " " & Value & vbCrLf '单行命令方式，内容要转义空格、换行符、引号等。
        '通用语法不需要转义特殊符号，要正确统计字节数：
        tmp = "*" & ParaCount & vbCrLf '声明参数总个数
        tmp = tmp & "$" & Len(CMD) & vbCrLf & CMD & vbCrLf '第一个参数（即命令）的长度及内容
        tmp = tmp & "$" & LenB(StrConv(Key, vbFromUnicode)) & vbCrLf & Key & vbCrLf '第二个参数（即变量名）的长度及内容
        If ParaCount > 2 Then tmp = tmp & "$" & LenB(StrConv(Value, vbFromUnicode)) & vbCrLf & Value & vbCrLf '第三个参数（即变量值）的长度及内容
        'MsgBox tmp
        sckRedis.SendData tmp
        RedisWait Timeout '等待返回响应数据
        If Cache = "" Then
            ErrMsg = "超时未得到响应"
        Else
            RedisCommand = Cache
        End If
    Else
        ErrMsg = "连接未就绪"
    End If
End Function


'从Redis中读取指定Key的数据。（由于Get和Set是VB内置语法，为避免冲突，只好用GetS和SetS做函数名）
Public Function GetS(ByVal Key As String) As String
    GetS = RedisCommand("GET", Key)
End Function

'向Redis指定Key的变量写入数据
Public Function SetS(ByVal Key As String, ByVal Value As String) As String
    SetS = RedisCommand("SET", Key, Value, 3)
End Function

'取出Redis队列Key[]末尾一条数据
Public Function rPop(ByVal Key As String) As String
    rPop = RedisCommand("RPOP", Key)
End Function

'取出Redis队列Key[]开头一条数据
Public Function lPop(ByVal Key As String) As String
    lPop = RedisCommand("LPOP", Key)
End Function

'向Redis队列Key[]开头插入一条数据
Public Function lPush(ByVal Key As String, ByVal Value As String) As String
    lPush = RedisCommand("LPUSH", Key, Value, 3)
End Function

'向Redis队列Key[]末尾插入一条数据
Public Function rPush(ByVal Key As String, ByVal Value As String) As String
    rPush = RedisCommand("RPUSH", Key, Value, 3)
End Function


'Redis响应数据
Private Sub sckRedis_DataArrival(ByVal bytesTotal As Long)
    sckRedis.GetData Cache
    Flag = "" '有响应了，表示当前工作完成，把标志清空。
End Sub

Private Sub sckRedis_Connect()
    If Flag = "con" Then Flag = "" '连接操作完成，清空标志
End Sub

Private Sub sckRedis_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ErrMsg = Now() & " 连接错误"
    sckRedis.Close
    If Flag <> "" Then Flag = ""
End Sub

