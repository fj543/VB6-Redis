VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Redis测试  （By fj543）"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6915
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton btnPop 
      Caption         =   "读Redis队列"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton btnPush 
      Caption         =   "写入Redis队列"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtData2 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   2640
      Width           =   6735
   End
   Begin VB.CommandButton btnRead 
      Caption         =   "读Redis数据"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox txtData 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":001A
      Top             =   120
      Width           =   6735
   End
   Begin VB.CommandButton btnWrite 
      Caption         =   "写数据到Redis"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim R As RedisClass '声明自定义类的变量

Private Sub btnPop_Click()
    txtData2.Text = R.rPop("urls") '从redis缓存队列"urls"中取出末尾一条数据
End Sub

Private Sub btnPush_Click()
    txtData2.Text = R.lPush("urls", txtData.Text) '向redis缓存队列"urls"的开头插入一条数据
End Sub

Private Sub btnRead_Click()
    txtData2.Text = R.GetS("test") '从redis中读取变量test的值
    If txtData2.Text = "" Then txtData2.Text = R.lastError
End Sub

Private Sub btnWrite_Click()
    txtData2.Text = R.SetS("test", txtData.Text) '将文本框内容写入到Redis中的变量test
    If txtData2.Text = "" Then txtData2.Text = R.lastError
End Sub

Private Sub Form_Load()
    Set R = New RedisClass '将变量实例化为可用的对象
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set R = Nothing
End Sub
