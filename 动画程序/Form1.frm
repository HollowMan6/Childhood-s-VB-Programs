VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "动画程序"
   ClientHeight    =   3090
   ClientLeft      =   6525
   ClientTop       =   3765
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "结束"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   -120
      Top             =   3480
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "暂停"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   -120
      OLEDragMode     =   1  'Automatic
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer 'x存入动画分解图形文件的序号
Rem 开始命令按钮中的程序
Private Sub Command1_Click()
  Timer1.Enabled = True '使计时器开始工作
End Sub
Rem 停止命令按钮中的程序
Private Sub Command2_Click()
  Timer1.Enabled = False  '使计时器停止工作
End Sub
Rem 结束命令按钮中的程序
Private Sub Command3_Click()
  End
End Sub
Rem 计时对象中的程序
Private Sub Timer1_Timer()
   x = x + 1               '产生一个分解图形文件的序号
   If x > 30 Then x = 1    '如果序号超过30将序号变为1
   Image1.Picture = LoadPicture(x & ".jpg")  '将序号为x的图形送入Image1对象
End Sub
