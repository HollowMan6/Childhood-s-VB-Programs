VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "倒计时程序"
   ClientHeight    =   1710
   ClientLeft      =   7875
   ClientTop       =   3540
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -240
      Top             =   1920
   End
   Begin VB.CommandButton Command3 
      Caption         =   "结束"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "还剩"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   -120
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "秒"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer  '存放秒值变量
Rem “开始”按钮对象中的程序
Private Sub Command1_Click()
  i = Text1.Text        '文本框对象中输入的秒值放入i变量
  Timer1.Enabled = True '使定时器开始工作进行循环
End Sub
Private Sub Command3_Click()
  End
End Sub

Rem 定时器对象中的程序
Private Sub Timer1_Timer()
  i = i - 1          '秒数递减
  Text1.Text = i     '将递减后的秒数值显示在文本框中
  If i = 0 Then Timer1.Enabled = False  '当秒值递减到0时关闭定时器，停止循环
End Sub
