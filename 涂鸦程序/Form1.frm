VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Ϳѻ����"
   ClientHeight    =   3510
   ClientLeft      =   -60
   ClientTop       =   -15
   ClientWidth     =   20490
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "             ˵��������ס������ʱ�ƶ����ߣ������Ҫ�������������һֱ��������Ҽ����ƶ���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Rem ����������������ڴ����ϵ�λ�ö�Ϊ�������
  Form1.CurrentX = X
  Form1.CurrentY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Rem ����ס����ƶ���껭��
  If Button = 1 Then Form1.Line -(X, Y)
  Rem ����ס�Ҽ��ƶ�����崰��������
  If Button = 2 Then Form1.Cls
End Sub


