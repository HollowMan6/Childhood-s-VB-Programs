VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "�ػ���ʽ��ȫ"
   ClientHeight    =   4095
   ClientLeft      =   4725
   ClientTop       =   2880
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command8 
      Caption         =   "�����ػ�(����Win8.1 64λ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Զ�̹ػ� 
      Caption         =   "Զ�̹ػ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "ת���߼������˵�����������(����Win8�����ϲ���ϵͳ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7080
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "ע��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000010&
      Caption         =   "ȡ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   72
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   8895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ػ�(��������,����Win8�����ϲ���ϵͳ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ػ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "             ��ѡ������һ�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "shutdown.exe -s -t 00"
End Sub

Private Sub Command2_Click()
Shell "shutdown.exe -s -hybrid -t 00"
End Sub

Private Sub Command3_Click()
Shell "shutdown.exe - h"
End
End Sub

Private Sub Command4_Click()
Shell "shutdown.exe -r -t 00"
End Sub

Private Sub Command5_Click()
Shell "shutdown.exe -a"
End
End Sub

Private Sub Command6_Click()
Shell "shutdown.exe -l"
End Sub

Private Sub Command7_Click()
Shell "shutdown.exe -r -o -t 00"
End Sub

Private Sub Command8_Click()
Shell "�����Թرյ���.exe"
End
End Sub

Private Sub Զ�̹ػ�_Click()
Shell "shutdown.exe -i", vbNormalFocus
End
End Sub
