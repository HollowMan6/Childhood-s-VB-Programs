VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
T = 1
S = 0
i = 0
n = Val(InputBox("����������������һЩ�Żᾫ׼�������鲻Ҫ̫��", "����", 10000000))
While T < n
i = 1 / T - 1 / (T + 2) + S
T = T + 4
S = i
Wend
MsgBox "Բ����" & 4 * S, vbInformation, "�����������"
End
End Sub
