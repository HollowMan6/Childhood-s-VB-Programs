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
s = 0
n = InputBox("����Ͷ������", "ģ��", 10000000)
For i = 1 To n
If Rnd > 0.5 Then s = s + 1
Next i
MsgBox "��������ĸ���Ϊ" & s / n, vbInformation, "�����������"
End
End Sub
