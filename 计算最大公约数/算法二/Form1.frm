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
m = InputBox("�������һ����Ȼ��", "�������Լ��")
n = InputBox("������ڶ�����Ȼ��", "�������Լ��")
While m / n <> Int(m / n)
c = m - n * Int(m / n)
m = n
n = c
Wend
MsgBox "���Լ��Ϊ" & n, vbInformation, "�����������"
End
End Sub
