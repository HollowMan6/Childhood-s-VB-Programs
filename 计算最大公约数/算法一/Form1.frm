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
a = InputBox("�������һ����Ȼ��", "�������Լ��")
b = InputBox("������ڶ�����Ȼ��", "�������Լ��")
While a Mod b <> 0
r = a Mod b
a = b
b = r
Wend
MsgBox "���Լ��Ϊ" & b, vbInformation, "�����������"
End
End Sub
