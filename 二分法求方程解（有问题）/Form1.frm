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
a = Val(InputBox("����������˵�ֵ", "���ַ����㷽�̽��ƽ�"))
b = Val(InputBox("���������Ҷ˵�ֵ", "���ַ����㷽�̽��ƽ�"))
c = Val(InputBox("�����������", "���ַ����㷽�̽��ƽ�"))
d = Val(InputBox("���뷽�̵����ߣ�ע���ʽ����*����/ָ����^", "���ַ����㷽�̽��ƽ�"))
Do
x0 = (a + b) / 2
x = a
f1 = d
x = x0
f2 = d
If f2 = 0 Then Exit Do
If f1 * f2 < 0 Then
b = x0
Else
a = x0
End If
Loop Until Abs(a - b) < c
MsgBox "���̵Ľ��ƽ�Ϊ" & x0, vbInformation, "�����������"
End
End Sub
