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
a = Val(InputBox("����a��������ϵ��ֵ��", "��һԪ���η��̸�"))
b = Val(InputBox("����b��һ����ϵ��ֵ��", "��һԪ���η��̸�"))
c = Val(InputBox("����c��������ϵ��ֵ��", "��һԪ���η��̸�"))
If b ^ 2 - 4 * a * c < 0 Then
MsgBox "������ʵ����", vbInformation, "�����������"
Else
MsgBox "��=" & b ^ 2 - 4 * a * c, vbInformation, "�����������"
MsgBox "��������=" & Sqr(b ^ 2 - 4 * a * c), vbInformation, "�����������"
MsgBox "���̵�һ��ʵ����X=" & -(b + Sqr(b ^ 2 - 4 * a * c)) / 2 * a, vbInformation, "�����������"
MsgBox "���̵���һ��ʵ����X=" & -(b - Sqr(b ^ 2 - 4 * a * c)) / 2 * a, vbInformation, "�����������"
End If
End
End Sub
