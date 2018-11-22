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
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
a = Val(InputBox("输入区间左端点值", "二分法计算方程近似解"))
b = Val(InputBox("输入区间右端点值", "二分法计算方程近似解"))
c = Val(InputBox("输入误差限制", "二分法计算方程近似解"))
d = Val(InputBox("输入方程的左半边，注意格式乘用*除用/指数用^", "二分法计算方程近似解"))
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
MsgBox "方程的近似解为" & x0, vbInformation, "结果出来啦！"
End
End Sub
