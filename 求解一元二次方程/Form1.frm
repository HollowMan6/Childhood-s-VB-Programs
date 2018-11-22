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
a = Val(InputBox("输入a（二次项系数值）", "求一元二次方程根"))
b = Val(InputBox("输入b（一次项系数值）", "求一元二次方程根"))
c = Val(InputBox("输入c（常数项系数值）", "求一元二次方程根"))
If b ^ 2 - 4 * a * c < 0 Then
MsgBox "方程无实数根", vbInformation, "结果出来啦！"
Else
MsgBox "Δ=" & b ^ 2 - 4 * a * c, vbInformation, "结果出来啦！"
MsgBox "Δ开方后=" & Sqr(b ^ 2 - 4 * a * c), vbInformation, "结果出来啦！"
MsgBox "方程的一个实数根X=" & -(b + Sqr(b ^ 2 - 4 * a * c)) / 2 * a, vbInformation, "结果出来啦！"
MsgBox "方程的另一个实数根X=" & -(b - Sqr(b ^ 2 - 4 * a * c)) / 2 * a, vbInformation, "结果出来啦！"
End If
End
End Sub
