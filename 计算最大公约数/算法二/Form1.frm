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
m = InputBox("请输入第一个自然数", "计算最大公约数")
n = InputBox("请输入第二个自然数", "计算最大公约数")
While m / n <> Int(m / n)
c = m - n * Int(m / n)
m = n
n = c
Wend
MsgBox "最大公约数为" & n, vbInformation, "结果出来啦！"
End
End Sub
