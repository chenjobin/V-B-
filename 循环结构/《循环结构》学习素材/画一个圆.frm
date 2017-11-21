VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6255
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Yuan(rr)
  Randomize
  r = Int(Rnd * 255)
  g = Int(Rnd * 255)
  b = Int(Rnd * 255)
  
  Rem 圆心坐标
  x = Form1.ScaleWidth / 2
  y = Form1.ScaleHeight / 2
  rr = rr * 10
  Form1.Circle (x, y), rr, RGB(r, g, b)
End Sub



