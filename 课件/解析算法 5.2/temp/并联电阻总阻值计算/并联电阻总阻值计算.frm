VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "计算并联电阻的总阻值"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0000FFFF&
      Height          =   2220
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2520
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "并联组内各电阻阻值"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "输入电阻值"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "并联的总电阻值"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As Double
Sub Text2_Click()
 List1.Clear
 rs = 0
 Text1.Text = ""
 Text2.Text = ""
End Sub
Sub Text2_KeyPress(KeyAscii As Integer)
 Dim r As Double
 If KeyAscii = 13 Then
    r = Val(Text2.Text)   '输入的是一个有效的电阻值
    If r > 0 Then
      rs = rs + 1 / r     '有效电阻值的倒数累加到rs
      List1.AddItem Str(r)
    End If
    Text2.Text = ""
 End If
End Sub
Sub Command1_Click()
 If rs > 0 Then Text1.Text = Str(1 / rs) Else Text1.Text = "无连接"
End Sub
