VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   4335
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1680
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "结果"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "最终值："
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "1-1/2+1/3-1/4+...-1/最终值的值为："
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
j = Text2.Text
s = 0
For i = 1 To j Step 1
If i Mod 2 = 1 Then
  s = s + 1 / i
ElseIf i Mod 2 = 0 Then
  s = s - 1 / i
End If
Next i
Text1.Text = Str(s)
  




End Sub

