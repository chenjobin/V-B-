VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "计算n的阶乘"
   ClientHeight    =   2835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   21.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   7500
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   675
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "n!="
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "输入N="
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer, n As Integer
Dim f As Double
n = Val(Text1.Text)
f = 1
i = 1
Do While i <= n
    f = f * i
    i = i + 1
Loop
Text2.Text = Str(f)
End Sub
