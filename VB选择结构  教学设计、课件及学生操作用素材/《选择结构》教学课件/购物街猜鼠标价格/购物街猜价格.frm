VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   12615
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "结束"
      Height          =   615
      Left            =   10680
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   615
      Left            =   8880
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   0
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "元"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10800
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "鼠标价格大竞猜"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1455
      Left            =   8520
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "请输入你猜的价格"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   660
      Left            =   8520
      TabIndex        =   1
      Top             =   2880
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   9800
      Left            =   0
      Picture         =   "购物街猜价格.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer
Rem
Private Sub Form_Load()
Randomize
a = 47
End Sub
Private Sub Command1_Click()

x = Text1.Text
b = Val(x)
If a = b Then
Label2.Caption = "恭喜你猜中了！"
Else
cc
End If
End Sub
Rem
Public Sub cc()
If b < a Then
Label2.Caption = "你猜的价格低了，请再猜！"
Else
Label2.Caption = "你猜的价格高了，请再猜！"
End If
End Sub
Rem
Private Sub Command2_Click()
End
End Sub

