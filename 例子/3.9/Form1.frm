VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "匀加速运动1"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8160
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
   ScaleHeight     =   5295
   ScaleWidth      =   8160
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   555
      Left            =   3360
      TabIndex        =   3
      Top             =   990
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   555
      Left            =   3360
      TabIndex        =   2
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "q确定"
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "初始速度"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "加速度"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "时间"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "dsad"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   3480
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
v0 = Val(Text1.Text)
a = Val(Text2.Text)
t = Val(Text3.Text)
Label4.Caption = Str(v0 + a * t)
End Sub

Private Sub Label2_Click()

End Sub

