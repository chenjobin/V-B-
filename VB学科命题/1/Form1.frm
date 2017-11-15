VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "政治练习题"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6390
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "判断"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
      Begin VB.OptionButton Option1 
         Caption         =   "货币是财富的象征"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3120
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "货币能够表现商品的价值"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "货币的本质是一般等价物"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "货币是商品，具有价值"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "1.货币能充当价值尺度职能的根本原因是"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "政治单选题"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
answer = "货币是商品，具有价值"
If Label3.Caption = answer Then
    MsgBox "恭喜你答对了"
Else
    MsgBox "哦no,答错了"
End If
Command1.Enabled = False
Command2.Visible = True
End Sub


Private Sub Command2_Click()
End
End Sub

Private Sub Option1_Click(Index As Integer)
Label3.Caption = Option1(Index).Caption

End Sub
