VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   10665
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   8520
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   3
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "过程的调用 "
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   1
      Top             =   6960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   4680
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim a As String
Label2.Caption = ("请输入你的姓名:")
a = Text2.Text
If a <> "" Then
List1.AddItem a + "我知道你了，哈哈!"
End If
End Sub

Private Sub Label1_Dblclick()
List1.AddItem "你双击我了，哈哈!"
Command1_Click
End Sub

Private Sub Text1_Change()
List1.AddItem "你输入东西了，哈哈"
Command1_Click
End Sub
