VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   LinkTopic       =   "Form2"
   ScaleHeight     =   4380
   ScaleWidth      =   7845
   StartUpPosition =   3  '窗口缺省
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
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "请输入一个自然数:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "函数的调用"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   2880
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function ou(a) As String
If a Mod 2 = 0 Then
  ou = "该数为偶数！"
Else
  ou = "该数为奇数！"
End If
End Function



Private Sub Text1_Change()
Dim a As Integer
a = Val(Text1.Text)
Label2.Caption = ou(a)
End Sub
