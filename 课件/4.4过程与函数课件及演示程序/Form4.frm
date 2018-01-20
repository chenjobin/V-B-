VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form4"
   ScaleHeight     =   6945
   ScaleWidth      =   7410
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Height          =   1095
      Left            =   3840
      TabIndex        =   3
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   3840
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   3840
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求最大公约数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2520
      TabIndex        =   0
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "a,b的最大公约数是:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "b="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "a="
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
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, m, n As Integer
Private Sub Command1_Click()

a = Val(Text1.Text)
b = Val(Text2.Text)

Text3.Text = Str(f(a, b))

End Sub
Public Function f(m, n) As Integer
Do While m <> n
   Do While m > n
     m = m - n
   Loop
   Do While n > m
      n = n - m
   Loop
Loop
   f = m
End Function
