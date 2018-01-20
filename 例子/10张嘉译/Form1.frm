VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   5055
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   2760
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "华氏转摄氏"
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "摄氏温度"
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "华氏温度"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
f = Val(Text1.Text)
c = (f - 32) * 5 / 9
Text2.Text = Str(c)
End Sub

Private Sub Command2_Click()

End Sub
