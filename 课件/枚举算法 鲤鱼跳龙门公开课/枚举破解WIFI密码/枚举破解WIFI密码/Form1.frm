VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   7455
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3960
      Top             =   3240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "√‹¬Î£∫"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "SSID"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "√∂æŸ∆∆Ω‚WIFI√‹¬Î"
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
a = 1234
b = 1000
End Sub

Private Sub Timer1_Timer()
If b <> a Then
b = b + 1
Label2.Caption = Str(b)
Else
Timer1.Enabled = False
Label2.Caption = "√‹¬Î «" + Str(b)
End If

End Sub
