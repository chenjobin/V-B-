VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7260
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1380
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Randomize
a = Int(Rnd() * 19) + 1
Text1.Text = Str(a)
End Sub

