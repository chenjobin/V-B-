VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   BeginProperty Font 
      Name            =   "ËÎÌå"
      Size            =   36
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   9360
   ScaleWidth      =   13980
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   360
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   3915
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   4365
      Left            =   5400
      Picture         =   "Form2.frx":332C8
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   6015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Randomize
Label1.Caption = Int(Rnd() * 9000 + 1000)
End Sub

