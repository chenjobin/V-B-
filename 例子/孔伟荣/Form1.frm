VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9915
   ClientLeft      =   4875
   ClientTop       =   3570
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   12930
   Begin VB.CommandButton Command3 
      Caption         =   "开始抽奖"
      Height          =   975
      Left            =   8640
      TabIndex        =   4
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重输"
      Height          =   495
      Left            =   10440
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   7800
      Top             =   2880
   End
   Begin VB.Label Label1 
      Caption         =   "请输入照片张数"
      Height          =   615
      Left            =   8400
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   5400
      Left            =   360
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   6840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim renshu As Integer


Private Sub Command1_Click()
Static kz As Integer

If kz Mod 2 = 0 Then
Timer1.Enabled = True
Else
Timer1.Enabled = False
End If
kz = kz + 1

End Sub

Private Sub Command2_Click()
Command1.Enabled = True
renshu = Int(Val(Text1.Text))
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Command1.Enabled = False

End Sub

Private Sub Timer1_Timer()
Static m As Integer
Randomize

m = Int(Rnd() * renshu) + 1

Image1.Picture = LoadPicture(App.Path & "\img1\" & Trim(Str(m)) & ".jpg")

End Sub
