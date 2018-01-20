VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "抽奖程序 by jinshanglei"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14490
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   14490
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "重输"
      Height          =   255
      Left            =   13320
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "确认"
      Height          =   255
      Left            =   12360
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
      Height          =   270
      Left            =   13560
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   12960
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始抽奖"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "请输入照片张数:"
      Height          =   375
      Left            =   12240
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   7065
      Left            =   240
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   11700
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


