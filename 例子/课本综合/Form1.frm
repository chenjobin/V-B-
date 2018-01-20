VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "登录程序"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   6885
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5280
      TabIndex        =   10
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   720
      Top             =   3840
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   3840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   4560
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   3240
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   2760
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "验证码"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "密码"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Integer, t As Integer

Private Sub Command1_Click()
Dim uname As String, psd As String, vcode As Integer
uname = Text1.Text
psd = Text2.Text
vcode = Val(Text3.Text)
If vcode <> Val(Label4.Caption) Then
    num = num - 1
    Label5.Caption = "验证码有误，还有" & Str(num) & "次机会"
ElseIf uname <> "admin" Or psd <> "123" Then
    num = num - 1
    Label5.Caption = "用户名或密码有误，还有" & Str(num) & "次机会"
Else
    Label5.Caption = "验证成功！"
    Timer1.Interval = 200
    Image1.Visible = True
End If
If num < 1 Then
    MsgBox "3次机会已到，程序结束"
    End
End If
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Form_Load()
Randomize
Label4.Caption = Int(Rnd * 9000) + 1000
Image1.Visible = False
num = 3
t = 60
Timer2.Interval = 1000
End Sub



Private Sub Timer1_Timer()
Dim n As Integer, ss As String
ss = Label5.Caption
n = Len(ss)
Label5.Caption = Mid(ss, 2, n - 1) + Mid(ss, 1, 1)
End Sub

Private Sub Timer2_Timer()
Label6.Caption = "还剩余" & Str(t) & "秒"
t = t - 1
If t < 0 Then
    MsgBox "时间到，程序结束"
    End
End If
End Sub
