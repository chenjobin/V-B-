VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "QQ用户登录"
   ClientHeight    =   3195
   ClientLeft      =   4395
   ClientTop       =   3405
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   Picture         =   "QQ.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4860
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退　出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确  定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Let user = Text1.Text '将输入文本框1中的账号赋给变量a
  Let password = Text2.Text '将输入文本框2中的密码赋给变量b
  Rem 对输入的用户名和口令进行判断
  If (user = "666" And password = "666") Then
    dlcg '调用登录成功模块
  Else
    dlsb '调用登录失败模块
  End If
End Sub
Private Sub Command2_Click()
  End
End Sub

Private Sub Form_Load()
  Text2.PasswordChar = "*" '使输入的字符都为*号
  Text2.MaxLength = 6  '使Text1对象中最多只能输入6个字符
End Sub
