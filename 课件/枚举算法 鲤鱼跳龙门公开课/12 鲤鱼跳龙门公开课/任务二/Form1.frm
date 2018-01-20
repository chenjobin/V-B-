VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7260
   ScaleWidth      =   14325
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "破解字母密码"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "破解5位随机数密码"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "破解4位数密码"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "设置字母密码"
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置5位随机数密码"
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "设置4位数密码"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2640
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Long

Private Sub Command1_Click()          'command1 设置4位数密码
a = 1234
Command4.Enabled = True
End Sub

Private Sub Command4_Click()          'command4 破解4位数密码
For i = 1000 To 9999
If i = a Then
Label1.Caption = "密码是" + Str(i)
End If
Next i
Command2.Enabled = True
End Sub

Private Sub Command2_Click()         'command2 设置5位数随机密码并开启command5
Randomize

End Sub

Private Sub Command5_Click()           'command5 破解5位数密码并开启command3

End Sub

