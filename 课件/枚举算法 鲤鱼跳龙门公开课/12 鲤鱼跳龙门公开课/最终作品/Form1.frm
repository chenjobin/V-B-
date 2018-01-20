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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   8520
      Top             =   5760
   End
   Begin VB.CommandButton Command8 
      Caption         =   "跳过龙门"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6600
      TabIndex        =   10
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "打开龙门"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6600
      TabIndex        =   9
      Top             =   4800
      Width           =   1695
   End
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
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   2640
      Picture         =   "Form1.frx":6BF37
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
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
Dim b As String
Private Sub Command1_Click()
a = 1234
Command4.Enabled = True
End Sub

Private Sub Command2_Click()
a = Int(Rnd * 90000 + 10000)
Image1.Visible = False
Command5.Enabled = True
End Sub

Private Sub Command3_Click()
b = "fish"
Command6.Enabled = True
End Sub

Private Sub Command4_Click()          '破解4位数密码
For i = 0 To 9999
If i = a Then
Label1.Caption = "密码是" + Str(i)
End If
Next i
Command2.Enabled = True
End Sub

Private Sub Command5_Click()           '破解5位数密码
For i = 10000 To 99999
If i = a Then
Label2.Caption = "密码是" + Str(i)
End If
Next i
Command3.Enabled = True
End Sub

Private Sub Command6_Click()               '破解字母密码fish
Dim a(1 To 4) As Integer
Dim b(1 To 4) As String
a1 = Asc("f")
a2 = Asc("i")
a3 = Asc("s")
a4 = Asc("h")
For i = 97 To 122 Step 1
If i = a1 Then b1 = Chr(i)
If i = a2 Then b2 = Chr(i)
If i = a3 Then b3 = Chr(i)
If i = a4 Then b4 = Chr(i)
Next i
Label3.Caption = "密码是" + b1 + b2 + b3 + b4
Command7.Enabled = True
Command8.Enabled = True

End Sub

Private Sub Command7_Click()                      '打开龙门
Form1.Picture = LoadPicture("lm.jpg")
Image1.Visible = True
End Sub

Private Sub Command8_Click()                     '启动定时器
Timer1.Enabled = True
Timer1.Interval = 100

End Sub

Private Sub Timer1_Timer()                           '游向龙门
If Image1.Left < Form1.Width Then
Image1.Left = Image1.Left + 100
Else
Timer1.Enabled = False
Form1.Picture = LoadPicture("cg.jpg")
End If
End Sub
