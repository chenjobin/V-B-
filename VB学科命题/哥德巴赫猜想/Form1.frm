VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "哥德巴赫猜想"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   9135
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "子函数计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "函数计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   3000
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "表达式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "请输入一个大于2的偶数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "任何大于2的偶数都是素数之和？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()    '调用函数过程的主过程
    Dim i, S, N As Integer
    N = Text1.Text                        '将文本框中的数字赋给变量N
    For i = 3 To N Step 2                     '产生3到N中间的奇数
        S = i                            '将变量I的值赋给变量S
        If HQSS(S) = 1 Then              '如果函数值为1，则说明I是素数
            S = N - i                    '将N-1的值赋给变量S
            If HQSS(S) = 1 Then Exit For    '如果函数值为1，则脱离循环
        End If
    Next i
    S = HSJG(i, N)                         '调用显示函数程序
    If N = 4 Then Label4.Caption = "4=2+2"   '如果N=4，重新显示结果
End Sub

Public Function HQSS(ByVal S1 As Integer)   '判断是否为素数的函数过程
    Dim J As Integer
    K1 = 1                              '变量K1为1时，是素数，否则不是
    For J = 2 To Sqr(S1)                '进行素数判断
        If S1 / J = Int(S1 / J) Then K1 = 0: Exit For   '如果可以整除，则不为素数，退出循环
    Next J
    HQSS = K1       '将判断结果赋函数
End Function

Public Function HSJG(ByVal i As Integer, ByVal N As Integer)
    Label4.Caption = Str(N) + "=" + Str(i) + "+" + Str(N - i)
End Function

Private Sub command2_click()    '调用子过程的主过程
Dim i, S, N, K As Integer
N = Text1.Text
For i = 3 To N Step 2
    S = i: Call ZQSS(S, K)       '调判断素数子程序
    If K = 1 Then
        S = N - i: Call ZQSS(S, K)      '调判断素数子程序
        If K = 1 Then Exit For
    End If
Next i
Call HSJG(i, N)     '调显示子程序
If N = 4 Then Label4.Caption = "4=2+2"
End Sub
Public Sub ZQSS(ByVal S1 As Integer, K1 As Integer)  '判断是否为素数的子过程
Dim J As Integer
K1 = 1
For J = 2 To Sqr(S1)
    If S1 / J = Int(S1 / J) Then K1 = 0: Exit For
Next J
End Sub
Public Sub XSJG(ByVal i As Integer, ByVal N As Integer)
    Label4.Caption = Str(N) + "=" + Str(i) + "+" + Str(N - i)
End Sub
Private Sub command3_click()
    End
End Sub
