VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "圆周率的计算"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   9315
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Text            =   "100"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   9
      Left            =   7440
      TabIndex        =   40
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   8
      Left            =   7440
      TabIndex        =   39
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   7
      Left            =   7440
      TabIndex        =   38
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   6
      Left            =   7440
      TabIndex        =   37
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   5
      Left            =   7440
      TabIndex        =   36
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   4
      Left            =   7440
      TabIndex        =   35
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   3
      Left            =   7440
      TabIndex        =   34
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   2
      Left            =   7440
      TabIndex        =   33
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   7440
      TabIndex        =   32
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   7440
      TabIndex        =   31
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   9
      Left            =   5760
      TabIndex        =   29
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   8
      Left            =   5760
      TabIndex        =   28
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   7
      Left            =   5760
      TabIndex        =   27
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   6
      Left            =   5760
      TabIndex        =   26
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   5
      Left            =   5760
      TabIndex        =   25
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   4
      Left            =   5760
      TabIndex        =   24
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   3
      Left            =   5760
      TabIndex        =   23
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   2
      Left            =   5760
      TabIndex        =   22
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   5760
      TabIndex        =   21
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   1
      Left            =   5760
      TabIndex        =   20
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Text            =   "1000"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "（2）蒙特卡洛计算"
      Height          =   435
      Left            =   7320
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "（1） 迭代法计算"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000005&
      Caption         =   "Pai="
      Height          =   375
      Left            =   4680
      TabIndex        =   43
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000005&
      Caption         =   "Pai="
      Height          =   375
      Left            =   4680
      TabIndex        =   42
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000005&
      Caption         =   "迭代次数："
      Height          =   375
      Left            =   5040
      TabIndex        =   41
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "x 轴点分布"
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   8
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "y 轴点分布"
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   30
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "0.9-1.0"
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   19
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "0.8-0.9"
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   18
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "0.7-0.8"
      Height          =   255
      Index           =   6
      Left            =   4920
      TabIndex        =   17
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "0.6-0.7"
      Height          =   255
      Index           =   5
      Left            =   4920
      TabIndex        =   16
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "0.5-0.6"
      Height          =   255
      Index           =   4
      Left            =   4920
      TabIndex        =   15
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "0.4-0.5"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   14
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "0.3-0.4"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   13
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "0.2-0.3"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   12
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      Caption         =   "0.1-0.2"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   11
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Caption         =   "0-0.1"
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "每次投点数："
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   4125
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   2520
      Width           =   4470
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "2-使用蒙特卡洛法"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "1-使用迭代公式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   360
      Picture         =   "Form1.frx":3C2C2
      Top             =   840
      Width           =   4125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Pai As Double
    Dim B As Double
    Dim I, N As Integer
    N = Val(Text6.Text)
    Pai = 1
    B = 1
    For I = 2 To N
        B = B * (I - 1) / (2 * I - 1)
        Pai = Pai + B
    Next
    Text1.Text = Pai * 2

End Sub

Private Sub Command2_Click()
Randomize Timer
Dim I As Long, N As Long, S As Long, J As Long
Dim ar(1 To 10), br(1 To 10) As Long
Dim Pai As Double, A As Double, B As Double, C As Double
I = 1: S = 0
N = Val(Text3.Text)
While I <= N
    A = Rnd: B = Rnd
    C = Sqr(A * A + B * B)
    J = Int(A * 10) + 1
    ar(J) = ar(J) + 1
    J = Int(B * 10) + 1
    br(J) = br(J) + 1
    
    If C <= 1 Then
        S = S + 1
    End If
    I = I + 1
Wend
Pai = S * 4 / N
Text2.Text = Pai
For J = 1 To 10
    Text4(J - 1).Text = Str(ar(J)) + " , (" + Str(ar(J) - N / 10) + ")"
    Text5(J - 1).Text = Str(br(J)) + " , (" + Str(br(J) - N / 10) + ")"
Next

End Sub

