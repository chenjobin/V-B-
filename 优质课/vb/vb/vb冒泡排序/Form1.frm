VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6960
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "排序"
      Height          =   495
      Left            =   5400
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5280
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   2760
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2760
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "输入数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "排序后的数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "输入的原始数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d(1 To 128) As Integer
Dim ct As Integer
Function adj(a As Integer, n As Integer) As String
Dim sa As String
sa = a: na = Len(a)
For i = 1 To n - na
    sa = " " + sa
Next i
adj = sa
End Function
Private Sub Command1_Click()
For i = 1 To ct - 1                                 '完成升序冒泡排序
    For j = ct To i + 1 Step -1
        If d(i) > d(j) Then
            kt = d(i): d(i) = d(j): d(j) = kt
        End If
    Next j
Next i
For i = 1 To ct
    List2.AddItem adj(Str(i), 3) + adj(Str(d(i)), 6)
Next i
ct = 0
List1.AddItem ""
List2.AddItem ""
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ct = ct + 1
    d(ct) = Val(Text1.Text)
    List1.AddItem adj(Str(ct), 3) + adj(Str(d(ct)), 6)
    Text1.Text = "": Text1.SetFocus
End If
End Sub
