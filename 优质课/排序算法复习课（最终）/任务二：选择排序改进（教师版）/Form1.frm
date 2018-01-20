VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "改进选择排序"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5805
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "排序"
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成数据"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   2400
      ItemData        =   "Form1.frx":0000
      Left            =   3840
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "Form1.frx":0004
      Left            =   360
      List            =   "Form1.frx":0006
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "排序后数据："
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "排序前数据："
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const n = 10
Dim a(1 To n) As Integer
Private Sub Command2_Click()
Dim i As Integer
Dim t As Integer
Dim iMin As Integer
Dim iMax As Integer
List2.Clear

p = 1: q = 10
Do While p < q
  iMin = p: iMax = p
  For i = p + 1 To q
    If a(i) < a(iMin) Then iMin = i
    If a(i) > a(iMax) Then iMax = i
  Next i
  t = a(iMin): a(iMin) = a(p): a(p) = t
  If iMax = p Then iMax = iMin
  t = a(iMax): a(iMax) = a(q): a(q) = t
  p = p + 1
  q = q - 1
Loop


For i = 1 To n
   List2.AddItem Str(a(i))
Next
End Sub

Private Sub Command1_Click()
  List1.Clear
  For i = 1 To 10
    a(i) = 1 + Int(Rnd * 100)
    List1.AddItem Str(a(i))
  Next
End Sub

