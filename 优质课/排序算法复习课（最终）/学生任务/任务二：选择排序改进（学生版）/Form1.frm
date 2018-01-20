VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "改进选择排序"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   6720
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "排序"
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "生成数据"
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   2400
      ItemData        =   "Form1.frx":0000
      Left            =   4200
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "Form1.frx":0004
      Left            =   720
      List            =   "Form1.frx":0006
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "排序后数据："
      Height          =   615
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "排序前数据："
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   360
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
  For i = ________ To q
    If a(i) ____ a(iMin) Then iMin = i
    If a(i) ____ a(iMax) Then iMax = i
  Next i
  t = a(iMin): a(iMin) = a(p): a(p) = t
  If iMax = p Then _________
  t = a(iMax): a(iMax) = a(q): a(q) = t
  p = p + 1
  q = __________
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

