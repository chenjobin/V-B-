VERSION 5.00
Begin VB.Form 冒泡和排序算法 
   Caption         =   "冒泡和排序算法"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   7800
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List3 
      Height          =   3120
      Left            =   3720
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "选择排序"
      Height          =   615
      Left            =   5640
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "冒泡排序"
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "随机产生8个数"
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      Height          =   3120
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3120
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "选择排序结果"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "冒泡排序结果"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "原始数据"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "冒泡和排序算法"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim a(1 To 8) As Integer  '定义一个数组


Private Sub Command1_Click()  '产生8个随机数
   Randomize  '随机数初始化
List1.Clear '原始数据清空
List2.Clear   '将排序后的列表数据清空
List3.Clear   '将排序后的列表数据清空
For i = 1 To 8
   a(i) = Int(Rnd * 100) 'Rnd 函数返回的随机数介于 0 和 1 之间，可等于 0，但不等于 1
   List1.AddItem Str(a(i))  '将数据显示到原始数据列表中
Next i

End Sub

Private Sub Command2_Click()  '对8个数进行冒泡法排序


End Sub

Private Sub Command3_Click() '对8个数进行选择法排序

   
End Sub

