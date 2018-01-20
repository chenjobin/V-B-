VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "冒泡算法"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   7800
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "选择排序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "随机产生20个数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "冒泡排序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      Height          =   3120
      ItemData        =   "排序.frx":0000
      Left            =   3120
      List            =   "排序.frx":0002
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3120
      ItemData        =   "排序.frx":0004
      Left            =   240
      List            =   "排序.frx":0006
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "排序后的数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "原始数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 20) As Integer                  '定义一个数组，存放待排序数据
Dim n As Integer                           '变量n待排序数据个数

Private Sub Command1_Click()               '冒泡排序
    Dim i As Integer, j As Integer, temp As Integer
    List2.Clear                            '将排序后的数据列表框清空
    '冒泡排序
    
    
    
    
    
    
    
    '冒泡排序结束
    For i = 1 To n
        List2.AddItem Str(i) & ":" & Str(a(i))      '将排序结果输出到列表框
    Next i
End Sub


Private Sub Command2_Click()             '产生20个随机数
   n = 20
   Randomize                             '随机数初始化
   List1.Clear                           '将原始数据列表框清空
   List2.Clear                           '将排序后的数据列表框清空
   For i = 1 To n
      a(i) = Int(Rnd * 1000)             'Rnd 函数返回[0,1)随机实数
   List1.AddItem Str(i) & ":" & Str(a(i))            '将待排序数据显示到原始数据列表中
Next

End Sub

Private Sub Command3_Click()
Dim i As Integer, j As Integer, k As Integer, temp As Integer
    List2.Clear                            '将排序后的数据列表框清空
    '选择排序
    
    
    
    
    
    
    
    '选择排序结束
    For i = 1 To n
        List2.AddItem Str(i) & ":" & Str(a(i))      '将排序结果输出到列表框
    Next i
End Sub
