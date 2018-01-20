VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "查找算法的程序实现"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   9300
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "生成数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "排序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "对分查找"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "顺序查找"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      ItemData        =   "Form1.frx":0000
      Left            =   5880
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      ItemData        =   "Form1.frx":0004
      Left            =   480
      List            =   "Form1.frx":0006
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label5 
      Height          =   855
      Left            =   2640
      TabIndex        =   11
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "查找内容："
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
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "排序后的数据："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "排序前的数据："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const n = 10
Dim d(1 To n) As Integer
Dim sorted As Boolean          '判断是否已经排序
Dim flag As Boolean          '用设置逻辑变量的方式控制循环条件（程序中不使用Exit语句）
Private Sub Command1_Click()   '顺序查找
  key = Val(Text1.Text)               '查找的数据
  For i =_______________                   '依次查找
   
   If _______________ Then                '找到了数据
     
      Label4.Caption = "在数组的第" + Str(i) + "个位置"
      ______________
    End If
    Next i
    If i =_____________ Then         '判断未找到数据
      Label4.Caption = "在数组中没有找到数据" + Str(key)
    End If
End Sub


Private Sub Command2_Click()          '对分查找
 Dim i As Integer, j As Integer, key As Integer, m As Integer
  If Not sorted Then                                 '进行对分查找时，需先对数据进行排序
      MsgBox "进行对分查找时，需先对数据进行排序"
      Exit Sub
  End If
   '对分查找
   key = Val(Text1.Text)             '查找的数据
   i = 1: j = n
   Do While ______________
    
    m = _________________          '中间位置m的计算
   If d(m) = key Then
       Label5.Caption = "在数组的第" + Str(m) + "个位置"
     Exit Do
   End If
   If d(m) < key Then            '判断数据的位置在前半部分还是后半部分
      
      _______________________
    
    Else
    
      _______________________
      
    End If
  Loop
  if _____________ Then
     
     Label5.Caption = "在数组中没有找到数据" + Str(key)
     
  End If

End Sub

Private Sub Command3_Click()

   Dim i As Integer, j As Integer, temp As Integer
   List2.Clear
   sorted = True                     '设置已排序
   For i = 1 To n - 1              '冒泡排序 递增
     For j = n To i + 1 Step -1    '从下向上冒泡
       If d(j) < d(j - 1) Then       '如果下一个数据小，则同上一个数据进行交换
         temp = d(j)
         d(j) = d(j - 1)
         d(j - 1) = temp
        End If
     Next j
   Next i

For i = 1 To n                   '在右边列表中显示排序后的数据
   List2.AddItem Str(i) + " :" + Str(d(i))
Next i
End Sub

Private Sub Command4_Click()
Randomize             '初始化随机函数
sorted = False        '是否已经排序
 List1.Clear
  For i = 1 To 10
    d(i) = 1 + Int(Rnd * 100)
    List1.AddItem Str(i) + " :" + Str(d(i))
  Next
End Sub

