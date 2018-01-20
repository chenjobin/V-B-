VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "冒泡排序变式（教师版）"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   11445
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   $"J.frx":0000
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   1080
      TabIndex        =   6
      Top             =   3720
      Width           =   8295
      Begin VB.CommandButton Command2 
         Caption         =   "排序"
         Height          =   495
         Left            =   5160
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Height          =   2220
         ItemData        =   "J.frx":001A
         Left            =   6480
         List            =   "J.frx":001C
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox List3 
         Height          =   2220
         ItemData        =   "J.frx":001E
         Left            =   3600
         List            =   "J.frx":0020
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "排序后数据："
         Height          =   375
         Left            =   6240
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "排序前数据："
         Height          =   375
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   2775
         Left            =   360
         Picture         =   "J.frx":0022
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   $"J.frx":8E8F
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton Command1 
         Caption         =   "沉泡排序"
         Height          =   495
         Left            =   5040
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Height          =   2220
         ItemData        =   "J.frx":8EB7
         Left            =   6360
         List            =   "J.frx":8EB9
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   2220
         ItemData        =   "J.frx":8EBB
         Left            =   3600
         List            =   "J.frx":8EBD
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "排序后数据："
         Height          =   495
         Left            =   6240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "排序前数据："
         Height          =   495
         Left            =   3600
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   2775
         Left            =   360
         Picture         =   "J.frx":8EBF
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 5) As Integer
Private Sub Command1_Click()
Dim i As Integer
Dim j As Integer
Dim t As Integer

For i = 5 To 2 Step -1
   For j = 1 To i - 1
       If a(j) < a(j + 1) Then
        t = a(j + 1): a(j + 1) = a(j): a(j) = t
      
       End If
   Next j
Next i

For i = 1 To 5
List2.AddItem Str(a(i))
Next i
End Sub


Private Sub Form_Load()
a(1) = 9: a(2) = 5: a(3) = 8: a(4) = 6: a(5) = 2
For i = 1 To 5
List1.AddItem Str(a(i))
List3.AddItem Str(a(i))
Next
End Sub
Private Sub Command2_Click()
Dim i As Integer
Dim j As Integer
Dim t As Integer

For i = 1 To 4
   For j = i + 1 To 5
       If a(i) > a(j) Then
        t = a(i): a(i) = a(j): a(j) = t
      
       End If
   Next j
Next i

For i = 1 To 5
List4.AddItem Str(a(i))
Next i
End Sub
