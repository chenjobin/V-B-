VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ð�������㷨������ʵ��"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ʼð������>>>"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<<�Զ���������"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ListBox List2 
      Height          =   4545
      ItemData        =   "ð������.frx":0000
      Left            =   5040
      List            =   "ð������.frx":0002
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   4545
      ItemData        =   "ð������.frx":0004
      Left            =   120
      List            =   "ð������.frx":0006
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "������������"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d(1 To 1000) As Integer
Dim num As Integer

Private Sub Command1_Click() '�����������ť
num = Val(Text1.Text)
If num <= 0 Or num >= 1000 Then
   Text1.Text = "���ݸ���������"
   Exit Sub
End If
Randomize             '��ʼ���������
List1.Clear
For i = 1 To num
 j = Round(Rnd * 1000)  'rnd��������һ��[0��1)֮��������
 List1.AddItem Str(j)
 d(i) = j
 Next
End Sub

Private Sub Command2_Click()  'ð������ť

   List2.Clear
For i = 1 To num - 1     'ð������ ����
   For j = num To i + 1 Step -1    '�ӵ�����ð��
       If d(j) < d(j - 1) Then     ' ����¸�����С��������һ�����ݽ���
         temp = d(j)
         d(j) = d(j - 1)
         d(j - 1) = temp
        End If
    Next j
Next i


For i = 1 To num       '���б�2����ʾ����������
 List2.AddItem d(i)
Next i
End Sub
