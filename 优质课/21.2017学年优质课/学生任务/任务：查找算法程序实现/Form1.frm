VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�����㷨�ĳ���ʵ��"
   ClientHeight    =   5370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�Էֲ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "˳�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Height          =   735
      Left            =   2640
      TabIndex        =   11
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�������ݣ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���������ݣ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����ǰ�����ݣ�"
      BeginProperty Font 
         Name            =   "����"
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
Dim sorted As Boolean          '�ж��Ƿ��Ѿ�����
Dim flag As Boolean          '�������߼������ķ�ʽ����ѭ�������������в�ʹ��Exit��䣩
Private Sub Command1_Click()   '˳�����
  key = Val(Text1.Text)               '���ҵ�����
  For i =_______________                   '���β���
   
   If _______________ Then                '�ҵ�������
     
      Label4.Caption = "������ĵ�" + Str(i) + "��λ��"
      ______________
    End If
    Next i
    If i =_____________ Then         '�ж�δ�ҵ�����
      Label4.Caption = "��������û���ҵ�����" + Str(key)
    End If
End Sub


Private Sub Command2_Click()          '�Էֲ���
 Dim i As Integer, j As Integer, key As Integer, m As Integer
  If Not sorted Then                                 '���жԷֲ���ʱ�����ȶ����ݽ�������
      MsgBox "���жԷֲ���ʱ�����ȶ����ݽ�������"
      Exit Sub
  End If
   '�Էֲ���
   key = Val(Text1.Text)             '���ҵ�����
   i = 1: j = n
   Do While ______________
    
    m = _________________          '�м�λ��m�ļ���
   If d(m) = key Then
       Label5.Caption = "������ĵ�" + Str(m) + "��λ��"
     Exit Do
   End If
   If d(m) < key Then            '�ж����ݵ�λ����ǰ�벿�ֻ��Ǻ�벿��
      
      _______________________
    
    Else
    
      _______________________
      
    End If
  Loop
  if _____________ Then
     
     Label5.Caption = "��������û���ҵ�����" + Str(key)
     
  End If

End Sub

Private Sub Command3_Click()

   Dim i As Integer, j As Integer, temp As Integer
   List2.Clear
   sorted = True                     '����������
   For i = 1 To n - 1              'ð������ ����
     For j = n To i + 1 Step -1    '��������ð��
       If d(j) < d(j - 1) Then       '�����һ������С����ͬ��һ�����ݽ��н���
         temp = d(j)
         d(j) = d(j - 1)
         d(j - 1) = temp
        End If
     Next j
   Next i

For i = 1 To n                   '���ұ��б�����ʾ����������
   List2.AddItem Str(i) + " :" + Str(d(i))
Next i
End Sub

Private Sub Command4_Click()
Randomize             '��ʼ���������
sorted = False        '�Ƿ��Ѿ�����
 List1.Clear
  For i = 1 To 10
    d(i) = 1 + Int(Rnd * 100)
    List1.AddItem Str(i) + " :" + Str(d(i))
  Next
End Sub

