VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7260
   ScaleWidth      =   14325
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command8 
      Caption         =   "��������"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6600
      TabIndex        =   10
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "������"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6600
      TabIndex        =   9
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�ƽ���ĸ����"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�ƽ�5λ���������"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�ƽ�4λ������"
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "����4λ��ĸ����"
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����5λ���������"
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����4λ������"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   2640
      Picture         =   "Form1.frx":6BF37
      Stretch         =   -1  'True
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2640
      TabIndex        =   6
      Top             =   4680
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Long
Dim b As String
Private Sub Command1_Click()      '����4λ������1234
a = 1234
Command4.Enabled = True
End Sub

Private Sub Command4_Click()          '�ƽ�4λ������
For i = 0 To 9999
If i = a Then
Label1.Caption = "������" + Str(i)
End If
Next i
Command2.Enabled = True
End Sub

Private Sub Command2_Click()      '����5λ�����������
Randomize
a = Int(Rnd * 90000 + 10000)
Command5.Enabled = True
End Sub

Private Sub Command5_Click()           '�ƽ�5λ������
For i = 10000 To 99999
If i = a Then
Label2.Caption = "������" + Str(i)
End If
Next i
Command3.Enabled = True
End Sub

Private Sub Command3_Click()           '������ĸ����fish

End Sub

Private Sub Command6_Click()               '�ƽ���ĸ����fish


End Sub

Private Sub Command7_Click()                      '������
Form1.Picture = LoadPicture("lm.jpg")
Image1.Visible = True
End Sub


