VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ѡ���㷨"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   7800
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�������20������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ѡ������"
      BeginProperty Font 
         Name            =   "����"
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
      ItemData        =   "ѡ������.frx":0000
      Left            =   3120
      List            =   "ѡ������.frx":0002
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   3120
      ItemData        =   "ѡ������.frx":0004
      Left            =   240
      List            =   "ѡ������.frx":0006
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "����������(����)"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ԭʼ����"
      BeginProperty Font 
         Name            =   "����"
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
Dim d(1 To 20) As Integer                  '����һ�����飬��Ŵ���������
Dim n As Integer                           '����n���������ݸ���

Private Sub Command1_Click()               'ѡ������


End Sub


Private Sub Command2_Click()          '����20�������
   n = 20
   Randomize                             '�������ʼ��
   List1.Clear                           '��ԭʼ�����б�����
   List2.Clear                           '�������������б�����
   For i = 1 To n
      d(i) = Int(Rnd * 1000)             'Rnd ��������[0,1)���ʵ��
      List1.AddItem Str(i) & ":" & Str(d(i))            '��������������ʾ��ԭʼ�����б���
   Next i
End Sub

