VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "������ϰ��"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6390
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   4920
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ж�"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   5895
      Begin VB.OptionButton Option1 
         Caption         =   "�����ǲƸ�������"
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
         Index           =   3
         Left            =   3120
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�����ܹ�������Ʒ�ļ�ֵ"
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
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "���ҵı�����һ��ȼ���"
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
         Index           =   1
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "��������Ʒ�����м�ֵ"
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
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "1.�����ܳ䵱��ֵ�߶�ְ�ܵĸ���ԭ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "���ε�ѡ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
answer = "��������Ʒ�����м�ֵ"
If Label3.Caption = answer Then
    MsgBox "��ϲ������"
Else
    MsgBox "Ŷno,�����"
End If
Command1.Enabled = False
Command2.Visible = True
End Sub


Private Sub Command2_Click()
End
End Sub

Private Sub Option1_Click(Index As Integer)
Label3.Caption = Option1(Index).Caption

End Sub
