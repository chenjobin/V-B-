VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�����"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "����"
      Size            =   21.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   5475
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�Ƚ�"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   675
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "HHH"
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "����ߣ�cm��"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "����ߣ�cm��"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'�ֱ��ȡ�������ת������ڱ���a��b��
a = Val(Text1.Text)
b = Val(Text2.Text)

If a > b Then
    Max = a
Else
    Max = b
End If
Label4.Caption = "�ϸߵ�ͬѧ�������" + Str(Max) + "cm"
End Sub

