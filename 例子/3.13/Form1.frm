VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "����100�׵ȼ���ѯ"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7530
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
   ScaleHeight     =   4035
   ScaleWidth      =   7530
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��ѯ"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   555
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "�ȼ�����"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "�ɼ����룩"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
t = Val(Text1.Text)
If t <= 10.25 Then
    Text2.Text = "���Ҽ��˶�����"
ElseIf t <= 10.5 Then
    Text2.Text = "�˶�����"
ElseIf t <= 10.93 Then
    Text2.Text = "һ���˶�Ա"
ElseIf t <= 11.74 Then
    Text2.Text = "�����˶�Ա"
Else
    Text2.Text = "û�еȼ�"
End If
End Sub
