VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�Ƚ���������ʾ�ϴ���"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "�ϴ������"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "��������������ͬ������"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Let a = Text1.Text '�������ı���1�е��˺Ÿ�������a
  Let b = Text2.Text '�������ı���2�е����븳������b
  If a > b Then
    Text3.Text = a
  Else
    Text3.Text = b
  End If
End Sub
Private Sub Command2_Click()
  End
End Sub

