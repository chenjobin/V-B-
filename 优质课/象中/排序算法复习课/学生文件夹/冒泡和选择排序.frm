VERSION 5.00
Begin VB.Form ð�ݺ������㷨 
   Caption         =   "ð�ݺ������㷨"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   7800
   StartUpPosition =   3  '����ȱʡ
   Begin VB.ListBox List3 
      Height          =   3120
      Left            =   3720
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ѡ������"
      Height          =   615
      Left            =   5640
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ð������"
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�������8����"
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      Height          =   3120
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3120
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "ѡ��������"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ð��������"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ԭʼ����"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "ð�ݺ������㷨"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim a(1 To 8) As Integer  '����һ������


Private Sub Command1_Click()  '����8�������
   Randomize  '�������ʼ��
List1.Clear 'ԭʼ�������
List2.Clear   '���������б��������
List3.Clear   '���������б��������
For i = 1 To 8
   a(i) = Int(Rnd * 100) 'Rnd �������ص���������� 0 �� 1 ֮�䣬�ɵ��� 0���������� 1
   List1.AddItem Str(a(i))  '��������ʾ��ԭʼ�����б���
Next i

End Sub

Private Sub Command2_Click()  '��8��������ð�ݷ�����


End Sub

Private Sub Command3_Click() '��8��������ѡ������

   
End Sub

