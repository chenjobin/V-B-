VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��°ͺղ���"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   9135
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
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
      Left            =   6960
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�Ӻ�������"
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
      Left            =   3840
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������"
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
      Left            =   480
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
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
      Left            =   5760
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   3000
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "���ʽ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "������һ������2��ż��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "�κδ���2��ż����������֮�ͣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_click()    '���ú������̵�������
    Dim i, S, N As Integer
    N = Text1.Text                        '���ı����е����ָ�������N
    For i = 3 To N Step 2                     '����3��N�м������
        S = i                            '������I��ֵ��������S
        If HQSS(S) = 1 Then              '�������ֵΪ1����˵��I������
            S = N - i                    '��N-1��ֵ��������S
            If HQSS(S) = 1 Then Exit For    '�������ֵΪ1��������ѭ��
        End If
    Next i
    S = HSJG(i, N)                          '������ʾ��������
    If N = 4 Then Label4.Caption = "4=2+2"   '���N=4��������ʾ���
End Sub

Public Function HQSS(ByVal S1 As Integer)   '�ж��Ƿ�Ϊ�����ĺ�������
    Dim J As Integer
    K1 = 1                              '����K1Ϊ1ʱ����������������
    For J = 2 To Sqr(S1)                '���������ж�
        If S1 / J = Int(S1 / J) Then K1 = 0: Exit For   '���������������Ϊ�������˳�ѭ��
    Next J
    HQSS = K1       '���жϽ��������
End Function

Public Function HSJG(ByVal i As Integer, ByVal N As Integer)
    Label4.Caption = Str(N) + "=" + Str(i) + "+" + Str(N - i)
End Function

Private Sub command2_click()    '�����ӹ��̵�������
Dim i, S, N, K As Integer
N = Text1.Text
For i = 3 To N Step 2
    S = i: Call ZQSS(S, K)       '���ж������ӳ���
    If K = 1 Then
        S = N - 1: Call ZQSS(S, K)      '���ж������ӳ���
        If K = 1 Then Exit For
    End If
Next i
Call XSJG(i, N)     '����ʾ�ӳ���
If N = 4 Then Label4.Caption = "4=2+2"
End Sub
Public Sub ZQSS(ByVal S1 As Integer, K1 As Integer)  '�ж��Ƿ�Ϊ�������ӹ���
Dim J As Integer
K1 = 1
For J = 2 To Sqr(S1)
    If S1 / J = Int(S1 / J) Then K1 = 0: Exit For
Next J
End Sub
Public Sub XSJG(ByVal i As Integer, ByVal N As Integer)
    Label4.Caption = Str(N) + "=" + Str(i) + "+" + Str(N - i)
End Sub
Private Sub command3_click()
    End
End Sub
