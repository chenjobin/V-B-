VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���ڼ�����"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   7605
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ����"
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "������:"
      Height          =   975
      Left            =   720
      TabIndex        =   9
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "��"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "��"
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "��"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�������ڣ�"
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
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function leap(y As Integer) As Integer '�ж��Ƿ�Ϊ���꺯��
  If y Mod 100 = 0 Then   '�Ƿ���1�����Ƿ���0
    If y Mod 400 = 0 Then leap = 1 Else leap = 0
  Else
    If y Mod 4 = 0 Then leap = 1 Else leap = 0
  End If
End Function
Private Sub Command1_Click()
Dim year As Integer, month, day, w, c, y, m, d, ok As Integer
year = Val(Text1.Text)
month = Val(Text2.Text)
day = Val(Text3.Text)
ok = 0     '�ж�����������Ƿ���Ч������ƽ�ꡢ���ꡢ��ͬ�·ݵĲ�ͬ������ж�
If (year >= 1900) And (month = 1 Or month = 3 Or month = 5 Or month = 7 Or month = 8 Or month = 10 Or month = 12) And (day >= 1 And day <= 31) Then ok = 1
If (year >= 1900) And (month = 4 Or month = 6 Or month = 9 Or month = 11) And (day >= 1 And day <= 30) Then ok = 1
If (year >= 1900) And leap(year) = 1 And month = 2 And (day >= 1 And day <= 29) Then ok = 1
If (year >= 1900) And leap(year) = 0 And month = 2 And (day >= 1 And day <= 28) Then ok = 1
If ok = 0 Then
 Text4.Text = "���������д���"
Else
  If month = 1 Or month = 2 Then
    year = year - 1
    month = month + 12
  End If
  c = year \ 100   'ȡ��ݵ�ǰ��λ
  y = year Mod 100   'ȡ��ݵĺ���λ
  m = month
  d = day
  w = Int(c / 4) - 2 * c + y + Int(y / 4) + Int(13 * (m + 1) / 5) + d - 1   '���չ�ʽ
  Text4.Text = Str(c) + Str(y) + Str(m) + Str(d) + Str(w)
  w = (w + 700) Mod 7 + 1  '�����7��������w����700��֤����һ��������
  Text4.Text = WeekdayName(w)  'WeekdayName�������Խ���ֵת����������ʽ��ֵ1ת���������գ�ֵ2ת��������һ������ֵ7ת����������
 End If
End Sub

Private Sub Text1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub
