VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ש������"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4785
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   3120
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "����Բ�뾶��"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim r As Single
    Dim k, c, v As Integer
    List1.Clear
    r = Val(Text1.Text)
    k = Int(r)
    c = 0
    For j = 1 To k
        v = Int(Sqr(r ^ 2 - j ^ 2))
        c = c + v
    Next j
    List1.AddItem ("�����߳�Ϊ1��ש����")
    List1.AddItem (Str(4 * c) + "��")
End Sub

