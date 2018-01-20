VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   6420
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1455
      Left            =   600
      TabIndex        =   0
      Top             =   2040
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
k = 1
N = Val(Text1.Text)
For i = 2 To N - 1 Step 1
If N Mod i = 0 Then k = 0
Next i
If k = 1 Then
MsgBox Str(N) + "是素数"
Else
MsgBox Str(N) + "不是素数"
End If
End Sub
