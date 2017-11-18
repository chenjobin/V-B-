Attribute VB_Name = "Module1"
Public Sub dlcg()
   MsgBox "恭喜你登录成功！"
   Form1.Text1.Text = ""
   Form1.Text1.SetFocus '设置焦点
   Form1.Text2.Text = ""
End Sub
Public Sub dlsb()
   MsgBox "账号或密码错误！请重新输入！", vbOKOnly, "重输口令"
   Form1.Text1.Text = ""
   Form1.Text1.SetFocus '设置焦点
   Form1.Text2.Text = ""
End Sub
