Public Class Form用户登录
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = 13) Then
            TextBox2.Enabled = True
            TextBox2.Focus()
        End If
    End Sub
    Private Sub Form3_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.Enabled = False
        TextBox2.PasswordChar = "*"
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If (e.KeyCode = 13) Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SQLstring = "Select UserName,PassWord,UserType,ChineseName from UserInfo where UserName=" & "'" & TextBox1.Text & "'"

            rs.Open(SQLstring, cn, 1, 2)
            If (rs.RecordCount = 0) Then
                MsgBox("不存在此用户！")
                TextBox1.Clear()
                TextBox2.Clear()
                TextBox1.Focus()
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            Else
                If (rs("PassWord").Value <> TextBox2.Text) Then
                    MsgBox("用户密码不正确！")
                    TextBox2.Clear()
                    TextBox2.Focus()
                    rs.Close()
                    cn.Close()
                    rs = Nothing
                    cn = Nothing
                    Exit Sub
                Else
                    UserName = TextBox1.Text
                    UserType = rs("UserType").Value
                    ChineseName = rs("ChineseName").Value
                    Form主窗体.StatusStrip1.Items(2).Text = "          用户名：" + ChineseName
                    LoginUser(UserType)
                    rs.Close()
                    cn.Close()
                    Me.Close()
                    Form主窗体.Show()
                End If
            End If
        End If
    End Sub
End Class