Public Class Form嘉睦零件生产状态查询

    Private Sub Form47_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form嘉睦零件最新信息查询.Show()
    End Sub

    Private Sub Form47_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form47_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 100
    End Sub
End Class