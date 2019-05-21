Public Class Form个人加工详单

    Private Sub Form79_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Form工序个人工时查询.Show()
    End Sub

    Private Sub Form79_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form79_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 200
    End Sub

End Class