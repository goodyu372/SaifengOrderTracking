Public Class Form嘉睦零件未完成订单信息

    Private Sub Form48_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form嘉睦零件最新信息查询.Show()
    End Sub

    Private Sub Form48_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 100
    End Sub
End Class