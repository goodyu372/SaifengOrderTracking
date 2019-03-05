Public Class Form图纸信息工艺卡

    Private Sub Form10_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Select Case ProcessForm
            Case "form4"
                Form订单查询.Show()
            Case "form17"
                Form按客户查询未完成零件状况.Show()
            Case "form16"
                Form按生产单号查询未完成零件状况.Show()
            Case "form18"
                Form按图号查询未完成零件.Show()
            Case "form29"
                Form所有未完成生产单.Show()
            Case "form90"
                Form按客户订单号查询生产单.Show()
        End Select
    End Sub

    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i = 1 To DataGridView1.Columns.Count - 1
            If ((DataGridView1.Columns(i).Name = "ID") Or (DataGridView1.Columns(i).Name = "CustDWG") Or (DataGridView1.Columns(i).Name = "SFOrder") Or (DataGridView1.Columns(i).Name = "DrawNo") Or (DataGridView1.Columns(i).Name = "PartName") Or (DataGridView1.Columns(i).Name = "SPOrder") Or (DataGridView1.Columns(i).Name = "ProcMaker") Or (DataGridView1.Columns(i).Name = "ProcDate") Or (DataGridView1.Columns(i).Name = "CustCode") Or (DataGridView1.Columns(i).Name = "ProcCode")) Then
                DataGridView1.Columns(i).Visible = False
            End If
        Next
    End Sub
End Class