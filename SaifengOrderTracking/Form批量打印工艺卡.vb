Public Class Form批量打印工艺卡

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            If TextBox1.Text = "" Then
                MsgBox("单号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * from CustList Order by CustCode"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            rs.MoveFirst()
            Do While rs.EOF = False

                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class