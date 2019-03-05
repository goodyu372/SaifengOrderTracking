Public Class Form个人加工详单

    Private Sub Form79_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form工序个人工时查询.Show()
    End Sub

    Private Sub Form79_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form79_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 200
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If DataGridView1.Columns(e.ColumnIndex).Name = "ProcCode" Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SQLstring = "Select * from ProcCard where ProcCode=" & "'" & DataGridView1.Item("ProcCode", e.RowIndex).Value & "'"
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有工艺卡！")
                Exit Sub
            End If
            rs.MoveFirst()
            Form工序卡信息查询.DataGridView1.Rows.Clear()
            Form工序卡信息查询.DataGridView1.Columns.Clear()
            For i = 0 To rs.Fields.Count - 1
                Form工序卡信息查询.DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
            Next
            Do While rs.EOF = False
                Form工序卡信息查询.DataGridView1.Rows.Add(1)
                For j = 0 To DataGridView1.Columns.Count - 1
                    Form工序卡信息查询.DataGridView1.Item(j, Form工序卡信息查询.DataGridView1.Rows.Count - 1).Value = rs(j).Value
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Me.Hide()
            Form工序卡信息查询.Show()
        End If
    End Sub
End Class