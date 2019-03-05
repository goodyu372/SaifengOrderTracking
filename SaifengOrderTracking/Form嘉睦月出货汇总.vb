Public Class Form嘉睦月出货汇总

    Private Sub Form72_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            DataGridView1.Rows.Clear()
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring As String
            'SQLstring = "Select *,(100*year(送货日期)+month(送货日期)) as YearMonth from JamuOutput where (100*year(送货日期)+month(送货日期)) =" & "'" & TextBox1.Text & "'"
            SQLstring = "Select * from JamuOutputYM where YearMonth =" & "'" & TextBox1.Text & "'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)

            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有该年月的出货！")
                Exit Sub
            End If
            StatusStrip1.Items(0).Text = "本月出货款数：   " & rs.RecordCount
            rs.MoveFirst()
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
            Next
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(rs(i).Name, DataGridView1.Rows.Count - 1).Value = rs(i).Value
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            'SQLstring = "Select sum(总价) as Summary from JamuOutput where (100*year(送货日期)+month(送货日期)) =" & "'" & TextBox1.Text & "'"
            SQLstring = "Select sum(总价) as Summary from JamuOutputYM where YearMonth =" & Val(TextBox1.Text)
            rs.Open(SQLstring, cn, 1, 1)
            DataGridView1.Rows.Add(1)
            DataGridView1.Item("总价", DataGridView1.Rows.Count - 1).Value = rs("Summary").Value
            DataGridView1.Item("图号", DataGridView1.Rows.Count - 1).Value = "月出货汇总"
            cn.Close()
            rs = Nothing
            cn = Nothing
            DataGridView1.FirstDisplayedScrollingRowIndex = DataGridView1.Rows.Count - 1
        End If
    End Sub

    Private Sub Form72_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class