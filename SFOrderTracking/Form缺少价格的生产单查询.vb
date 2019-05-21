Public Class Form缺少价格的生产单查询

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select * from SumOrder where OrderSum=0 and YearMonth>201805 and CustCode<>'C001' and CustCode<>'C002' order by SFOrder desc"
        rs.Open(SQLstring, cn, 1, 1)
        If rs.RecordCount < 1 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            MsgBox("不存在无价格的订单！")
            Exit Sub
        End If
        rs.MoveFirst()
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Visible = True
        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
        Next
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = rs(i).Value
            Next
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub
    Private Sub Form84_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form84_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class