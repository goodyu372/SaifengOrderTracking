Public Class Form嘉睦订单特别显示

    Private Sub Form56_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form嘉睦订单查询.Show()
    End Sub
    Private Sub Form56_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 150
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            If TextBox1.Text = "" Then
                MsgBox("图号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * From JMOrderStock where (未交货数>0) and 图号 like " & "'%" & TextBox1.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            StatusStrip1.Items(0).Text = "" & rs.RecordCount
            DataGridView1.Rows.Clear()
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            End If
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = IIf(IsDBNull(rs(i).Value), "", rs(i).Value)
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            If TextBox2.Text = "" Then
                MsgBox("客户订单号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * From JMOrderStock where (未交货数>0) and 客户订单号 like " & "'%" & TextBox2.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            StatusStrip1.Items(0).Text = "" & rs.RecordCount
            DataGridView1.Rows.Clear()
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            End If
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = IIf(IsDBNull(rs(i).Value), "", rs(i).Value)
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If

    End Sub
    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * From JamuStock where 图号=" & "'" & DataGridView1.Item("图号", e.RowIndex).Value & "'"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        StatusStrip1.Items(0).Text = "" & rs.RecordCount
        Form嘉睦零件库存删除.DataGridView1.Rows.Clear()
        Form嘉睦零件库存删除.DataGridView1.Columns.Clear()
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If
        For i = 0 To rs.Fields.Count - 1
            Form嘉睦零件库存删除.DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
        Next
        Do While rs.EOF = False
            Form嘉睦零件库存删除.DataGridView1.Rows.Add(1)
            For i = 0 To rs.Fields.Count - 1
                Form嘉睦零件库存删除.DataGridView1.Item(i, Form嘉睦零件库存删除.DataGridView1.Rows.Count - 1).Value = rs(i).Value
            Next
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        JamuStockForm = "Form56"
        Me.Hide()
        Form嘉睦零件库存删除.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '更新备货标记
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("没有记录！")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        'Dim SQLstring As String
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        For i = 0 To DataGridView1.Rows.Count - 1
            SQLstring = "Select 备货标记 from JamuOrder where 序号=" & "'" & DataGridView1.Item("序号", i).Value & "'"
            rs.Open(SQLstring, cn, 1, 3)
            rs.Fields("备货标记").Value = DataGridView1.Item("备货标记", i).Value
            rs.Update()
            rs.Close()
        Next
        cn.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("备货标记更新完成！")
    End Sub
End Class