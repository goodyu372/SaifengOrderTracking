Public Class Form嘉睦订单查询
    Private Sub Form55_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        'Dim SQLstring As String
        Dim SQLstring As String
        'On Error Resume Next
        If (Label1.Text = "按未完成订单查询结果：") Then
            SQLstring = "Select * From JamuOrderBalance where (未交货数>0 and 备货数=0) and 交期<=" & "'" & DateTimePicker1.Value & "'"
        Else
            If (Label1.Text = "按未完成零件查询结果：") Then
                SQLstring = "Select 图号,sum(未交货数) as 缺货数 From JamuOrderBalance where (未交货数>0 and 备货数=0) and 交期<=" & "'" & DateTimePicker1.Value & "' group by 图号"
            End If
        End If

        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
        Next
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
                DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = rs.Fields(i).Value
            Next
            rs.MoveNext()
        Loop
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
    End Sub

    Private Sub 所有未完成订单查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 所有未完成订单查询ToolStripMenuItem.Click
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim da As New System.Data.OleDb.OleDbDataAdapter
        Dim ds As New DataSet
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        Dim SQLstring As String
        SQLstring = "Select * From JMOrderStock where (未交货数>0) order by 序号"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        da.Fill(ds, rs, "us")
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Refresh()
        cn.Close()
        rs = Nothing
        cn = Nothing
        da = Nothing
        ds = Nothing

        StatusStrip1.Items(0).Text = "记录条数：  " & DataGridView1.RowCount
        Label1.Text = "按未完成订单查询结果："
    End Sub

    Private Sub 未完成零件查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 未完成零件查询ToolStripMenuItem.Click
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        'Dim SQLstring As String
        Dim SQLstring As String
        'On Error Resume Next
        SQLstring = "Select * from JamuPartDetail"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If
        rs.MoveFirst()
        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
        Next
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = rs.Fields(i).Value
            Next
            rs.MoveNext()
        Loop
        Label1.Text = "按未完成零件查询结果："
    End Sub

    Private Sub Form55_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub Form55_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 200
    End Sub
    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)

        If (DataGridView1.Columns(e.ColumnIndex).Name = "图号") Then
            SQLstring = "Select * From JMOrderStock where (未交货数>0)  and 图号=" & "'" & DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value & "'"
            rs.Open(SQLstring, cn, 1, 1)
            rs.MoveFirst()
            Form嘉睦订单特别显示.DataGridView1.Rows.Clear()
            Form嘉睦订单特别显示.DataGridView1.Columns.Clear()
            For i = 0 To rs.Fields.Count - 1
                Form嘉睦订单特别显示.DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
            Next
            Form嘉睦订单特别显示.StatusStrip1.Items(0).Text = "记录条数： " & rs.RecordCount
            Do While rs.EOF = False
                Form嘉睦订单特别显示.DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    Form嘉睦订单特别显示.DataGridView1.Item(i, Form嘉睦订单特别显示.DataGridView1.Rows.Count - 1).Value = IIf(IsDBNull(rs(i).Value), "", rs(i).Value)
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Me.Hide()
            Form嘉睦订单特别显示.Show()
            Form嘉睦订单特别显示.WindowState = FormWindowState.Maximized
        End If
        If (DataGridView1.Columns(e.ColumnIndex).Name = "客户订单号") Then
            SQLstring = "Select * From JMOrderStock where (未交货数>0)  and 客户订单号=" & "'" & DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value & "'"
            rs.Open(SQLstring, cn, 1, 1)
            rs.MoveFirst()
            Form嘉睦订单特别显示.DataGridView1.Rows.Clear()
            Form嘉睦订单特别显示.DataGridView1.Columns.Clear()
            For i = 0 To rs.Fields.Count - 1
                Form嘉睦订单特别显示.DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
            Next
            Form嘉睦订单特别显示.StatusStrip1.Items(0).Text = "记录条数： " & rs.RecordCount
            Do While rs.EOF = False
                Form嘉睦订单特别显示.DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    Form嘉睦订单特别显示.DataGridView1.Item(i, Form嘉睦订单特别显示.DataGridView1.Rows.Count - 1).Value = IIf(IsDBNull(rs(i).Value), "", rs(i).Value)
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Me.Hide()
            Form嘉睦订单特别显示.Show()
            Form嘉睦订单特别显示.WindowState = FormWindowState.Maximized
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class