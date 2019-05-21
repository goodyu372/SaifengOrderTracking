Public Class Form嘉睦零件最新信息查询
    Public ss As New Microsoft.Office.Interop.Excel.Application
    Public xlsheet As Microsoft.Office.Interop.Excel.Worksheet
    Public xlbook As Microsoft.Office.Interop.Excel.Workbook
    Private Sub Form46_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form46_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 100
    End Sub
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim cn As New ADODB.Connection
        Dim rs, rs1 As New ADODB.Recordset
        Dim SQLstring As String
        Dim Dash As Integer
        cn.Open(CNsfmdb)

        If (DataGridView1.Columns(e.ColumnIndex).Name = "图号") Then
            Dash = InStr(DataGridView1.Item("图号", e.RowIndex).Value, "-")
            SQLstring = "Select * from SFOrderBase where DrawNo like " & "'%" & Mid(DataGridView1.Item("图号", e.RowIndex).Value, 1, Len(DataGridView1.Item("图号", e.RowIndex).Value) - Dash) & "%'"

            Form嘉睦零件生产状态查询.DataGridView1.Rows.Clear()
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("图号不存在，请检查！")
                Exit Sub
            End If
            rs.Close()
            SQLstring = "Select * from StatusInfo where ((Status not in (select Status where Status='完成入库' or Status='取消生产单' or Status='已终检')) and (Drawno like" & "'%" & Mid(DataGridView1.Item("图号", e.RowIndex).Value, 1, Dash - 1) & "%'))"
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此图号零件全部入库，请检查！")
                Exit Sub
            End If
            rs.MoveFirst()
            For i = 0 To rs.Fields.Count - 1
                Form嘉睦零件生产状态查询.DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
            Next
            Do While Not rs.EOF
                Form嘉睦零件生产状态查询.DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    Form嘉睦零件生产状态查询.DataGridView1.Item(i, Form嘉睦零件生产状态查询.DataGridView1.Rows.Count - 1).Value = IIf(IsDBNull(rs(i).Value), "", rs(i).Value)
                Next
                rs.MoveNext()
            Loop

            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Me.Hide()
            Form嘉睦零件生产状态查询.Show()
            Form嘉睦零件生产状态查询.WindowState = FormWindowState.Maximized
        End If
        If (DataGridView1.Columns(e.ColumnIndex).Name = "缺货数") Then
            Dash = InStr(DataGridView1.Item("图号", e.RowIndex).Value, "-")
            SQLstring = "Select * From JMOrderStock where (未交货数>0) and (图号 like" & "'%" & Mid(DataGridView1.Item("图号", e.RowIndex).Value, 1, Dash - 1) & "%')"

            Form嘉睦零件未完成订单信息.DataGridView1.Rows.Clear()
            Form嘉睦零件未完成订单信息.DataGridView1.Columns.Clear()
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("不存在未完成的该图号订单，请检查！")
                Exit Sub
            End If
            rs.MoveFirst()
            For i = 0 To rs.Fields.Count - 1
                Form嘉睦零件未完成订单信息.DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
            Next
            Do While Not rs.EOF
                Form嘉睦零件未完成订单信息.DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    Form嘉睦零件未完成订单信息.DataGridView1.Item(i, Form嘉睦零件未完成订单信息.DataGridView1.Rows.Count - 1).Value = IIf(IsDBNull(rs(i).Value), "", rs(i).Value)
                Next
                rs.MoveNext()
            Loop

            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Me.Hide()
            Form嘉睦零件未完成订单信息.Show()
            Form嘉睦零件未完成订单信息.WindowState = FormWindowState.Maximized
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            If TextBox1.Text = "" Then
                MsgBox("请输入图号！")
                Exit Sub
            End If
            Dim da As New System.Data.OleDb.OleDbDataAdapter
            Dim ds As New DataSet
            DataGridView1.DataSource = Nothing
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()

            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * from JamuPartDetail where 图号 like " & "'%" & TextBox1.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)

            StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            End If

            da.Fill(ds, rs, "us")
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.Refresh()
            For i = 0 To DataGridView1.Rows.Count - 1
                If (DataGridView1.Item("差额", i).Value < 0) Then
                    DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Red
                End If
            Next
            cn.Close()
            rs = Nothing
            cn = Nothing
            da = Nothing
            ds = Nothing
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim da As New System.Data.OleDb.OleDbDataAdapter
        Dim ds As New DataSet
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * from JamuPartDetail"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)

        StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If

        da.Fill(ds, rs, "us")
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Refresh()
        cn.Close()
        rs = Nothing
        cn = Nothing
        da = Nothing
        ds = Nothing

        For i = 0 To DataGridView1.Rows.Count - 1
            If (DataGridView1.Item("差额", i).Value < 0) Then
                DataGridView1.Rows(i).DefaultCellStyle.ForeColor = Color.Red
            End If
        Next
        WindowState = FormWindowState.Maximized
    End Sub
End Class