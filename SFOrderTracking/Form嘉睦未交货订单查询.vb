Public Class Form嘉睦未交货订单查询
    Private Sub Form54_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            Dim da As New System.Data.OleDb.OleDbDataAdapter
            Dim ds As New DataSet
            DataGridView1.DataSource = Nothing
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()

            If TextBox1.Text = "" Then
                MsgBox("请输入图号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * From JMOrderStock where 图号 like " & "'%" & TextBox1.Text & "%'" & " and 未交货数>0"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)

            da.Fill(ds, rs, "us")
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.Refresh()
            cn.Close()
            rs = Nothing
            cn = Nothing
            da = Nothing
            ds = Nothing

            For i = 0 To DataGridView1.Columns.Count - 1
                If (DataGridView1.Columns(i).Name = "单价") Or (DataGridView1.Columns(i).Name = "金额") Or (DataGridView1.Columns(i).Name = "备货数") Or (DataGridView1.Columns(i).Name = "图号单号") Then
                    DataGridView1.Columns(i).Visible = False
                End If
            Next
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        '按客户订单号
        If e.KeyCode = 13 Then
            Dim da As New System.Data.OleDb.OleDbDataAdapter
            Dim ds As New DataSet
            DataGridView1.DataSource = Nothing
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()

            If TextBox2.Text = "" Then
                MsgBox("请输入客户订单号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * From JMOrderStock where 客户订单号 like " & "'%" & TextBox2.Text & "%'" & " and 未交货数>0"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)

            da.Fill(ds, rs, "us")
            DataGridView1.DataSource = ds.Tables(0)
            DataGridView1.Refresh()
            cn.Close()
            rs = Nothing
            cn = Nothing
            da = Nothing
            ds = Nothing

            For i = 0 To DataGridView1.Columns.Count - 1
                If (DataGridView1.Columns(i).Name = "单价") Or (DataGridView1.Columns(i).Name = "金额") Or (DataGridView1.Columns(i).Name = "备货数") Or (DataGridView1.Columns(i).Name = "图号单号") Then
                    DataGridView1.Columns(i).Visible = False
                End If
            Next
        End If
    End Sub
    Private Sub Form54_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 200
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
        SQLstring = "Select * From JMOrderStock where (未交货数>0) order by 序号"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        StatusStrip1.Items(0).Text = "记录条数： " & rs.RecordCount

        da.Fill(ds, rs, "us")
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Refresh()
        cn.Close()
        rs = Nothing
        cn = Nothing
        da = Nothing
        ds = Nothing

        For i = 0 To DataGridView1.Columns.Count - 1
            If (DataGridView1.Columns(i).Name = "单价") Or (DataGridView1.Columns(i).Name = "金额") Or (DataGridView1.Columns(i).Name = "备货数") Or (DataGridView1.Columns(i).Name = "图号单号") Then
                DataGridView1.Columns(i).Visible = False
            End If
        Next
    End Sub
    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        If (((DataGridView1.Columns(e.ColumnIndex).Name = "DrawNo") Or (DataGridView1.Columns(e.ColumnIndex).Name = "图号")) And (e.RowIndex >= 0)) Then
            Shell(SFOrderBaseNetConnect, vbHide)
            '88888888888888888888888888 Select DrawNo to display 88888888888888888888888888888
            Dim sF As String
            Dim Path, SearchedFile As String
            Path = DWGPath
            If (Mid(DataGridView1.Item("图号", e.RowIndex).Value, Len(DataGridView1.Item("图号", e.RowIndex).Value) - 2, 3) = "-CH") Then
                SearchedFile = Mid(DataGridView1.Item("图号", e.RowIndex).Value, 1, Len(DataGridView1.Item("图号", e.RowIndex).Value) - 3)
            Else
                SearchedFile = DataGridView1.Item("图号", e.RowIndex).Value
            End If

            sF = Dir(Path, vbDirectory) ' 查找目录中第一个文件夹名称
            Do While sF <> "" ' 跳过当前的目录及上层目录
                If (InStr(1, UCase(sF), SearchedFile) > 0) Then
                    System.Diagnostics.Process.Start(Path + sF)
                    Exit Do
                End If
                sF = Dir()
            Loop
            Exit Sub
            '88888888888888888888 end of Select DrawNo to display 8888888888888888888888888888
            Shell(SFOrderBaseNetDisconnect, vbHide)
        End If
    End Sub
End Class