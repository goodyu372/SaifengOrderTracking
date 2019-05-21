Public Class Form外发记录查询

    Private Sub 查询外发未收回ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 查询外发未收回ToolStripMenuItem.Click
        Label1.Text = "已外发未收回清单："
        Label2.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        TextBox1.Visible = True
        TextBox2.Visible = True
        ComboBox1.Visible = True
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        DataGridView1.Rows.Clear()
        cn.Open(CNsfmdb)
        SQLstring = "Select * from SubRecord where isnull(BackDate,'')=''"
        rs.Open(SQLstring, cn, 1, 1)
        StatusStrip1.Items(0).Text = "记录条数:  " & rs.RecordCount
        If (rs.RecordCount = 0) Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If
        DataGridView1.Rows.Add(rs.RecordCount)
        rs.MoveFirst()
        Dim DataRow As Integer
        DataRow = 0
        Do While rs.EOF = False
            DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
            DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
            DataGridView1.Item("OKQty", DataRow).Value = rs.Fields("OKQty").Value
            DataGridView1.Item("SubCompany", DataRow).Value = rs.Fields("SubCompany").Value
            DataGridView1.Item("UnitPrice", DataRow).Value = rs.Fields("UnitPrice").Value
            DataGridView1.Item("SubDate", DataRow).Value = rs.Fields("SubDate").Value
            DataGridView1.Item("BackDate", DataRow).Value = rs.Fields("BackDate").Value
            DataGridView1.Item("Comment", DataRow).Value = rs.Fields("Comment").Value
            DataGridView1.Item("Operator", DataRow).Value = rs.Fields("Operator").Value
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop
        rs.Close()
        SQLstring = "Select SubCompany from SubRecord where isnull(BackDate,'')='' group by SubCompany Order by SubCompany"
        rs.Open(SQLstring, cn, 1, 1)
        ComboBox1.Items.Clear()
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If
        Do While rs.EOF = False
            ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing

    End Sub
    Private Sub Form26_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Select Case SubReordQuery
            Case "PMC"
                Form主窗体.Show()
            Case "NormalQuery"
                Form订单查询.Show()
            Case "SpecialQuery"
                Form主窗体.Show()
        End Select

    End Sub

    Private Sub Form26_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("OKQty", "数量")
        DataGridView1.Columns.Add("SubCompany", "供应商")
        DataGridView1.Columns.Add("UnitPrice", "外发单价")
        DataGridView1.Columns.Add("Comment", "外发项目备注")
        DataGridView1.Columns.Add("SubDate", "外发日期")
        DataGridView1.Columns.Add("BackDate", "收回日期")
        DataGridView1.Columns.Add("Operator", "经手人")
    End Sub

    Private Sub 查询外发已收回ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 查询外发已收回ToolStripMenuItem.Click
        Label1.Text = "外发已收回清单："
        Label2.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        TextBox1.Visible = True
        TextBox2.Visible = True
        ComboBox1.Visible = True
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        DataGridView1.Rows.Clear()
        cn.Open(CNsfmdb)
        SQLstring = "Select * from SubRecord where isnull(BackDate,'')<>''"
        rs.Open(SQLstring, cn, 1, 1)
        If (rs.RecordCount = 0) Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If
        DataGridView1.Rows.Add(rs.RecordCount)
        StatusStrip1.Items(0).Text = "记录条数:  " & rs.RecordCount
        rs.MoveFirst()
        Dim DataRow As Integer
        DataRow = 0
        Do While rs.EOF = False
            DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
            DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
            DataGridView1.Item("OKQty", DataRow).Value = rs.Fields("OKQty").Value
            DataGridView1.Item("SubCompany", DataRow).Value = rs.Fields("SubCompany").Value
            DataGridView1.Item("UnitPrice", DataRow).Value = rs.Fields("UnitPrice").Value
            DataGridView1.Item("SubDate", DataRow).Value = rs.Fields("SubDate").Value
            DataGridView1.Item("BackDate", DataRow).Value = rs.Fields("BackDate").Value
            DataGridView1.Item("Comment", DataRow).Value = rs.Fields("Comment").Value
            DataGridView1.Item("Operator", DataRow).Value = rs.Fields("Operator").Value
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop
        rs.Close()
        SQLstring = "Select SubCompany from SubRecord where isnull(BackDate,'')<>'' group by SubCompany Order by SubCompany"
        rs.Open(SQLstring, cn, 1, 1)
        ComboBox1.Items.Clear()
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If

        rs.MoveFirst()
        Do While rs.EOF = False
            ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub

    Private Sub Form26_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = 2 Then
            DataGridView1.Width = Me.Width - 100
            DataGridView1.Height = Me.Height - 200
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        '按生产单查询
        If e.KeyCode = 13 Then

            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            DataGridView1.Rows.Clear()
            cn.Open(CNsfmdb)
            If (Label1.Text = "已外发未收回清单：") Then
                SQLstring = "Select * from SubRecord where isnull(BackDate,'')='' and SFOrder=" & "'" & TextBox1.Text & "'"
            Else
                SQLstring = "Select * from SubRecord where isnull(BackDate,'')<>'' and SFOrder=" & "'" & TextBox1.Text & "'"
            End If

            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            End If
            DataGridView1.Rows.Add(rs.RecordCount)
            StatusStrip1.Items(0).Text = "记录条数:  " & rs.RecordCount
            rs.MoveFirst()
            Dim DataRow As Integer
            DataRow = 0
            Do While rs.EOF = False
                DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("OKQty", DataRow).Value = rs.Fields("OKQty").Value
                DataGridView1.Item("SubCompany", DataRow).Value = rs.Fields("SubCompany").Value
                DataGridView1.Item("UnitPrice", DataRow).Value = rs.Fields("UnitPrice").Value
                DataGridView1.Item("SubDate", DataRow).Value = rs.Fields("SubDate").Value
                DataGridView1.Item("BackDate", DataRow).Value = rs.Fields("BackDate").Value
                DataGridView1.Item("Comment", DataRow).Value = rs.Fields("Comment").Value
                DataGridView1.Item("Operator", DataRow).Value = rs.Fields("Operator").Value
                DataRow = DataRow + 1
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        '按图号查询
        If e.KeyCode = 13 Then

            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            DataGridView1.Rows.Clear()
            cn.Open(CNsfmdb)
            If (Label1.Text = "已外发未收回清单：") Then
                SQLstring = "Select * from SubRecord where isnull(BackDate,'')='' and DrawNo like " & "'%" & TextBox2.Text & "%'"
            Else
                SQLstring = "Select * from SubRecord where isnull(BackDate,'')<>'' and DrawNo like " & "'%" & TextBox2.Text & "%'"
            End If

            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            End If
            DataGridView1.Rows.Add(rs.RecordCount)
            StatusStrip1.Items(0).Text = "记录条数:  " & rs.RecordCount
            rs.MoveFirst()
            Dim DataRow As Integer
            DataRow = 0
            Do While rs.EOF = False
                DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("OKQty", DataRow).Value = rs.Fields("OKQty").Value
                DataGridView1.Item("SubCompany", DataRow).Value = rs.Fields("SubCompany").Value
                DataGridView1.Item("UnitPrice", DataRow).Value = rs.Fields("UnitPrice").Value
                DataGridView1.Item("SubDate", DataRow).Value = rs.Fields("SubDate").Value
                DataGridView1.Item("BackDate", DataRow).Value = rs.Fields("BackDate").Value
                DataGridView1.Item("Comment", DataRow).Value = rs.Fields("Comment").Value
                DataGridView1.Item("Operator", DataRow).Value = rs.Fields("Operator").Value
                DataRow = DataRow + 1
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        '按供应商查询
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        DataGridView1.Rows.Clear()
        cn.Open(CNsfmdb)
        If (Label1.Text = "已外发未收回清单：") Then
            SQLstring = "Select * from SubRecord where isnull(BackDate,'')='' and SubCompany=" & "'" & ComboBox1.SelectedItem.ToString & "'"
        Else
            SQLstring = "Select * from SubRecord where isnull(BackDate,'')<>'' and SubCompany=" & "'" & ComboBox1.SelectedItem.ToString & "'"
        End If

        rs.Open(SQLstring, cn, 1, 1)
        If (rs.RecordCount = 0) Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If
        DataGridView1.Rows.Add(rs.RecordCount)
        StatusStrip1.Items(0).Text = "记录条数:  " & rs.RecordCount
        rs.MoveFirst()
        Dim DataRow As Integer
        DataRow = 0
        Do While rs.EOF = False
            DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
            DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
            DataGridView1.Item("OKQty", DataRow).Value = rs.Fields("OKQty").Value
            DataGridView1.Item("SubCompany", DataRow).Value = rs.Fields("SubCompany").Value
            DataGridView1.Item("UnitPrice", DataRow).Value = rs.Fields("UnitPrice").Value
            DataGridView1.Item("SubDate", DataRow).Value = rs.Fields("SubDate").Value
            DataGridView1.Item("BackDate", DataRow).Value = rs.Fields("BackDate").Value
            DataGridView1.Item("Comment", DataRow).Value = rs.Fields("Comment").Value
            DataGridView1.Item("Operator", DataRow).Value = rs.Fields("Operator").Value
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub
End Class