Public Class Form嘉睦零件交期信息查询

    Private Sub Form45_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form45_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
            SQLstring = "Select * From JMOrderStock where (未交货数>0) and 图号 like " & "'%" & TextBox1.Text & "%' order by 交期"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            StatusStrip1.Items(0).Text = "" & rs.RecordCount
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
                If ((rs(i).Name = "单价") Or (rs(i).Name = "金额") Or (rs(i).Name = "图号单号")) Then
                    DataGridView1.Columns(i).Visible = False
                End If
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

    Private Sub Form45_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 150
    End Sub
End Class