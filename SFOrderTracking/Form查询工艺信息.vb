Public Class Form查询工艺信息
    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = 13 Then
            DataGridView1.Rows.Clear()
            If TextBox8.Text = "" Then
                MsgBox("工艺信息不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SQLstring = "Select CustCode,DWGInfo,SFOrder,DrawNo,ProcDesc From ProcCard where ProcDesc like " & "'%" & TextBox8.Text & "%'"
            rs.Open(SQLstring, cn, 1, 1)
            StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("找不到相关信息！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                DataGridView1.Item("CustCode", DataGridView1.Rows.Count - 1).Value = rs("CustCode").Value
                DataGridView1.Item("DWGInfo", DataGridView1.Rows.Count - 1).Value = rs("DWGInfo").Value
                DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = rs("SFOrder").Value
                DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = rs("DrawNo").Value
                DataGridView1.Item("ProcDesc", DataGridView1.Rows.Count - 1).Value = rs("ProcDesc").Value
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub Form65_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form65_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Columns.Add("CustCode", "客户代码")
        DataGridView1.Columns.Add("DWGInfo", "图号信息")
        DataGridView1.Columns.Add("SFOrder", "单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("ProcDesc", "工艺信息")
        'DataGridView1.DefaultCellStyle.Font.Size = 
    End Sub
End Class