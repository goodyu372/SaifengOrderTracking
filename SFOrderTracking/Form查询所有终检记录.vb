Public Class Form查询所有终检记录

    Private Sub Form34_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form34_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        Dim cn As New ADODB.Connection
        Dim rs, rs1 As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * from QARecord "
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)

        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
        Next

        rs.MoveFirst()
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
            DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
            DataGridView1.Item("CustCode", DataGridView1.Rows.Count - 1).Value = rs.Fields("CustCode").Value
            DataGridView1.Item("Qty", DataGridView1.Rows.Count - 1).Value = rs.Fields("Qty").Value
            DataGridView1.Item("DDate", DataGridView1.Rows.Count - 1).Value = rs.Fields("DDate").Value
            DataGridView1.Item("PartName", DataGridView1.Rows.Count - 1).Value = rs.Fields("PartName").Value
            DataGridView1.Item("ID", DataGridView1.Rows.Count - 1).Value = rs.Fields("ID").Value
            DataGridView1.Item("Operator", DataGridView1.Rows.Count - 1).Value = rs.Fields("Operator").Value
            DataGridView1.Item("InspectDate", DataGridView1.Rows.Count - 1).Value = rs.Fields("InspectDate").Value
            DataGridView1.Item("Comment", DataGridView1.Rows.Count - 1).Value = rs.Fields("Comment").Value
            DataGridView1.Item("OK", DataGridView1.Rows.Count - 1).Value = rs.Fields("OK").Value
            DataGridView1.Item("NG", DataGridView1.Rows.Count - 1).Value = rs.Fields("NG").Value
            rs.MoveNext()
        Loop
        StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub

    Private Sub Form34_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 50
        DataGridView1.Height = Me.Height - 100
    End Sub
End Class