Public Class Form查询返工记录

    Private Sub Form43_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form43_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select * from ReworkRecord order by ID"
        rs.Open(SQLstring, cn, 1, 1)
        DataGridView1.Columns.Add("ID", "序号")
        DataGridView1.Columns.Add("SFOrder", "单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("ReworkQty", "返工数量")
        DataGridView1.Columns.Add("ReworkDate", "返工日期")
        DataGridView1.Columns.Add("RWFinishDate", "返工收回日期")
        DataGridView1.Columns.Add("ReworkBackQty", "返工收回数量")
        DataGridView1.Columns.Add("Operator", "记录者")
        DataGridView1.Columns.Add("ReworkProcess", "返工工序")
        DataGridView1.Rows.Clear()
        If rs.RecordCount <> 0 Then
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = rs.Fields(i).Value
                Next
                rs.MoveNext()
            Loop
        End If
    End Sub
End Class