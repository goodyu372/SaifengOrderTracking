Public Class Form返工收回
    Public CellSelect As Integer
    Private Sub Form42_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form42_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select * from ReworkRecord where isnull(RWFinishDate,'')='' order by ID"
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
        CellSelect = 0
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        CellSelect = 1
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CellSelect = 0 Then
            MsgBox("请选择要收回的返工记录！")
            Exit Sub
        End If
        If (IsNumeric(DataGridView1.Item("ReworkBackQty", DataGridView1.SelectedCells(0).RowIndex).Value)) Then
        Else
            MsgBox("请输入收回的返工零件数量！")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select * from ReworkRecord where ID=" & "'" & DataGridView1.Item("ID", DataGridView1.SelectedCells(0).RowIndex).Value & "'"
        rs.Open(SQLstring, cn, 1, 3)
        rs.Fields("RWFinishDate").Value = Now
        rs.Fields("ReworkBackQty").Value = DataGridView1.Item("ReworkBackQty", DataGridView1.SelectedCells(0).RowIndex).Value
        rs.Update()
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        DataGridView1.Rows.RemoveAt(DataGridView1.SelectedCells(0).RowIndex)
        MsgBox("已完成收回所选择的返工记录！")
    End Sub
End Class