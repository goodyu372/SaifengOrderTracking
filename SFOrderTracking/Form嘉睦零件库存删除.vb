Public Class Form嘉睦零件库存删除
    Private Sub Form57_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Select Case JamuStockForm
            Case "Form56"
                Form嘉睦订单特别显示.Show()
            Case "Form58"
                Form嘉睦零件库存管理.Show()
        End Select
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        If e.KeyCode = 13 Then
            If TextBox1.Text = "" Then
                MsgBox("图号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            cn.Open(CNsfmdb)
            SQLstring = "Select * from JamuStock where 图号 like " & "'%" & TextBox1.Text & "%'"
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此图号的库存不存在！")
                Exit Sub
            End If
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
            Next
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = rs(i).Value
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
        Dim ID, num, 图号, 编号 As Object
        ID = DataGridView1.Item("ID", e.RowIndex).Value
        num = DataGridView1.Item("数量", e.RowIndex).Value
        图号 = DataGridView1.Item("图号", e.RowIndex).Value
        编号 = DataGridView1.Item("编号", e.RowIndex).Value
        If MsgBox($"图号{图号} 编号{编号} 数量{num} 确定要删除吗", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            'SQLstring = "Delete * from JamuStock where 编号=" & "'" & DataGridView1.Item("编号", e.RowIndex).Value & "'"
            SQLstring = $"Delete JamuStock where ID={ID}"
            rs.Open(SQLstring, cn, 1, 3)
            cn.Close()
            rs = Nothing
            cn = Nothing
            DataGridView1.Rows.RemoveAt(e.RowIndex)
            MsgBox($"图号{图号} ID{ID} 已成功删除！")
        End If
    End Sub

    Private Sub Form57_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim ID, num, 图号 As Object
        ID = DataGridView1.Item("ID", e.RowIndex).Value
        num = DataGridView1.Item("数量", e.RowIndex).Value
        图号 = DataGridView1.Item("图号", e.RowIndex).Value
        If e.ColumnIndex = 1 And e.RowIndex >= 0 And ID IsNot Nothing And num IsNot Nothing Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SQLstring = $"update JamuStock set 数量 = {num} where ID={ID}"
            rs.Open(SQLstring, cn, 1, 3)
            cn.Close()
            rs = Nothing
            cn = Nothing
            MsgBox($"{图号}已成功更改数量！")
        End If
    End Sub
End Class