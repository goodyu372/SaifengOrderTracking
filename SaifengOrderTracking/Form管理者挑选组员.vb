Public Class Form管理者挑选组员
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView2.Rows.Count <= 0 Then
            MsgBox("没有选好的组员！")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        For i = 0 To DataGridView2.Rows.Count - 1
            SQLstring = "Select 工号,姓名,在职标识 from SFHR where 工号=" & "'" & DataGridView2.Item("工号", i).Value & "'"
            rs.Open(SQLstring, cn, 1, 3)
            rs.Fields("在职标识").Value = ProcessName
            rs.Update()
            rs.Close()
        Next
        cn.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("组员选择完成!")
    End Sub
    Private Sub Form77_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form77_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView1.Columns.Add("工号", "工号")
        DataGridView1.Columns.Add("姓名", "姓名")
        DataGridView1.Columns.Add("在职标识", "在职标识")
        DataGridView2.Columns.Add("工号", "工号")
        DataGridView2.Columns.Add("姓名", "姓名")
        DataGridView2.Columns.Add("在职标识", "在职标识")
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select 工号,姓名,在职标识 from SFHR where 在职标识<>" & "'" & ProcessName & "'" & " and 在职标识<>'已离职' order by 姓名"
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            DataGridView1.Item("工号", DataGridView1.Rows.Count - 1).Value = rs("工号").Value
            DataGridView1.Item("姓名", DataGridView1.Rows.Count - 1).Value = rs("姓名").Value
            DataGridView1.Item("在职标识", DataGridView1.Rows.Count - 1).Value = rs("在职标识").Value
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).ReadOnly = True
            DataGridView2.Columns(i).ReadOnly = True
        Next
    End Sub
    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        DataGridView2.Rows.Add(1)
        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView2.Item(i, DataGridView2.Rows.Count - 1).Value = DataGridView1.Item(i, e.RowIndex).Value
        Next
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class