Public Class Form嘉睦零件库存管理
    Public ss As New Microsoft.Office.Interop.Excel.Application
    Public xlsheet As Microsoft.Office.Interop.Excel.Worksheet
    Public xlbook As Microsoft.Office.Interop.Excel.Workbook
    Public SelectDelete As String
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        cn.Open(CNsfmdb)
        SQLstring = "Select * from JamuStock order by 图号,编号"
        rs.Open(SQLstring, cn, 1, 1)
        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
        Next
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            MsgBox("没有库存记录！")
            Exit Sub
        End If

        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Item(i, DataGridView1.Rows.Count - 2).Value = rs(i).Value
            Next
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("图号", "图号")
        DataGridView1.Columns.Add("数量", "数量")
        DataGridView1.Columns.Add("编号", "编号")

    End Sub
    Private Sub Form58_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '零件入库确认
        'Check PartNumber in JamuStock
        If DataGridView1.Rows.Count < 2 Then
            MsgBox("没有入库记录！")
            Exit Sub
        End If
        For i = 0 To DataGridView1.Rows.Count - 2
            If (IsNumeric(DataGridView1.Item("数量", DataGridView1.Rows.Count - 2).Value)) Then
            Else
                MsgBox("请检查入库数量！")
                Exit Sub
            End If
        Next
        For i = 0 To DataGridView1.Rows.Count - 2
            If (IsDBNull(DataGridView1.Item("图号", DataGridView1.Rows.Count - 2).Value)) Then
                MsgBox("请检查图号！")
                Exit Sub
            Else
                If (DataGridView1.Item("图号", DataGridView1.Rows.Count - 2).Value = "") Then
                    MsgBox("图号不能为空！")
                    Exit Sub
                End If
            End If
        Next
        For i = 0 To DataGridView1.Rows.Count - 2
            If (IsDBNull(DataGridView1.Item("编号", DataGridView1.Rows.Count - 2).Value)) Then
                MsgBox("请检查编号！")
                Exit Sub
            Else
                If (DataGridView1.Item("编号", DataGridView1.Rows.Count - 2).Value = "") Then
                    MsgBox("编号不能为空！")
                    Exit Sub
                End If
            End If
        Next

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        Dim ID As Integer
        cn.Open(CNsfmdb)

        SQLstring = "Select * From JamuStock"
        rs.Open(SQLstring, cn, 1, 3)
        For i = 0 To DataGridView1.Rows.Count - 2
            rs.AddNew()
            rs.Fields("图号").Value = DataGridView1.Item("图号", i).Value
            rs.Fields("数量").Value = DataGridView1.Item("数量", i).Value
            rs.Fields("编号").Value = DataGridView1.Item("编号", i).Value
            'rs.Fields("入库日期").Value = Now
            rs.Update()
        Next
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        DataGridView1.Rows.Clear()
        MsgBox("嘉睦零件入库录入完成！")
    End Sub
    Private Sub Form58_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SelectDelete = ""
    End Sub
    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        SelectDelete = "Delete"
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        JamuStockForm = "Form58"
        Me.Hide()
        Form嘉睦零件库存删除.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class