Public Class Form取消工艺

    Private Sub Form28_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form28_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("SFOrder", "SFOrder")
        DataGridView1.Columns.Add("DrawNo", "DrawNo")
        DataGridView1.Columns.Add("DWGInfo", "DWGInfo")
        DataGridView1.Columns.Add("CustDWG", "CustDWG")
        DataGridView1.Columns.Add("CustCode", "CustCode")
        DataGridView1.Columns.Add("Status", "Status")
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            DataGridView1.Rows.Clear()
            cn.Open(CNsfmdb)
            SQLstring = "Select * from SFOrderBase where SFOrder=" & "'" & TextBox1.Text & "'"
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("无此单号！")
                Exit Sub
            End If

            DataGridView1.Rows.Add(rs.RecordCount)
            rs.MoveFirst()
            Dim DataRow As Integer
            DataRow = 0
            Do While rs.EOF = False
                DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("DWGInfo", DataRow).Value = rs.Fields("DWGInfo").Value
                DataGridView1.Item("CustDWG", DataRow).Value = rs.Fields("CustDWG").Value
                DataGridView1.Item("CustCode", DataRow).Value = rs.Fields("CustCode").Value
                DataGridView1.Item("Status", DataRow).Value = rs.Fields("Status").Value
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
        If e.KeyCode = 13 Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            DataGridView1.Rows.Clear()
            cn.Open(CNsfmdb)
            SQLstring = "Select * from SFOrderBase where DrawNo like " & "'%" & TextBox2.Text & "%'"
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("无此图号！")
                Exit Sub
            End If
            DataGridView1.Rows.Add(rs.RecordCount)
            rs.MoveFirst()
            Dim DataRow As Integer
            DataRow = 0
            Do While rs.EOF = False
                DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("DWGInfo", DataRow).Value = rs.Fields("DWGInfo").Value
                DataGridView1.Item("CustDWG", DataRow).Value = rs.Fields("CustDWG").Value
                DataGridView1.Item("CustCode", DataRow).Value = rs.Fields("CustCode").Value
                DataGridView1.Item("Status", DataRow).Value = rs.Fields("Status").Value
                DataRow = DataRow + 1
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value = "" Then
            MsgBox("未选择！")
            Exit Sub
        End If

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        Dim DWGInfo As String
        DWGInfo = DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value

        cn.Open(CNsfmdb)
        SQLstring = "Select * from ProcRecord where DWGInfo=" & "'" & DWGInfo & "'"
        rs.Open(SQLstring, cn, 1, 3)
        If rs.RecordCount > 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            MsgBox("该图号已有扫描记录，不能更改！")
            Exit Sub
        End If
        rs.Close()

        SQLstring = "delete ProcCard where DWGInfo=" & "'" & DWGInfo & "'"
        rs.Open(SQLstring, cn, 1, 3)

        SQLstring = "delete ProcRecord where DWGInfo=" & "'" & DWGInfo & "'"
        rs.Open(SQLstring, cn, 1, 3)

        SQLstring = "Select * from SFOrderBase where DWGInfo=" & "'" & DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value & "'"
        rs.Open(SQLstring, cn, 1, 3)
        rs.MoveFirst()
        rs.Fields("Status").Value = ""
        rs.Update()
        rs.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("工艺取消完成完成！")
    End Sub
End Class