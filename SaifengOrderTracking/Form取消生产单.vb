Public Class Form取消生产单

    Private Sub Form13_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form13_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView1.Columns.Add("DrawNo", "DrawNo")
        DataGridView1.Columns.Add("DDate", "DDate")
        DataGridView1.Columns.Add("Qty", "Qty")
        DataGridView1.Columns.Add("OrderType", "OrderType")
        DataGridView1.Columns.Add("Status", "Status")

        DataGridView2.Columns.Add("DrawNo", "DrawNo")
        DataGridView2.Columns.Add("DDate", "DDate")
        DataGridView2.Columns.Add("Qty", "Qty")
        DataGridView2.Columns.Add("Status", "Status")
        DataGridView2.Columns.Add("Comment", "Comment")
        For i = 0 To DataGridView2.Columns.Count - 1
            DataGridView2.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = 13) Then
            'shell(SFOrderBaseNetConnect, vbHide)
            DataGridView1.Rows.Clear()
            DataGridView2.Rows.Clear()
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * From SFOrderBase where SFOrder=" & "'" & TextBox1.Text & "'"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                'shell(SFOrderBaseNetDisconnect, vbHide)
                MsgBox("生产单号不对，请检查！")
                Exit Sub
            End If
            DataGridView1.Rows.Add(rs.RecordCount)
            Dim RowNo As Integer
            RowNo = 0
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Item("DDate", RowNo).Value = rs.Fields("DDate").Value
                DataGridView1.Item("DrawNo", RowNo).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("Qty", RowNo).Value = rs.Fields("Qty").Value
                DataGridView1.Item("OrderType", RowNo).Value = rs.Fields("OrderType").Value
                DataGridView1.Item("Status", RowNo).Value = rs.Fields("Status").Value
                RowNo = RowNo + 1
                rs.MoveNext()
            Loop
            rs.Close()
            'shell(SFOrderBaseNetDisconnect, vbHide)
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For i = 0 To DataGridView1.Rows.Count - 1
            If (Len(DataGridView1.Item(0, i).Value) <> 0) Then
                DataGridView2.Rows.Add()
                DataGridView2.Item("DrawNo", i).Value = DataGridView1.Item("DrawNo", i).Value
                DataGridView2.Item("DDate", i).Value = DataGridView1.Item("DDate", i).Value
                DataGridView2.Item("Qty", i).Value = DataGridView1.Item("Qty", i).Value
                DataGridView2.Item("Status", i).Value = DataGridView1.Item("Status", i).Value
            End If
        Next
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        DataGridView2.Rows.Add()
        Dim i As Integer
        i = DataGridView2.Rows.Count - 1
        DataGridView2.Item("DrawNo", i).Value = DataGridView1.Item("DrawNo", e.RowIndex).Value
        DataGridView2.Item("DDate", i).Value = DataGridView1.Item("DDate", e.RowIndex).Value
        DataGridView2.Item("Qty", i).Value = DataGridView1.Item("Qty", e.RowIndex).Value
        DataGridView2.Item("Status", i).Value = DataGridView1.Item("Status", e.RowIndex).Value
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'shell(SFOrderBaseNetConnect, vbHide)
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        'Dim DWGInfo, SFOrder, DrawNo, CustCode, CustDWG, ProcMaker As String
        Dim DrawNo As String
        cn.Open(CNsfmdb)
        Dim RowNo As Integer
        RowNo = 0
        For i = 0 To DataGridView2.Rows.Count - 1
            If (Len(DataGridView2.Item("Comment", i).Value) = 0) Then
                cn.Close()
                cn = Nothing
                'shell(SFOrderBaseNetDisconnect, vbHide)
                MsgBox("请增加备注！")
                Exit Sub
            End If
            DrawNo = DataGridView2.Item("DrawNo", i).Value
            'SQLstring = "Select * From SFOrderBase where ((SFOrder=" & "'" & TextBox1.Text & "'" & ") and (DrawNo=" & "'" & DrawNo & "'" & "))"
            SQLstring = "Select * From SFOrderBase where ((SFOrder=" & "'" & TextBox1.Text & "'" & ") and (DrawNo=" & "'" & DrawNo & "'" & "))"
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 3)
            rs.MoveFirst()
            Do While rs.EOF = False
                rs.Fields("Status").Value = "取消生产单"
                rs.Fields("Comment").Value = DataGridView2.Item("Comment", RowNo).Value
                rs.Update()
                rs.MoveNext()
            Loop
            rs.Close()
            RowNo = RowNo + 1
        Next
        cn.Close()
        rs = Nothing
        cn = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        MsgBox("成功取消生产单！")
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub
End Class