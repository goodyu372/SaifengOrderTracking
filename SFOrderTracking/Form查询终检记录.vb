Public Class Form查询终检记录

    Private Sub Form33_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form33_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * from QARecord order by ID"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
        Next
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub

    Private Sub Form33_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 100
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            DataGridView1.Rows.Clear()
            Dim cn As New ADODB.Connection
            Dim rs, rs1 As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select Top 200 * from QARecord where DrawNo like " & "'%" & TextBox2.Text & "%' Order by ID Desc"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有此图号的记录，请检查！")
                Exit Sub
            End If
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
            StatusStrip1.Items(0).Text = "记录条数： " & rs.RecordCount
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            DataGridView1.Rows.Clear()
            Dim cn As New ADODB.Connection
            Dim rs, rs1 As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select Top 200 * from QARecord where SFOrder=" & "'" & TextBox1.Text & "' Order by ID Desc"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有此单号的记录，请检查！")
                Exit Sub
            End If
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
            StatusStrip1.Items(0).Text = "记录条数： " & rs.RecordCount
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = 13 Then
            'Query by CustCode
            DataGridView1.Rows.Clear()
            Dim cn As New ADODB.Connection
            Dim rs, rs1 As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select Top 200 * from QARecord where CustCode=" & "'" & TextBox3.Text & "' Order by ID Desc"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有此客户的记录，请检查！")
                Exit Sub
            End If
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
            StatusStrip1.Items(0).Text = "记录条数： " & rs.RecordCount
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Query by the newest 200 items
        DataGridView1.Rows.Clear()
        Dim cn As New ADODB.Connection
        Dim rs, rs1 As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select Top 200 * from QARecord Order by ID Desc"
        cn.Open(CNsfmdb)
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)
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
        StatusStrip1.Items(0).Text = "记录条数： " & rs.RecordCount
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub
End Class