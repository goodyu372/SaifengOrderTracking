Public Class Form更改图号

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
                DataRow = DataRow + 1
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Form27_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form27_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("SFOrder", "SFOrder")
        DataGridView1.Columns.Add("DrawNo", "DrawNo")
        DataGridView1.Columns.Add("DWGInfo", "DWGInfo")
        DataGridView1.Columns.Add("CustDWG", "CustDWG")
        DataGridView1.Columns.Add("CustCode", "CustCode")
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
                DataRow = DataRow + 1
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'MsgBox(DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value)
        If DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value = "" Then
            MsgBox("未选择！")
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MsgBox("请输入更改后的图号！")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        Dim NewDWGInfo, NewDrawNo, NewCustDWG, DWGInfo As String
        DWGInfo = DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value
        NewDrawNo = TextBox3.Text
        NewCustDWG = DataGridView1.Item("CustCode", DataGridView1.SelectedCells(0).RowIndex).Value & NewDrawNo
        NewDWGInfo = NewCustDWG & DataGridView1.Item("SFOrder", DataGridView1.SelectedCells(0).RowIndex).Value
        cn.Open(CNsfmdb)

        SQLstring = "Select * from SFOrderBase where DWGInfo=" & "'" & DWGInfo & "'"
        rs.Open(SQLstring, cn, 1, 3)
        rs.MoveFirst()
        rs.Fields("DWGInfo").Value = NewDWGInfo
        rs.Fields("CustDWG").Value = NewCustDWG
        rs.Fields("DrawNo").Value = NewDrawNo
        rs.Update()
        rs.Close()

        SQLstring = "Select * from ProcCard where DWGInfo=" & "'" & DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value & "'"
        rs.Open(SQLstring, cn, 1, 3)
        If rs.RecordCount > 0 Then
            rs.MoveFirst()
            Do While rs.EOF = False
                rs.Fields("DWGInfo").Value = NewDWGInfo
                rs.Fields("CustDWG").Value = NewCustDWG
                rs.Fields("DrawNo").Value = NewDrawNo
                rs.Fields("ProcCode").Value = DataGridView1.Item("SFOrder", DataGridView1.SelectedCells(0).RowIndex).Value & NewDrawNo & Mid(rs.Fields("ProcCode").Value, Len(rs.Fields("ProcCode").Value) - 2, 3)
                rs.MoveNext()
            Loop
        End If
        rs.Close()

        SQLstring = "Select * from ProcRecord where DWGInfo=" & "'" & DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value & "'"
        rs.Open(SQLstring, cn, 1, 3)
        If rs.RecordCount > 0 Then
            rs.MoveFirst()
            Do While rs.EOF = False
                rs.Fields("DWGInfo").Value = NewDWGInfo
                rs.Fields("CustDWG").Value = NewCustDWG
                rs.Fields("ProcCode").Value = DataGridView1.Item("SFOrder", DataGridView1.SelectedCells(0).RowIndex).Value & NewDrawNo & Mid(rs.Fields("ProcCode").Value, Len(rs.Fields("ProcCode").Value) - 2, 3)
                rs.MoveNext()
            Loop
        End If
        rs.Close()

        rs = Nothing
        cn = Nothing
        MsgBox("图号更改完成！")
    End Sub
End Class