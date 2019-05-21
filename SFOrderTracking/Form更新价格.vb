Public Class Form更新价格

    Private Sub Form39_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form39_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        cn.Open(CNsfmdb)
        SQLstring = "Select * from TempPrice"
        rs.Open(SQLstring, cn, 1, 1)
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            MsgBox("没有要更新的价格！")
            Exit Sub
        End If

        rs.MoveFirst()
        Dim DataRow As Integer
        DataRow = 0
        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
        Next
        DataGridView1.Rows.Add(rs.RecordCount)
        Do While rs.EOF = False
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Item(i, DataRow).Value = rs.Fields(i).Value
            Next
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error Resume Next
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("没有要更新的价格！")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        For i = 0 To DataGridView1.Rows.Count - 1
            SQLstring = "Select * from SFOrderBase where SFOrder=" & "'" & DataGridView1.Item("SFOrder", i).Value & "' and DrawNo=" & "'" & DataGridView1.Item("DrawNo", i).Value & "'"
            rs.Open(SQLstring, cn, 1, 3)
            rs.Fields("UnitP").Value = DataGridView1.Item("UnitP", i).Value
            rs.Fields("RMBUnitP").Value = DataGridView1.Item("RMBUnitP", i).Value
            rs.Update()
            rs.Close()
        Next

        cn.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("已更新价格！")
    End Sub
End Class