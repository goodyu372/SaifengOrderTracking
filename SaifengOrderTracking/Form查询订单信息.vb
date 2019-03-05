Public Class Form查询订单信息

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            If TextBox1.Text = "" Then
                MsgBox("请输入单号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String

            SQLstring = "Select * from SFOrderBase where SFOrder=" & "'" & TextBox1.Text & "'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此单号不存在！")
                Exit Sub
            End If
            rs.MoveFirst()
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
            Next
            Dim TempRow As Integer
            TempRow = 0
            Do While Not rs.EOF
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(i, TempRow).Value = rs.Fields(i).Value
                Next
                TempRow = TempRow + 1
                rs.MoveNext()
            Loop
            For i = 0 To DataGridView1.RowCount - 1
                If (DataGridView1.Item("DDate", i).Value < Now.Date) Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                End If
            Next
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            If TextBox2.Text = "" Then
                MsgBox("请输入图号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String

            SQLstring = "Select * from SFOrderBase where DrawNo like " & "'%" & TextBox2.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此图号不存在！")
                Exit Sub
            End If
            rs.MoveFirst()
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
            Next
            Dim TempRow As Integer
            TempRow = 0
            Do While Not rs.EOF
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(i, TempRow).Value = rs.Fields(i).Value
                Next
                TempRow = TempRow + 1
                rs.MoveNext()
            Loop
            For i = 0 To DataGridView1.RowCount - 1
                If (DataGridView1.Item("DDate", i).Value < Now.Date) Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                End If
            Next
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing

        End If
    End Sub

    Private Sub Form40_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form40_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = 13 Then
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            If TextBox3.Text = "" Then
                MsgBox("请输入客户代码！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String

            SQLstring = "Select Top 1000 * from SFOrderBase where CustCode=" & "'" & TextBox3.Text & "' Order by ID desc"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此客户代码不存在！")
                Exit Sub
            End If
            rs.MoveFirst()
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
            Next
            Dim TempRow As Integer
            TempRow = 0
            Do While Not rs.EOF
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(i, TempRow).Value = rs.Fields(i).Value
                Next
                TempRow = TempRow + 1
                rs.MoveNext()
            Loop
            For i = 0 To DataGridView1.RowCount - 1
                If (DataGridView1.Item("DDate", i).Value < Now.Date) Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                End If
            Next
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing

        End If
    End Sub
End Class