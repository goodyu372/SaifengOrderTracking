Public Class Form增加供应商

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            If (TextBox1.Text = "pmc731023") Then
                'shell(SFOrderBaseNetConnect, vbHide)
                Dim cn As New ADODB.Connection
                Dim rs As New ADODB.Recordset
                Dim SQLstring As String
                cn.Open(CNsfmdb)
                DataGridView1.Rows.Clear()
                DataGridView1.Columns.Clear()
                SQLstring = "Select * from SubList"
                rs.Open(SQLstring, cn, 1, 3)
                rs.MoveFirst()
                For i = 1 To rs.Fields.Count
                    DataGridView1.Columns.Add(rs.Fields(i - 1).Name, rs.Fields(i - 1).Name)
                Next
                DataGridView1.Rows.Add(rs.RecordCount)
                Dim DataRow As Integer
                DataRow = 0
                Do While rs.EOF = False
                    For i = 1 To rs.Fields.Count
                        DataGridView1.Item(i - 1, DataRow).Value = rs.Fields(i - 1).Value
                    Next i
                    DataGridView1.Rows(DataRow).DefaultCellStyle.BackColor = Color.Pink
                    DataRow = DataRow + 1
                    rs.MoveNext()
                Loop
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                DataGridView1.Visible = True
                Button1.Visible = True
            Else
                MsgBox("密码错误！")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim QueryText As String
        QueryText = "Select * from SubList"
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        cn.Open(CNsfmdb)
        rs.Open(QueryText, cn, 1, 3)
        For i = 0 To DataGridView1.Rows.Count - 1
            If IsDBNull(DataGridView1.Item(0, i).Value) Then
            Else
                If (DataGridView1.Item(0, i).Value = "") Then
                    Exit For
                Else
                    If DataGridView1.Rows(i).DefaultCellStyle.BackColor <> Color.Pink Then
                        rs.AddNew()
                        rs.Fields("供应商名称").Value = DataGridView1.Item("供应商名称", i).Value
                        rs.Fields("联系电话1").Value = DataGridView1.Item("联系电话1", i).Value
                        rs.Fields("联系电话2").Value = DataGridView1.Item("联系电话2", i).Value
                        rs.Fields("加工项目").Value = DataGridView1.Item("加工项目", i).Value
                        rs.Update()
                        DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Pink
                    End If
                End If
            End If
        Next
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("新增供应商信息输入完毕！")
        Button1.Visible = False

    End Sub

    Private Sub Form30_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form30_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class