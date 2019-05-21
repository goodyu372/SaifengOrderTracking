Public Class Form员工离职录入
    Public existingWorkers As Integer

    Private Sub Form81_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form81_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        SQLstring = "Select * from SFHR"
        rs.Open(SQLstring, cn, 1, 2)
        existingWorkers = rs.RecordCount
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
            'DataGridView1.Rows(DataRow).DefaultCellStyle.BackColor = Color.Pink
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim QueryText As String
        QueryText = "Select * from SFHR"
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        cn.Open(CNsfmdb)

        For i = 0 To DataGridView1.Rows.Count - 1
            If (DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow) Then
                QueryText = "Select * from SFHR where 工号=" & "'" & DataGridView1.Item("工号", i).Value & "'"
                rs.Open(QueryText, cn, 1, 2)
                rs.Fields("在职标识").Value = "已离职"
                'rs.Fields("更新日期").Value = Today
                rs.Update()
                rs.Close()
            End If
        Next
        cn.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("员工离职信息输入完毕！")
    End Sub
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Yellow Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Nothing
        Else
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Yellow
        End If
    End Sub
End Class