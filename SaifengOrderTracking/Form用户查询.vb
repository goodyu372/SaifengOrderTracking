Public Class Form用户查询

    Private Sub Form20_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form20_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Shell(SFOrderBaseNetConnect, vbHide)
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        Dim QueryText As String
        QueryText = "Select UserName,UserType,ChineseName from UserInfo"

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        cn.Open(CNsfmdb)
        rs.Open(QueryText, cn, 1, 1)
        Dim DataRow As Integer
        DataRow = 0
        rs.MoveFirst()
        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
        Next
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
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
        Shell(SFOrderBaseNetDisconnect, vbHide)
    End Sub
End Class