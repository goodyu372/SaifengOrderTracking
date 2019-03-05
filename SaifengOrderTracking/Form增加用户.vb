Public Class Form增加用户

    Private Sub Form21_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form21_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Shell(SFOrderBaseNetConnect, vbHide)
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        Dim QueryText As String
        QueryText = "Select UserName,UserType,ChineseName from UserInfo"

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        cn.Open(CNsfmdb)
        'rs.Open(QueryText, cn, 1, 2)
        rs.Open(QueryText, cn, 1, 1)
        Dim DataRow As Integer
        DataRow = 0
        rs.MoveFirst()
        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
        Next
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            DataGridView1.Rows(DataRow).DefaultCellStyle.BackColor = Color.Yellow
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim QueryText As String
        QueryText = "Select * from UserInfo"
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        cn.Open(CNsfmdb)
        rs.Open(QueryText, cn, 1, 2)
        For i = 0 To DataGridView1.Rows.Count - 1
            If IsDBNull(DataGridView1.Item(0, i).Value) Then
            Else
                If DataGridView1.Rows(i).DefaultCellStyle.BackColor <> Color.Yellow Then
                    If (DataGridView1.Item(0, i).Value = "") Then
                        Exit For
                    End If
                    rs.AddNew()
                    rs.Fields("UserName").Value = DataGridView1.Item("UserName", i).Value
                    rs.Fields("Password").Value = "123456"
                    rs.Fields("Changetimes").Value = 0
                    rs.Fields("UserType").Value = DataGridView1.Item("UserType", i).Value
                    rs.Fields("ChineseName").Value = DataGridView1.Item("ChineseName", i).Value
                    rs.Update()
                End If
            End If
        Next
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("增加用户完毕！")
    End Sub
End Class