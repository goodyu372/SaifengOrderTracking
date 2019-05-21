Public Class Form客户信息查询及维护
    Public OldCustomer As Integer
    Private Sub Form22_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form22_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'shell(SFOrderBaseNetConnect, vbHide)
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        SQLstring = "Select * from CustList"
        rs.Open(SQLstring, cn, 1, 2)
        OldCustomer = rs.RecordCount
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
        'shell(SFOrderBaseNetDisconnect, vbHide)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim QueryText As String
        QueryText = "Select * from CustList"
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        cn.Open(CNsfmdb)
        rs.Open(QueryText, cn, 1, 2)
        For i = OldCustomer To DataGridView1.Rows.Count - 2
            rs.AddNew()
            rs.Fields("CustCode").Value = DataGridView1.Item("CustCode", i).Value
            rs.Fields("CustName").Value = DataGridView1.Item("CustName", i).Value
            rs.Fields("Contact").Value = DataGridView1.Item("Contact", i).Value
            rs.Fields("联系方式1").Value = DataGridView1.Item("联系方式1", i).Value
            rs.Fields("联系方式2").Value = DataGridView1.Item("联系方式2", i).Value
            rs.Fields("赛峰业务员").Value = DataGridView1.Item("赛峰业务员", i).Value
            rs.Update()
        Next
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("增加客户完毕！")
        Button1.Visible = False
    End Sub
End Class