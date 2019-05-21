Public Class Form员工信息维护
    Public existingWorkers As Integer
    Private Sub Form23_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form23_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'shell(SFOrderBaseNetConnect, vbHide)
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
        QueryText = "Select * from SFHR"
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        cn.Open(CNsfmdb)
        rs.Open(QueryText, cn, 1, 2)
        For i = existingWorkers To DataGridView1.Rows.Count - 1
            If IsDBNull(DataGridView1.Item(0, i).Value) Then
            Else
                If (DataGridView1.Item(0, i).Value = "") Then
                    Exit For
                End If
                rs.AddNew()
                rs.Fields("工号").Value = DataGridView1.Item("工号", i).Value
                rs.Fields("姓名").Value = DataGridView1.Item("姓名", i).Value
                rs.Fields("性别").Value = DataGridView1.Item("性别", i).Value
                rs.Fields("部门名称").Value = DataGridView1.Item("部门名称", i).Value
                rs.Fields("职位").Value = DataGridView1.Item("职位", i).Value
                rs.Fields("在职标识").Value = DataGridView1.Item("职位", i).Value
                rs.Fields("工种名称").Value = DataGridView1.Item("工种名称", i).Value
                rs.Update()
            End If
        Next
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("员工信息输入完毕！")
        Button1.Visible = False
    End Sub
End Class