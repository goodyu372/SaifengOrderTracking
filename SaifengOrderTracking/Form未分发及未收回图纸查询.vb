Public Class Form未分发及未收回图纸查询

    Private Sub Form19_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form19_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("ReceiveDate", "收图日期")
        DataGridView1.Columns.Add("ReleaseDate", "给工艺日期")
        DataGridView1.Columns.Add("ReturnDate", "工艺交回日期")
        DataGridView1.Columns.Add("ProcMaker", "编工艺者")

        Dim cn As New ADODB.Connection
        Dim rs, rs1 As New ADODB.Recordset
        Dim SQLstring, CNString As String
        On Error Resume Next
        Dim CustCode As String
        Dim QueryText As String

        cn.Open(CNsfmdb)
        '*********************************************** First check PMC_DrawingRecord records ***************************************************************
        QueryText = "Select * from PMC_DrawingRecord where ((isnull(ReceiveDate,'')='') or (isnull(ReturnDate,'')='')  or (isnull(ReleaseDate,'')='') ) order by SForder,DrawNo"
        rs.Open(QueryText, cn, 1, 1)
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs1 = Nothing
            rs = Nothing
            cn = Nothing
            MsgBox("没有未处理完的图纸！")
        End If
        rs.MoveFirst()
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
            DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
            DataGridView1.Item("ReceiveDate", DataGridView1.Rows.Count - 1).Value = rs.Fields("ReceiveDate").Value
            DataGridView1.Item("ReleaseDate", DataGridView1.Rows.Count - 1).Value = rs.Fields("ReleaseDate").Value
            DataGridView1.Item("ReturnDate", DataGridView1.Rows.Count - 1).Value = rs.Fields("ReturnDate").Value
            DataGridView1.Item("ProcMaker", DataGridView1.Rows.Count - 1).Value = rs.Fields("ProcMaker").Value
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs1 = Nothing
        rs = Nothing
        cn = Nothing
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class