Public Class Form嘉睦零件库存查询
    Private Sub Form50_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form50_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        cn.Open(CNsfmdb)
        'SQLstring = "Select 图号, iif(isnull(StockIn,'')='',0,StockIn) as JMIn, iif(isnull(StockOut,'')='',0,StockOut) as JMOut from JMStock"
        SQLstring = "Select * from JMStock"
        rs.Open(SQLstring, cn, 1, 1)
        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
        Next
        StatusStrip1.Items(0).Text = "记录条数： " & rs.RecordCount
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = rs(i).Value
            Next
            rs.MoveNext()
        Loop
    End Sub
    Private Sub Form50_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 100
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            If TextBox1.Text = "" Then
                MsgBox("请输入图号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            cn.Open(CNsfmdb)
            'SQLstring = "Select 图号, iif(isnull(StockIn,'')='',0,StockIn) as JMIn, iif(isnull(StockOut,'')='',0,StockOut) as JMOut from JMStock"
            SQLstring = "Select * from JMStock where 图号 like " & "'%" & TextBox1.Text & "%'"
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有该图号的库存!")
                Exit Sub
            End If
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
            Next
            StatusStrip1.Items(0).Text = "记录条数： " & rs.RecordCount
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            End If
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = rs(i).Value
                Next
                rs.MoveNext()
            Loop
        End If
    End Sub
End Class