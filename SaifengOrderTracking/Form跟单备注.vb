Public Class Form跟单备注

    Private Sub Form85_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form85_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("ID", "序号")
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("PartName", "零件名称")
        DataGridView1.Columns.Add("Qty", "数量")
        DataGridView1.Columns.Add("Creator", "下单者")
        DataGridView1.Columns.Add("CDate", "下单日期")
        DataGridView1.Columns.Add("DDate", "交货日期")
        DataGridView1.Columns.Add("CustOrder", "客户订单号")
        DataGridView1.Columns.Add("CustCode", "客户代码")
        DataGridView1.Columns.Add("Status", "状态")
        DataGridView1.Columns.Add("FollowNote", "跟单备注")
        For i = 0 To DataGridView1.Columns.Count - 1
            If (DataGridView1.Columns(i).Name = "FollowNote") Then
                DataGridView1.Columns(i).ReadOnly = False
            Else
                DataGridView1.Columns(i).ReadOnly = True
            End If
        Next
    End Sub
    Private Sub Form85_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 200
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            If TextBox1.Text = "" Then
                MsgBox("请输入客户订单号！")
                Exit Sub
            End If

            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SQLstring = "Select * from SFOrderBase where CustOrder like " & "'%" & TextBox1.Text & "%'"
            rs.Open(SQLstring, cn, 1, 1)
            StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
            DataGridView1.Rows.Clear()
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有该客户订单号！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To DataGridView1.Columns.Count - 1
                    DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = rs(DataGridView1.Columns(i).Name).Value
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            If TextBox2.Text = "" Then
                MsgBox("请输入赛峰单号！")
                Exit Sub
            End If

            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SQLstring = "Select * from SFOrderBase where SFOrder=" & "'" & TextBox2.Text & "'"
            rs.Open(SQLstring, cn, 1, 1)
            StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
            DataGridView1.Rows.Clear()
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有此赛峰单号！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To DataGridView1.Columns.Count - 1
                    DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = rs(DataGridView1.Columns(i).Name).Value
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.Rows.Count < 1 Then
            MsgBox("没有要备注的记录！")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)

        For i = 0 To DataGridView1.Rows.Count - 1
            SQLstring = "Select * from SFOrderBase where ID=" & DataGridView1.Item("ID", i).Value
            rs.Open(SQLstring, cn, 1, 3)
            rs("FollowNote").Value = DataGridView1.Item("FollowNote", i).Value
            rs.Update()
            rs.Close()
        Next
        cn.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("跟单备注完成！")
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = 13 Then
            If TextBox3.Text = "" Then
                MsgBox("请输入图号！")
                Exit Sub
            End If

            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SQLstring = "Select * from SFOrderBase where DrawNo like " & "'%" & TextBox3.Text & "%'"
            rs.Open(SQLstring, cn, 1, 1)
            StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
            DataGridView1.Rows.Clear()
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有该图号！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To DataGridView1.Columns.Count - 1
                    DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = rs(DataGridView1.Columns(i).Name).Value
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
End Class