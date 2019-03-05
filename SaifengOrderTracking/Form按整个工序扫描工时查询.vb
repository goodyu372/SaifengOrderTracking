Public Class Form按整个工序扫描工时查询
    Private Sub Form44_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form44_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("工序", "工序")
        DataGridView1.Columns.Add("人数", "人数")
        DataGridView1.Columns.Add("时期", "时期")
        DataGridView1.Columns.Add("工时", "工时")
    End Sub

    Private Sub Form44_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 200
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            If DataGridView1.Rows.Count = 0 Then
                MsgBox("请先选择工序！")
                Exit Sub
            End If
            If TextBox1.Text = "" Then
                MsgBox("日期不能为空！")
                Exit Sub
            End If
            Label4.Text = "日期 " & TextBox1.Text & "数据"
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            For i = 0 To DataGridView1.Rows.Count - 1
                SQLstring = "Select sum(Hours) as WorkTime from DailyWorkTime where DateMark=" & "'" & TextBox1.Text & "' and Operator=" & "'" & DataGridView1.Item(1, i).Value & "'"
                rs.Open(SQLstring, cn, 1, 1)
                DataGridView1.Item("工时", i).Value = Format(IIf(IsNumeric(rs(0).Value), rs(0).Value, 0), "0.0")
                DataGridView1.Item("时期", i).Value = TextBox1.Text
                rs.Close()
            Next
            cn.Close()
            rs = Nothing
            cn = Nothing
            Label4.Text = "日期： " & TextBox1.Text & " 数据"
            Label4.Visible = True
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            If DataGridView1.Rows.Count = 0 Then
                MsgBox("请先选择工序！")
                Exit Sub
            End If
            If TextBox2.Text = "" Then
                MsgBox("月份不能为空！")
                Exit Sub
            End If
            Label4.Text = "月份 " & TextBox1.Text & "数据"
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            For i = 0 To DataGridView1.Rows.Count - 1
                SQLstring = "Select sum(Hours) as WorkTime from MonthlyWorkTime where MonthMark=" & "'" & TextBox2.Text & "' and Operator=" & "'" & DataGridView1.Item(1, i).Value & "'"
                rs.Open(SQLstring, cn, 1, 1)
                DataGridView1.Item("工时", i).Value = Format(IIf(IsNumeric(rs(0).Value), rs(0).Value, 0), "0.0")
                DataGridView1.Item("时期", i).Value = TextBox2.Text
                rs.Close()
            Next
            cn.Close()
            rs = Nothing
            cn = Nothing
            Label4.Text = "月份： " & TextBox2.Text & " 数据"
            Label4.Visible = True
        End If
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If InStr(Label4.Text, "日期") > 0 Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            Dim j As Integer
            cn.Open(CNsfmdb)
            SQLstring = "Select * from ProcRecord"
            rs.Open(SQLstring, cn, 1, 1)
            Form个人加工详单.DataGridView1.Rows.Clear()
            Form个人加工详单.DataGridView1.Columns.Clear()
            For i = 0 To rs.Fields.Count - 1
                Form个人加工详单.DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
            Next
            rs.Close()

            SQLstring = "Select * from ProcRecord where (10000*year(RecordTime)+100*month(RecordTime)+day(RecordTime))=" & Val(DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value) & " and Operator=" & "'" & DataGridView1.Item("姓名", DataGridView1.SelectedCells(0).RowIndex).Value & "'"
            Form个人加工详单.Label1.Text = DataGridView1.Item("姓名", DataGridView1.SelectedCells(0).RowIndex).Value & "    " & DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value & "  记录："
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount > 0 Then
                rs.MoveFirst()
                Do While rs.EOF = False
                    Form个人加工详单.DataGridView1.Rows.Add(1)
                    For j = 0 To rs.Fields.Count - 1
                        Form个人加工详单.DataGridView1.Item(j, Form个人加工详单.DataGridView1.Rows.Count - 1).Value = rs(j).Value
                    Next
                    rs.MoveNext()
                Loop
            End If
            rs.Close()

            cn.Close()
            rs = Nothing
            cn = Nothing
            Me.Hide()
            Form个人加工详单.WindowState = FormWindowState.Maximized
            Form个人加工详单.Show()
        ElseIf (InStr(Label4.Text, "月份") > 0) Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            Dim j As Integer
            cn.Open(CNsfmdb)
            SQLstring = "Select * from ProcRecord"
            rs.Open(SQLstring, cn, 1, 1)
            Form个人加工详单.DataGridView1.Rows.Clear()
            Form个人加工详单.DataGridView1.Columns.Clear()
            For i = 0 To rs.Fields.Count - 1
                Form个人加工详单.DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
            Next
            rs.Close()

            SQLstring = "Select * from ProcRecord where (100*year(RecordTime)+month(RecordTime))=" & Val(DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value) & " and Operator=" & "'" & DataGridView1.Item("姓名", DataGridView1.SelectedCells(0).RowIndex).Value & "'"
            Form个人加工详单.Label1.Text = DataGridView1.Item("姓名", DataGridView1.SelectedCells(0).RowIndex).Value & "    " & DataGridView1.Item(2, DataGridView1.SelectedCells(0).RowIndex).Value & "  记录："
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount > 0 Then
                rs.MoveFirst()
                Do While rs.EOF = False
                    Form个人加工详单.DataGridView1.Rows.Add(1)
                    For j = 0 To rs.Fields.Count - 1
                        Form个人加工详单.DataGridView1.Item(j, Form个人加工详单.DataGridView1.Rows.Count - 1).Value = rs(j).Value
                    Next
                    rs.MoveNext()
                Loop
            End If
            rs.Close()

            cn.Close()
            rs = Nothing
            cn = Nothing
            Me.Hide()
            Form个人加工详单.WindowState = FormWindowState.Maximized
            Form个人加工详单.Show()
        End If
    End Sub
End Class