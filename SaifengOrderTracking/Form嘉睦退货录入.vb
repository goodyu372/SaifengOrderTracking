Public Class Form嘉睦退货录入
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '退货确认
        If TextBox4.Text = "" Then
            MsgBox("请选择退货日期！")
            Exit Sub
        End If
        For i = 0 To DataGridView2.Rows.Count - 1
            If DataGridView2.Item("Qty", i).Value > 0 Then
                MsgBox("退货数应为负数！")
                Exit Sub
            End If
        Next
        Dim cn As New ADODB.Connection
        Dim rsOutput As New ADODB.Recordset
        Dim SQLOutput As String
        Dim ID As Integer

        cn.Open(CNsfmdb)
        SQLOutput = "Select * from JamuOutput order by 序号"
        rsOutput.Open(SQLOutput, cn, 1, 3)
        If rsOutput.RecordCount = 0 Then
            ID = 0
        Else
            rsOutput.MoveLast()
            ID = rsOutput("序号").Value
        End If

        For i = 0 To DataGridView2.Rows.Count - 1
            rsOutput.AddNew()
            ID = ID + 1
            rsOutput.Fields("序号").Value = ID
            rsOutput.Fields("图号").Value = DataGridView2.Item("图号", i).Value
            rsOutput.Fields("数量").Value = DataGridView2.Item("数量", i).Value
            rsOutput.Fields("单价").Value = DataGridView2.Item("单价", i).Value
            rsOutput.Fields("总价").Value = DataGridView2.Item("总价", i).Value
            rsOutput.Fields("客户订单号").Value = DataGridView2.Item("客户订单号", i).Value
            rsOutput.Fields("备注").Value = DataGridView2.Item("备注", i).Value
            rsOutput.Fields("送货日期").Value = TextBox4.Text
            rsOutput.Update()
        Next
        rsOutput.Close()
        cn.Close()
        rsOutput = Nothing
        cn = Nothing
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        TextBox4.Text = ""
        MsgBox("嘉睦退货输入完成！")
    End Sub

    Private Sub Form49_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form49_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("序号", "序号")
        DataGridView1.Columns.Add("图号", "图号")
        DataGridView1.Columns.Add("数量", "数量")
        DataGridView1.Columns.Add("单价", "单价")
        DataGridView1.Columns.Add("总价", "总价")
        DataGridView1.Columns.Add("客户订单号", "客户订单号")
        DataGridView1.Columns.Add("备注", "备注")
        DataGridView1.Columns.Add("送货日期", "送货日期")

        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView2.Columns.Add("图号", "图号")
        DataGridView2.Columns.Add("数量", "退货数量")
        DataGridView2.Columns.Add("客户订单号", "客户订单号")
        DataGridView2.Columns.Add("单价", "单价")
        DataGridView2.Columns.Add("总价", "总价")
        DataGridView2.Columns.Add("备注", "备注")
    End Sub
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        DataGridView2.Rows.Add(1)
        On Error Resume Next
        For i = 0 To DataGridView2.Columns.Count - 1
            DataGridView2.Item(i, DataGridView2.Rows.Count - 1).Value = DataGridView1.Item(DataGridView2.Columns(i).Name, e.RowIndex).Value
            DataGridView2.Item("备注", DataGridView2.Rows.Count - 1).Value = "退货"
        Next
        DataGridView2.FirstDisplayedScrollingRowIndex = DataGridView2.Rows.Count - 1
    End Sub
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        TextBox4.Text = DatePart(DateInterval.Year, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Day, DateTimePicker1.Value)
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            '按图号查询
            DataGridView1.Rows.Clear()
            If TextBox2.Text = "" Then
                MsgBox("图号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select 序号,图号,数量,单价,总价,客户订单号,备注,送货日期 from JamuOutput where 图号 like " & "'%" & TextBox2.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("不存在该图号的出货！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(rs(i).Name, DataGridView1.Rows.Count - 1).Value = rs(i).Value
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing

        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = 13 Then
            '按客户订单号查询
            DataGridView1.Rows.Clear()
            If TextBox3.Text = "" Then
                MsgBox("客户订单号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select 序号,图号,数量,单价,总价,客户订单号,备注,送货日期 from JamuOutput where 客户订单号 like " & "'%" & TextBox3.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("不存在该客户订单号的出货记录！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(rs(i).Name, DataGridView1.Rows.Count - 1).Value = rs(i).Value
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing

        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '全部选择
        DataGridView2.Rows.Clear()
        If DataGridView1.Rows.Count = 0 Then
            Exit Sub
        End If
        For i = 0 To DataGridView1.Rows.Count - 1
            DataGridView2.Rows.Add(1)
            On Error Resume Next
            For j = 0 To DataGridView2.Columns.Count - 1
                DataGridView2.Item(j, DataGridView2.Rows.Count - 1).Value = DataGridView1.Item(DataGridView2.Columns(j).Name, i).Value
            Next
        Next
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView2.Rows.RemoveAt(DataGridView2.SelectedCells(0).RowIndex)
    End Sub
End Class