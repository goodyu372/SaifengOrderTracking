Public Class Form客户退货输入
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
        Dim rs, rsOutput As New ADODB.Recordset
        Dim SQLstring, SQLOutput As String
        Dim ID, TotalOutput As Integer

        cn.Open(CNsfmdb)
        SQLOutput = "Select * from OutputRecord order by ID"
        rsOutput.Open(SQLOutput, cn, 1, 3)
        If rsOutput.RecordCount = 0 Then
            ID = 0
        Else
            rsOutput.MoveLast()
            ID = rsOutput("ID").Value
        End If

        For i = 0 To DataGridView2.Rows.Count - 1
            rsOutput.AddNew()
            ID = ID + 1
            rsOutput.Fields("ID").Value = ID
            rsOutput.Fields("SFOrder").Value = DataGridView2.Item("SFOrder", i).Value
            rsOutput.Fields("DrawNo").Value = DataGridView2.Item("DrawNo", i).Value
            rsOutput.Fields("PartName").Value = DataGridView2.Item("PartName", i).Value
            rsOutput.Fields("CustOrder").Value = DataGridView2.Item("CustOrder", i).Value
            rsOutput.Fields("OutputQty").Value = DataGridView2.Item("Qty", i).Value
            rsOutput.Fields("UnitP").Value = DataGridView2.Item("UnitP", i).Value
            'rsOutput.Fields("TotalMoney").Value = DataGridView2.Item("", i).Value
            rsOutput.Fields("Remark").Value = DataGridView2.Item("Remark", i).Value
            rsOutput.Fields("CustCode").Value = DataGridView2.Item("CustCode", i).Value
            rsOutput.Fields("OutputDate").Value = TextBox4.Text
            SQLstring = "Select sum(OutputQty) as TotalOutput from OutputRecord where (SFOrder=" & "'" & DataGridView2.Item("SFOrder", i).Value & "'" & " and DrawNo=" & "'" & DataGridView2.Item("DrawNo", i).Value & "'" & ") group by SFOrder,DrawNo"
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                TotalOutput = DataGridView2.Item("Qty", i).Value
            Else
                rs.MoveFirst()
                TotalOutput = rs.Fields("TotalOutput").Value + DataGridView2.Item("Qty", i).Value
            End If
            rsOutput.Update()
            rs.Close()
            SQLstring = "Select * from SFOrderBase where (SFOrder=" & "'" & DataGridView2.Item("SFOrder", i).Value & "'" & " and DrawNo=" & "'" & DataGridView2.Item("DrawNo", i).Value & "'" & ")"
            rs.Open(SQLstring, cn, 1, 3)
            rs.MoveFirst()
            rs.Fields("Status").Value = "客户退货"
            rs.Update()
            rs.Close()
        Next
        rsOutput.Close()
        cn.Close()
        rsOutput = Nothing
        cn = Nothing
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        TextBox4.Text = ""
        MsgBox("出货输入完成！")
    End Sub

    Private Sub Form49_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form49_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("ID", "序号")
        DataGridView1.Columns.Add("CustCode", "客户代码")
        DataGridView1.Columns.Add("CustOrder", "客户订单号")
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("PartName", "零件名称")
        DataGridView1.Columns.Add("Qty", "订单数量")
        DataGridView1.Columns.Add("TotalOut", "已出货数量")
        DataGridView1.Columns.Add("UnitP", "单价")
        DataGridView1.Columns.Add("Status", "状态")

        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView2.Columns.Add("SFOrder", "赛峰单号")
        DataGridView2.Columns.Add("DrawNo", "图号")
        DataGridView2.Columns.Add("PartName", "零件名称")
        DataGridView2.Columns.Add("CustOrder", "客户订单号")
        DataGridView2.Columns.Add("Qty", "退货数量")
        DataGridView2.Columns.Add("UnitP", "单价")
        'DataGridView2.Columns.Add("TotalMoney", "金额")
        DataGridView2.Columns.Add("CustCode", "客户代码")
        DataGridView2.Columns.Add("Remark", "备注")
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            '按生产单号查询

            DataGridView1.Rows.Clear()
            'DataGridView1.Columns.Clear()

            If TextBox1.Text = "" Then
                MsgBox("生产单号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring, SQLOutput As String
            Dim ID As Integer
            'SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,Status from SFOrderBase where SFOrder=" & "'" & TextBox1.Text & "' and ID not in (Select ID from TotalOutput where TotalOut>=Qty)"
            SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,TotalOut from TotalOutput where SFOrder=" & "'" & TextBox1.Text & "' and TotalOut>0"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("不存在此生产单号的出货！")
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

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        DataGridView2.Rows.Add(1)
        On Error Resume Next
        For i = 0 To DataGridView2.Columns.Count - 1
            DataGridView2.Item(i, DataGridView2.Rows.Count - 1).Value = DataGridView1.Item(DataGridView2.Columns(i).Name, e.RowIndex).Value
            DataGridView2.Item("Qty", DataGridView2.Rows.Count - 1).Value = -DataGridView1.Item("TotalOut", e.RowIndex).Value
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
            Dim SQLstring, SQLOutput As String
            Dim ID As Integer
            'SQLstring = "Select CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,Status from SFOrderBase where DrawNo like " & "'%" & TextBox2.Text & "%' and Status<>'完成入库'"
            'SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,Status from SFOrderBase where DrawNo like " & "'" & TextBox2.Text & "' and ID not in (Select ID from TotalOutput where TotalOut>=Qty)"
            SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,TotalOut from TotalOutput where DrawNo like " & "'%" & TextBox2.Text & "%' and TotalOut>0"
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
            Dim SQLstring, SQLOutput As String
            Dim ID As Integer
            'SQLstring = "Select CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,Status from SFOrderBase where CustOrder like " & "'%" & TextBox3.Text & "%' and Status<>'完成入库'"
            'SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,Status from SFOrderBase where CustOrder like " & "'" & TextBox3.Text & "' and ID not in (Select ID from TotalOutput where TotalOut>=Qty)"
            SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,TotalOut from TotalOutput where CustOrder like " & "'%" & TextBox3.Text & "%' and TotalOut>0"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("不存在该客户订单号！")
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
                DataGridView2.Item("Qty", DataGridView2.Rows.Count - 1).Value = -DataGridView1.Item("TotalOut", i).Value
            Next
        Next
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView2.Rows.RemoveAt(DataGridView2.SelectedCells(0).RowIndex)
    End Sub
End Class