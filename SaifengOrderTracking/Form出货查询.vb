Public Class Form出货查询

    Private Sub Form50_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form50_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("ID", "序号")
        DataGridView1.Columns.Add("CustOrder", "客户订单号")
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("PartName", "品名")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("OutputQty", "数量")
        DataGridView1.Columns.Add("UnitP", "单价")
        DataGridView1.Columns.Add("TotalMoney", "金额")
        DataGridView1.Columns.Add("Remark", "备注")
        DataGridView1.Columns.Add("CustCode", "客户代码")
        DataGridView1.Columns.Add("OutputDate", "出货日期")
        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
        Dim cn As New ADODB.Connection
        Dim rs, rsOutput As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select CustCode from CustList Order by CustCode"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        Do While rs.EOF = False
            ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub

    Private Sub Form50_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 200
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            DataGridView1.Rows.Clear()
            If TextBox1.Text = "" Then
                MsgBox("生产单号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * from OutputRecord where SFOrder=" & "'" & TextBox1.Text & "'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("该生产单号无出货记录！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(rs(i).Name, DataGridView1.Rows.Count - 1).Value = rs(i).Value
                    DataGridView1.Item("TotalMoney", DataGridView1.Rows.Count - 1).Value = IIf(IsNumeric(rs("UnitP").Value), rs("UnitP").Value, 0) * rs("OutputQty").Value
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
            DataGridView1.Rows.Clear()
            If TextBox2.Text = "" Then
                MsgBox("图号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * from OutputRecord where DrawNo like " & "'%" & TextBox2.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("该图号无出货记录！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(rs(i).Name, DataGridView1.Rows.Count - 1).Value = rs(i).Value
                    DataGridView1.Item("TotalMoney", DataGridView1.Rows.Count - 1).Value = IIf(IsNumeric(rs("UnitP").Value), rs("UnitP").Value, 0) * rs("OutputQty").Value
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
        '按客户订单查询出货
        If e.KeyCode = 13 Then
            DataGridView1.Rows.Clear()
            If TextBox3.Text = "" Then
                MsgBox("图号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * from OutputRecord where CustOrder like " & "'%" & TextBox3.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("该客户订单号无出货记录！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(rs(i).Name, DataGridView1.Rows.Count - 1).Value = rs(i).Value
                    DataGridView1.Item("TotalMoney", DataGridView1.Rows.Count - 1).Value = IIf(IsNumeric(rs("UnitP").Value), rs("UnitP").Value, 0) * rs("OutputQty").Value
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        DataGridView1.Rows.Clear()
        Dim cn As New ADODB.Connection
        Dim rs, rsOutput As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * from OutputRecord where CustCode=" & "'" & ComboBox1.SelectedItem.ToString & "'"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            MsgBox("该客户无出货记录！")
            Exit Sub
        End If
        rs.MoveFirst()
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Item(rs(i).Name, DataGridView1.Rows.Count - 1).Value = rs(i).Value
                DataGridView1.Item("TotalMoney", DataGridView1.Rows.Count - 1).Value = IIf(IsNumeric(rs("UnitP").Value), rs("UnitP").Value, 0) * rs("OutputQty").Value
            Next
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        DataGridView1.Rows.Clear()
        TextBox4.Text = DatePart(DateInterval.Year, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Day, DateTimePicker1.Value)
        Dim cn As New ADODB.Connection
        Dim rs, rsOutput As New ADODB.Recordset
        Dim SQLstring As String
        'SQLstring = "Select * from OutputRecord where DatePart(Year,OutputDate)=" & "'" & DatePart(DateInterval.Year, DateTimePicker1.Value) & "'" & " and DatePart(Month,OutputDate)=" & "'" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "'" & " and DatePart(Day,OutputDate)=" & "'" & DatePart(DateInterval.Day, DateTimePicker1.Value) & "'"
        SQLstring = "Select * from OutputRecord where DatePart(Year,OutputDate)=" & "'" & DatePart(DateInterval.Year, DateTimePicker1.Value) & "'" & " and DatePart(Month,OutputDate)=" & "'" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "'" & " and DatePart(Day,OutputDate)=" & "'" & DatePart(DateInterval.Day, DateTimePicker1.Value) & "'"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            MsgBox("该日期无出货记录！")
            Exit Sub
        End If
        rs.MoveFirst()
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Item(rs(i).Name, DataGridView1.Rows.Count - 1).Value = rs(i).Value
                DataGridView1.Item("TotalMoney", DataGridView1.Rows.Count - 1).Value = IIf(IsNumeric(rs("UnitP").Value), rs("UnitP").Value, 0) * rs("OutputQty").Value
            Next
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub
End Class