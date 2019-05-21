Public Class Form输入生产单
    Public ss As New Microsoft.Office.Interop.Excel.Application
    Public xlsheet As Microsoft.Office.Interop.Excel.Worksheet
    Public xlbook As Microsoft.Office.Interop.Excel.Workbook
    Public Currency As String
    Private Sub Form12_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form12_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'shell(SFOrderBaseNetConnect, vbHide)

        Currency = "RMB"
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        'Dim SQLstring As String
        Dim SQLstring, CNString As String
        'On Error Resume Next
        '********************************** Textbox1 connection ***************************************************************
        'List down the SFOrder
        SQLstring = "Select * From SFOrderBase order by SFOrder"
        cn.Open(CNsfmdb)
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveLast()
        'MsgBox(rs.RecordCount)
        TextBox1.Text = rs.Fields("SFOrder").Value
        rs.Close()
        '*************************** end of combobox1 connection *****************************************************************
        '********************************** combobox1connection *****************************************************************
        'List down the Customer code
        SQLstring = "Select CustCode From CustList order by CustCode"
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()

        'MsgBox(rs.RecordCount)
        Do While Not rs.EOF
            ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        rs.Close()
        '*************************** end of combobox1 connection *****************************************************************
        '********************************** combobox2connection *****************************************************************
        'List down OrderType
        SQLstring = "Select * From OrderType"
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        'MsgBox(rs.RecordCount)
        Do While Not rs.EOF
            ComboBox2.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        rs.Close()
        '*************************** end of combobox2 connection *****************************************************************
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        cn.Close()
        rs = Nothing
        cn = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox2.Text = ComboBox1.SelectedItem.ToString
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'shell(SFOrderBaseNetConnect, vbHide)
        If (TextBox2.Text = "") Then
            'shell(SFOrderBaseNetDisconnect, vbHide)
            MsgBox("请选择客户代码！")
            Exit Sub
        End If
        If (TextBox4.Text = "") Then
            'shell(SFOrderBaseNetDisconnect, vbHide)
            MsgBox("请选择生产单种类！")
            Exit Sub
        End If
        '**************************** 自动生成生产单号 **************fffffff**************************
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        'Dim SQLstring As String
        Dim SQLstring As String
        'On Error Resume Next
        SQLstring = "Select * From SFOrderBase order by SFOrder"

        cn.Open(CNsfmdb)
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 3)
        rs.MoveLast()
        'MsgBox(rs.RecordCount)
        TextBox1.Text = rs.Fields("SFOrder").Value
        TextBox3.Text = rs.Fields("SFOrder").Value + 1
        rs.Close()
        '***************************** end of 自动生成生产单号 *********************************
        For i = 2 To xlsheet.UsedRange.Rows.Count
            If (CStr(xlsheet.Cells(i, 1).value) = "") Then
                MsgBox("请检查所有零件图号！")
                'shell(SFOrderBaseNetDisconnect, vbHide)
                Exit Sub
            End If
        Next
        For i = 2 To xlsheet.UsedRange.Rows.Count
            If (IsNumeric(xlsheet.Cells(i, 4).value)) Then
                'shell(SFOrderBaseNetDisconnect, vbHide)
            Else
                MsgBox("数量格式有误，请检查！")
                Exit Sub
            End If
        Next
        For i = 2 To xlsheet.UsedRange.Rows.Count
            If (IsDate(xlsheet.Cells(i, 7).value)) Then
                'shell(SFOrderBaseNetDisconnect, vbHide)
            Else
                MsgBox("日期格式有误，请检查！")
                Exit Sub
            End If
        Next
        For i = 2 To xlsheet.UsedRange.Rows.Count
            If ((IsNumeric(xlsheet.Cells(i, 6).value)) Or (CStr(xlsheet.Cells(i, 6).value) = "")) Then
                'shell(SFOrderBaseNetDisconnect, vbHide)
            Else
                MsgBox("单价格式有误，请检查！")
                Exit Sub
            End If
        Next
        '******************************************* 输入订单内容 **********************************************************************
        Dim ID As Integer
        'List down the SFOrder
        SQLstring = "Select * From SFOrderBase order by ID"

        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 3)
        rs.MoveLast()
        ID = rs.Fields("ID").Value

        For i = 2 To xlsheet.UsedRange.Rows.Count
            ID = ID + 1
            rs.AddNew()
            rs.Fields("ID").Value = ID
            rs.Fields("DWGInfo").Value = TextBox2.Text + CStr(xlsheet.Cells(i, 1).value) + TextBox3.Text
            rs.Fields("CustDWG").Value = TextBox2.Text + CStr(xlsheet.Cells(i, 1).value)
            rs.Fields("CustCode").Value = TextBox2.Text
            rs.Fields("CustOrder").Value = xlsheet.Cells(i, 3).value
            rs.Fields("SFOrder").Value = TextBox3.Text
            rs.Fields("DrawNo").Value = xlsheet.Cells(i, 1).value
            rs.Fields("Creator").Value = ChineseName
            'rs.Fields("Unit").Value = IIf(IsDBNull(DataGridView2.Item(4, i),"件",DataGridView2.Item(4, i).Value)
            If (CStr(xlsheet.Cells(i, 5).value) = "") Then
                rs.Fields("Unit").Value = "件"
            Else
                rs.Fields("Unit").Value = CStr(xlsheet.Cells(i, 5).value)
            End If
            rs.Fields("CDate").Value = Now.Date
            rs.Fields("PartName").Value = CStr(xlsheet.Cells(i, 2).value)
            'rs.Fields("Ver").Value = DataGridView2.Item(2, i).Value
            rs.Fields("DDate").Value = xlsheet.Cells(i, 7).value
            rs.Fields("Comment").Value = xlsheet.Cells(i, 8).value
            rs.Fields("Qty").Value = xlsheet.Cells(i, 4).value
            rs.Fields("UnitP").Value = xlsheet.Cells(i, 6).value
            rs.Fields("Currency").Value = Currency
            'rs.Fields("RMBUnitP").Value = DataGridView2.Item(5, i).Value
            rs.Fields("OrderType").Value = TextBox4.Text
            rs.Fields("Status").Value = "完成下单"
            rs.Update()
        Next

        '*******************************************************************************************************************************
        TextBox2.Clear()
        TextBox4.Clear()
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        ss.DisplayAlerts = False
        xlbook.Close()
        ss.Quit()
        xlsheet = Nothing
        xlbook = Nothing
        ss = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)
        MsgBox("输入生产单完成！")
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox4.Text = ComboBox2.SelectedItem.ToString
        TextBox3.Focus()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ss = CreateObject("Excel.Application")
        xlbook = ss.Workbooks.Add
        If xlbook.Sheets.Count > 1 Then
            For i = xlbook.Sheets.Count To 2
                xlbook.Sheets(i).delete()
            Next
        End If

        ss.Visible = True
        xlsheet = xlbook.Sheets(1)
        xlsheet.Name = "test"
        xlsheet.Cells(1, 1) = "图号"
        xlsheet.Cells(1, 2) = "零件名称"
        xlsheet.Cells(1, 3) = "客户订单号"
        xlsheet.Cells(1, 4) = "数量"
        xlsheet.Cells(1, 5) = "单位"
        xlsheet.Cells(1, 6) = "单价"
        xlsheet.Cells(1, 7) = "交期"
        xlsheet.Cells(1, 8) = "备注"
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Currency = "EUR"
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Currency = "USD"
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Currency = "GBP"
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        Currency = "RMB"
    End Sub
End Class