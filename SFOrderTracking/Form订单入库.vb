Public Class Form订单入库

    Private Sub Form14_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form14_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'shell(SFOrderBaseNetConnect, vbHide)
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring, CNString As String
        On Error Resume Next
        '********************************** combobox1connection *****************************************************************
        'List down the SFOrder
        cn.Open(CNsfmdb)
        SQLstring = "Select CustCode From CustList order by CustCode"
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()

        Do While Not rs.EOF
            ComboBox2.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        rs.Close()
        '*************************** end of combobox1 connection *****************************************************************
        '****************************************  ************************************************************************
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("CustCode", "客户")
        DataGridView1.Columns.Add("DDate", "交期")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("PartName", "零件名称")
        DataGridView1.Columns.Add("Qty", "订单数量")
        DataGridView1.Columns.Add("Unit", "单位")
        'DataGridView2.Columns.Add("日期", "日期")
        '****************************************  ********************************************************************
        '****************************************  ************************************************************************
        DataGridView2.Columns.Add("SFOrder", "赛峰单号")
        DataGridView2.Columns.Add("CustCode", "客户")
        DataGridView2.Columns.Add("DDate", "交期")
        DataGridView2.Columns.Add("DrawNo", "图号")
        DataGridView2.Columns.Add("PartName", "零件名称")
        DataGridView2.Columns.Add("Qty", "订单数量")
        DataGridView2.Columns.Add("Unit", "单位")
        DataGridView2.Columns.Add("InStoreQty", "入库数量")
        DataGridView2.Columns.Add("Comment", "备注")
        DataGridView2.Columns.Add("Location", "仓库位置")
        DataGridView2.Columns("InStoreQty").DefaultCellStyle.BackColor = Color.Pink
        DataGridView2.Columns("Comment").DefaultCellStyle.BackColor = Color.Pink
        DataGridView2.Columns("Location").DefaultCellStyle.BackColor = Color.Pink
        For i = 0 To DataGridView2.Columns.Count - 1
            DataGridView2.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        'DataGridView2.Columns.Add("日期", "日期")
        '****************************************  ********************************************************************
        rs = Nothing
        cn = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        DataGridView2.Rows.Add(1)
        DataGridView2.Item("SFOrder", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("SFOrder", e.RowIndex).Value
        DataGridView2.Item("CustCode", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("CustCode", e.RowIndex).Value
        DataGridView2.Item("DrawNo", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("DrawNo", e.RowIndex).Value
        DataGridView2.Item("DDate", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("DDate", e.RowIndex).Value
        DataGridView2.Item("Qty", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("Qty", e.RowIndex).Value
        DataGridView2.Item("PartName", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("PartName", e.RowIndex).Value
        DataGridView2.Item("Unit", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("Unit", e.RowIndex).Value
        DataGridView2.Item("InStoreQty", DataGridView2.Rows.Count - 1).ReadOnly = False
        DataGridView2.Item("Comment", DataGridView2.Rows.Count - 1).ReadOnly = False
        DataGridView2.Item("Location", DataGridView2.Rows.Count - 1).ReadOnly = False

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error Resume Next
        DataGridView2.Rows.RemoveAt(DataGridView2.SelectedCells(0).RowIndex)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '*************************** 入库记录 ************************************************************************
        For i = 0 To DataGridView2.Rows.Count - 1
            If (IsNumeric(DataGridView2.Item("InStoreQty", i).Value) = False) Then
                MsgBox("请检查入库数量！")
                Exit Sub
            End If
            If (IsDBNull(DataGridView2.Item("Location", i).Value) = True) Then
                MsgBox("请检查仓库位置！")
                Exit Sub
            End If
        Next
        'shell(SFOrderBaseNetConnect, vbHide)
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring, CNString As String
        On Error Resume Next
        'DataGridView1.Rows.Clear()
        '**********************************  *****************************************************************
        'List down the SFOrder
        cn.Open(CNsfmdb)
        Dim ID As Integer
        For i = 0 To DataGridView2.Rows.Count - 1
            SQLstring = "Select * from SFInStore order by ID"
            ID = rs.Fields("ID").Value + 1
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 3)
            rs.MoveLast()
            rs.AddNew()
            rs.Fields("ID").Value = ID
            rs.Fields("SFOrder").Value = DataGridView2.Item("SFOrder", i).Value
            rs.Fields("CustCode").Value = DataGridView2.Item("CustCode", i).Value
            rs.Fields("DrawNo").Value = DataGridView2.Item("DrawNo", i).Value
            rs.Fields("DDate").Value = DataGridView2.Item("DDate", i).Value
            rs.Fields("Qty").Value = DataGridView2.Item("Qty", i).Value
            rs.Fields("PartName").Value = DataGridView2.Item("PartName", i).Value
            rs.Fields("Unit").Value = DataGridView2.Item("Unit", i).Value
            rs.Fields("InStoreQty").Value = DataGridView2.Item("InStoreQty", i).Value
            rs.Fields("Comment").Value = DataGridView2.Item("Comment", i).Value
            rs.Fields("Location").Value = DataGridView2.Item("Location", i).Value
            rs.Fields("Creator").Value = ChineseName
            rs.Fields("InStoreDate").Value = DateValue(Now)
            rs.Update()
            rs.Close()
            SQLstring = "Select SFOrder,DrawNo,Status From SFOrderBase where ((SFOrder=" & "'" & DataGridView2.Item("SFOrder", i).Value & "'" & ") and (DrawNo=" & "'" & DataGridView2.Item("DrawNo", i).Value & "'" & "))"
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 3)
            rs.MoveFirst()
            rs.Fields("Status").Value = "完成入库"
            rs.Update()
            rs.Close()
        Next

        'shell(SFOrderBaseNetDisconnect, vbHide)
        DataGridView2.Rows.Clear()
        DataGridView1.Rows.Clear()
        MsgBox("完成入库记录！")
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'shell(SFOrderBaseNetConnect, vbHide)
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring, CNString As String
        On Error Resume Next
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        '********************************** combobox2connection *****************************************************************
        'List down the SFOrder
        SQLstring = "Select SFOrder,CustCode,DrawNo,PartName,Qty,DDate,Unit From SFOrderBase where ((CustCode=" & "'" & ComboBox2.SelectedItem.ToString & "'" & ") and ((Status<>" & "'" & "取消生产单" & "'" & ") and (Status<>" & "'" & "完成入库" & "'" & ")))"
        cn.Open(CNsfmdb)
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)
        If (rs.RecordCount = 0) Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            'shell(SFOrderBaseNetDisconnect, vbHide)
            MsgBox("该客户订单已全部入库！")
            Exit Sub
        End If
        rs.MoveFirst()
        Do While Not rs.EOF
            DataGridView1.Rows.Add(1)
            DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
            DataGridView1.Item("CustCode", DataGridView1.Rows.Count - 1).Value = rs.Fields("CustCode").Value
            DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
            DataGridView1.Item("DDate", DataGridView1.Rows.Count - 1).Value = rs.Fields("DDate").Value
            DataGridView1.Item("Qty", DataGridView1.Rows.Count - 1).Value = rs.Fields("Qty").Value
            DataGridView1.Item("PartName", DataGridView1.Rows.Count - 1).Value = rs.Fields("PartName").Value
            DataGridView1.Item("Unit", DataGridView1.Rows.Count - 1).Value = rs.Fields("Unit").Value
            rs.MoveNext()
        Loop
        rs.Close()
        '*************************** end of combobox2 connection *****************************************************************
        rs = Nothing
        cn = Nothing
        DataGridView1.Focus()
        'shell(SFOrderBaseNetDisconnect, vbHide)

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs)
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            'shell(SFOrderBaseNetConnect, vbHide)
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring, CNString As String
            On Error Resume Next
            DataGridView1.Rows.Clear()
            DataGridView2.Rows.Clear()
            '********************************** combobox2connection *****************************************************************
            'List down the SFOrder
            SQLstring = "Select SFOrder,CustCode,DrawNo,PartName,Qty,DDate,Unit From SFOrderBase where ((DrawNo like" & "'%" & TextBox1.Text & "%'" & ") and (Status not in ('取消生产单','完成入库')))"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                'shell(SFOrderBaseNetDisconnect, vbHide)
                MsgBox("该图号订单已全部入库！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While Not rs.EOF
                DataGridView1.Rows.Add(1)
                DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("CustCode", DataGridView1.Rows.Count - 1).Value = rs.Fields("CustCode").Value
                DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("DDate", DataGridView1.Rows.Count - 1).Value = rs.Fields("DDate").Value
                DataGridView1.Item("Qty", DataGridView1.Rows.Count - 1).Value = rs.Fields("Qty").Value
                DataGridView1.Item("PartName", DataGridView1.Rows.Count - 1).Value = rs.Fields("PartName").Value
                DataGridView1.Item("Unit", DataGridView1.Rows.Count - 1).Value = rs.Fields("Unit").Value
                rs.MoveNext()
            Loop
            rs.Close()
            '*************************** end of combobox2 connection *****************************************************************
            rs = Nothing
            cn = Nothing
            DataGridView1.Focus()
            'shell(SFOrderBaseNetDisconnect, vbHide)

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            'shell(SFOrderBaseNetConnect, vbHide)
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring, CNString As String
            On Error Resume Next
            DataGridView1.Rows.Clear()
            DataGridView2.Rows.Clear()
            '********************************** combobox1connection *****************************************************************
            'List down the SFOrd
            SQLstring = "Select SFOrder,CustCode,DrawNo,PartName,Qty,DDate,Unit From SFOrderBase where ((SFOrder=" & "'" & TextBox2.Text & "'" & ") and (((Status not in ('取消生产单','完成入库')) or (isnull(status,'N')='N'))))"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                'shell(SFOrderBaseNetDisconnect, vbHide)
                MsgBox("该订单已全部入库！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While Not rs.EOF
                DataGridView1.Rows.Add(1)
                DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("CustCode", DataGridView1.Rows.Count - 1).Value = rs.Fields("CustCode").Value
                DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("DDate", DataGridView1.Rows.Count - 1).Value = rs.Fields("DDate").Value
                DataGridView1.Item("Qty", DataGridView1.Rows.Count - 1).Value = rs.Fields("Qty").Value
                DataGridView1.Item("PartName", DataGridView1.Rows.Count - 1).Value = rs.Fields("PartName").Value
                DataGridView1.Item("Unit", DataGridView1.Rows.Count - 1).Value = rs.Fields("Unit").Value
                rs.MoveNext()
            Loop
            rs.Close()
            '*************************** end of combobox1 connection *****************************************************************
            rs = Nothing
            cn = Nothing
            DataGridView1.Focus()
            'shell(SFOrderBaseNetDisconnect, vbHide)
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class