Public Class Form收料扫描

    Private Sub Form7_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Visible = True
        Form主窗体.Show()
    End Sub
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '********************************** combobox1connection *****************************************************************
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * from ReceiveList order by '序号'"

        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()

        Do While Not rs.EOF
            ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        rs.Close()
        '*************************** end of combobox1 connection *****************************************************************
        '**************************** title of datagridview1 intialization ******************************************************
        SQLstring = "Select * from ProcRecord"
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()

        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        'MsgBox("First Records in ProcCard is " & rs.RecordCount)
        For i = 1 To rs.Fields.Count
            If (rs.Fields(i - 1).Name = "ProcCode") Then
                TextBox6.Text = i - 1 '将ProcCode列号赋值给TextBox6
            End If
            DataGridView1.Columns.Add(rs.Fields(i - 1).Name, rs.Fields(i - 1).Name)
            DataGridView1.Columns(i - 1).ReadOnly = True
            If ((rs.Fields(i - 1).Name = "OKQty") Or (rs.Fields(i - 1).Name = "ScrapQty")) Then
                DataGridView1.Columns(i - 1).ReadOnly = False
                DataGridView1.Columns(i - 1).DataGridView.DefaultCellStyle.ForeColor = Color.Red
            Else
                DataGridView1.Columns(i - 1).ReadOnly = True
                DataGridView1.Columns(i - 1).DataGridView.DefaultCellStyle.ForeColor = Color.Black
            End If
        Next
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        '************************* end of title of datagridview *****************************************************************
        'shell(SFOrderBaseNetDisconnect, vbHide)
    End Sub

    Private Sub Form7_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        Form主窗体.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox1.Text = ComboBox1.SelectedItem.ToString
        TextBox1.Visible = True
        TextBox3.Enabled = True
        TextBox3.Focus()
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If (e.KeyCode = 13) Then
            'shell(SFOrderBaseNetConnect, vbHide)
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring, CNString As String

            On Error Resume Next

            'SQLstring = "Select DrawNo, SFOrder,CustCode CustDWG, ProcSN, ProcName, ProcDesc, Qty, ProcQty from ProcCard"
            SQLstring = "Select * from SFHR where 工号=" & "'" & (TextBox3.Text) & "'"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                'shell(SFOrderBaseNetDisconnect, vbHide)
                MsgBox("该工号不存在，请找系统管理员！")
                Exit Sub
            End If
            'MsgBox("First Records in ProcCard is " & rs.RecordCount)
            TextBox4.Text = rs.Fields("姓名").Value
            TextBox4.Visible = True
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            TextBox3.Enabled = False
            TextBox2.Enabled = True
            TextBox2.Focus()
        End If
        'shell(SFOrderBaseNetDisconnect, vbHide)
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If (e.KeyCode = 13) Then
            'shell(SFOrderBaseNetConnect, vbHide)
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String

            On Error Resume Next

            SQLstring = "Select * from ProcCard where ProcCode=" & "'" & (TextBox2.Text) & "'"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                MsgBox("系统工艺信息缺乏！")
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                TextBox2.Clear()
                TextBox2.Focus()
                'shell(SFOrderBaseNetDisconnect, vbHide)
                Exit Sub
            End If
            Dim DWGInfo As String
            rs.MoveFirst()
            DataGridView1.Rows.Add(1)
            Dim DataRow As Integer
            DataRow = DataGridView1.RowCount - 1
            DataGridView1.Columns("OKQty").ReadOnly = False
            DataGridView1.Columns("ScrapQty").ReadOnly = False

            DataGridView1.Item("DWGInfo", DataRow).Value = rs.Fields("DWGInfo").Value
            DataGridView1.Item("CustDWG", DataRow).Value = rs.Fields("CustDWG").Value
            DataGridView1.Item("ProcCode", DataRow).Value = rs.Fields("ProcCode").Value
            DataGridView1.Item("Operator", DataRow).Value = TextBox4.Text
            DataGridView1.Item("ProcQty", DataRow).Value = rs.Fields("ProcQty").Value
            DataGridView1.Item("ScrapQty", DataRow).Value = "0"
            DataGridView1.Item("Qty", DataRow).Value = rs.Fields("Qty").Value
            DataGridView1.Item("ProcQty", DataRow).Value = rs.Fields("ProcQty").Value
            DataGridView1.Item("RecordTime", DataRow).Value = Now
            DataGridView1.Item("RecordType", DataRow).Value = "收料"
            DataGridView1.Item("Status", DataRow).Value = (Mid(rs.Fields("ProcCode").Value, Len(rs.Fields("ProcCode").Value) - 2, 3)) & TextBox1.Text + "-收料"


            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            TextBox2.Clear()
            TextBox2.Focus()
            'shell(SFOrderBaseNetDisconnect, vbHide)
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'shell(SFOrderBaseNetConnect, vbHide)
        Dim ID As Integer

        If (TextBox1.Text = "") Then
            'shell(SFOrderBaseNetDisconnect, vbHide)
            MsgBox("请选择收料工序！")
            Exit Sub
        End If
        If (DataGridView1.RowCount = 0) Then
            'shell(SFOrderBaseNetDisconnect, vbHide)
            MsgBox("没有收料单！")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        Dim DWGInfo As String

        cn.Open(CNsfmdb)

        For i = 0 To DataGridView1.RowCount - 1
            SQLstring = "Select * from ProcRecord order by ID"
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 3)
            rs.MoveLast()
            ID = rs.Fields("ID").Value
            rs.AddNew()
            rs.Fields("ID").Value = ID + 1
            rs.Fields("DWGInfo").Value = DataGridView1.Item("DWGInfo", i).Value
            rs.Fields("CustDWG").Value = DataGridView1.Item("CustDWG", i).Value
            rs.Fields("ProcCode").Value = DataGridView1.Item("ProcCode", i).Value
            rs.Fields("Operator").Value = DataGridView1.Item("Operator", i).Value
            rs.Fields("OKQty").Value = DataGridView1.Item("OKQty", i).Value
            rs.Fields("ScrapQty").Value = DataGridView1.Item("ScrapQty", i).Value
            rs.Fields("Qty").Value = DataGridView1.Item("Qty", i).Value
            rs.Fields("ProcQty").Value = DataGridView1.Item("ProcQty", i).Value
            rs.Fields("RecordTime").Value = DataGridView1.Item("RecordTime", i).Value
            rs.Fields("RecordType").Value = DataGridView1.Item("RecordType", i).Value
            rs.Fields("Status").Value = Mid(DataGridView1.Item("ProcCode", i).Value, Len(DataGridView1.Item("ProcCode", i).Value) - 2, 3) & TextBox1.Text + "-收料"
            rs.Update()
            rs.Close()
            '888888888888888888888888888888 更新订单状况 status 8888888888888888888888888888888888888888888
            DWGInfo = DataGridView1.Item("DWGInfo", i).Value
            'cn.Open(CNsfmdb)
            'SQLstring = "Select * from SFOrderBase where DWGInfo=" & "'" & DWGInfo & "'"
            SQLstring = "Select * from SFOrderBase where DWGInfo like " & "'%" & DWGInfo & "%'"
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 3)
            rs.MoveFirst()

            rs.Fields("Status").Value = Mid(DataGridView1.Item("ProcCode", i).Value, Len(DataGridView1.Item("ProcCode", i).Value) - 2, 3) & TextBox1.Text + "-收料"
            rs.Update()
            rs.Close()
            '8888888888888888888888888888 end of 更新订单状况 status 88888888888888888888888888888888888888
        Next
        cn.Close()
        rs = Nothing
        cn = Nothing
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        DataGridView1.Rows.Clear()
        'shell(SFOrderBaseNetDisconnect, vbHide)
        MsgBox("收料扫描完成！")

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub
End Class