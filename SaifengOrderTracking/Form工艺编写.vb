Public Class Form工艺编写
    Private Sub Form11_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '*************************** end of combobox1 connection *****************************************************************
        Dim QueryText As String

        QueryText = "Select * from Resource"
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset

        cn.Open(CNsfmdb)
        rs.Open(QueryText, cn, 1, 1)
        Dim dd As New DataGridViewComboBoxColumn
        Do While rs.EOF = False
            dd.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        '**************************************** 编写工艺列表 ************************************************************************
        DataGridView2.Columns.Add("工序号", "工序号")
        DataGridView2.Columns.Add(dd)
        DataGridView2.Columns.Add("工序内容和注意事项", "工序内容和注意事项")
        DataGridView2.Columns.Add("准备工时h", "准备工时h")
        DataGridView2.Columns.Add("单件工时h", "单件工时h")
        DataGridView2.Columns.Add("该工序预计单件成本¥", "该工序预计单件成本¥")
        For i = 0 To DataGridView2.Columns.Count - 1
            DataGridView2.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

        Next
        DataGridView2.Columns("工序号").ReadOnly = True
        DataGridView2.Columns(1).HeaderText = "资源名称"
        DataGridView2.Columns(1).Name = "资源名称"

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'shell(SFOrderBaseNetConnect, vbHide)
        TextBox_已选图号.Text = ComboBox2.SelectedItem.ToString '已选图号
        TextBox3.Text = TextBox5.Text + TextBox_已选图号.Text
        TextBox4.Text = TextBox5.Text + TextBox_已选图号.Text + TextBox_生产单号.Text

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        'Dim SQLstring As String
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select * From SFOrderBase where ((SFOrder=" & "'" & TextBox_生产单号.Text & "'" & ") and (DrawNo=" & "'" & ComboBox2.SelectedItem.ToString & "'" & "))"
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        TextBox6.Text = rs.Fields("Qty").Value
        TextBox7.Text = rs.Fields("Qty").Value
        '*************************** end of combobox2 connection *****************************************************************
        rs.Close()
        '88888888888888888888888888888888 对应图号工艺列表 88888888888888888888888888888888888888888888888888888888888888888888888888
        'SQLstring = "Select DrawNo, SFOrder,CustCode CustDWG, ProcSN, ProcName, ProcDesc, Qty, ProcQty From ProcCard"
        'SQLstring = "Select DWGInfo,CustDWG,DrawNo,Qty,CustCode From CardList where DrawNo=" & "'" & ComboBox2.SelectedItem.ToString & "'"
        SQLstring = "Select DWGInfo,CustDWG,DrawNo,Qty,CustCode From ProcCard where DrawNo=" & "'" & ComboBox2.SelectedItem.ToString & "' group by DWGInfo,CustDWG,DrawNo,Qty,CustCode order by DWGInfo,CustDWG,DrawNo,Qty,CustCode"
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)
        If (rs.RecordCount = 0) Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            'shell(SFOrderBaseNetDisconnect, vbHide)
            Exit Sub
        End If
        rs.MoveFirst()
        For i = 1 To rs.Fields.Count
            DataGridView1.Columns.Add(rs.Fields(i - 1).Name, rs.Fields(i - 1).Name)
        Next
        DataGridView1.Rows.Add(rs.RecordCount)

        Dim DataRow As Integer
        DataRow = 0
        Do While rs.EOF = False
            For i = 1 To rs.Fields.Count
                DataGridView1.Item(i - 1, DataRow).Value = rs.Fields(i - 1).Value
            Next i
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop
        rs.Close()
        '88888888888888888888888888 end of 对应图号工艺列表结束 8888888888888888888888888888888888888888888888888888888888888888888888
        rs = Nothing
        cn = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If (e.RowIndex >= 0) Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            'Dim SQLstring As String
            Dim SQLstring As String
            Dim ProcSN As Integer

            DataGridView2.Rows.Clear()
            DataGridView2.Columns.Clear()

            DataGridView2.Columns.Add("工序号", "工序号")
            DataGridView2.Columns.Add("资源名称", "资源名称")
            DataGridView2.Columns.Add("工序内容和注意事项", "工序内容和注意事项")
            DataGridView2.Columns.Add("准备工时h", "准备工时h")
            DataGridView2.Columns.Add("单件工时h", "单件工时h")
            DataGridView2.Columns.Add("该工序预计单件成本¥", "该工序预计单件成本¥")
            For i = 0 To DataGridView2.Columns.Count - 1
                DataGridView2.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            DataGridView2.Columns("工序号").ReadOnly = True

            cn.Open(CNsfmdb)
            SQLstring = "Select * From ProcCard where DWGInfo=" & "'" & DataGridView1.Item(0, e.RowIndex).Value & "' order by ProcSN"
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            DataGridView2.Rows.Add(rs.RecordCount)

            Dim DataRow As Integer
            DataRow = 0
            ProcSN = 10
            Do While rs.EOF = False
                For i = 1 To rs.Fields.Count
                    'DataGridView1.Item(i - 1, DataRow).Value = rs.Fields(i - 1).Value
                    If (rs.Fields(i - 1).Name = "ProcSN") Then
                        DataGridView2.Item(0, DataRow).Value = ProcSN
                    End If
                    If (rs.Fields(i - 1).Name = "ProcName") Then
                        DataGridView2.Item(1, DataRow).Value = rs.Fields(i - 1).Value
                    End If
                    If (rs.Fields(i - 1).Name = "ProcDesc") Then
                        DataGridView2.Item(2, DataRow).Value = rs.Fields(i - 1).Value
                    End If
                    If (rs.Fields(i - 1).Name = "PreTime") Then
                        DataGridView2.Item(3, DataRow).Value = rs.Fields(i - 1).Value
                    End If
                    If (rs.Fields(i - 1).Name = "UnitTime") Then
                        DataGridView2.Item(4, DataRow).Value = rs.Fields(i - 1).Value
                    End If
                    'DataGridView2.Item(6, DataRow).Value = Val(rs.Fields("PreTime").Value) + Val(TextBox7.Text) * Val(rs.Fields("UnitTime").Value)
                Next i
                DataRow = DataRow + 1
                ProcSN = ProcSN + 10
                rs.MoveNext()
            Loop

            rs.Close()
            '显示当前工件的报价
            SQLstring = "Select UnitP from SFOrderBase where SFOrder=" & "'" & TextBox_生产单号.Text & "' and DrawNo='" + TextBox_已选图号.Text + "'"
            rs.Open(SQLstring, cn, 1, 1)
            Dim price As String
            price = rs.Fields(0).Value.ToString()
            Dim str As String
            If String.IsNullOrEmpty(price) Then
                str = "无报价"
            Else
                str = price
            End If
            Label_Price.Text = "当前报价￥：" + str

            rs.Close()
            '*************************** end of combobox2 connection *****************************************************************
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (TextBox_已选图号.Text = "") Then
            MsgBox("请选择图号！")
            Exit Sub
        End If
        If (DataGridView2.Item("资源名称", 0).Value = "") Then
            MsgBox("请先编写工艺卡！")
            Exit Sub
        End If

        '888888888888888888888888888888888888888888888888 存入已编写工艺 888888888888888888888888888888888888888888888888888888888888888888888888888
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String

        On Error Resume Next

        cn.Open(CNsfmdb)


        Dim 是否需要预估成本, 是否有成本单价, Label预计成本 As String
        Dim 成本单价, 预计单件成本, 预计总成本 As Decimal

        For i = 0 To DataGridView2.Rows.Count - 1
            Dim ProcName As String
            ProcName = DataGridView2.Item(1, i).Value '工序名称
            If String.IsNullOrEmpty(ProcName) = False Then  '如果工序名称不是空
                SQLstring = "select * From Resource where Resource='" + ProcName + "'"
                rs.Open(SQLstring, cn, 1, 1)
                是否需要预估成本 = rs.Fields(1).Value
                是否有成本单价 = rs.Fields(2).Value
                成本单价 = rs.Fields(3).Value
                If 是否需要预估成本 = "是" Then
                    If 是否有成本单价 = "是" Then
                        If DataGridView2.Item(4, i).Value = Nothing Or DataGridView2.Item(4, i).Value.ToString() = "" Then  '该工序单件工时判断用户输入没有,如果用户没输入则报警提示
                            MsgBox($"第{i + 1}行请输入单件工时！")
                            Return
                        Else
                            If DataGridView2.Item(5, i).Value = Nothing Or DataGridView2.Item(5, i).Value.ToString() = "" Then   '该工序预计单件成本列判断用户输入没有,如果用户没输入
                                DataGridView2.Item(5, i).Value = Decimal.Parse(DataGridView2.Item(4, i).Value) * 成本单价
                            End If
                        End If

                    Else '如果没有成本单价
                        If DataGridView2.Item(5, i).Value = Nothing Or DataGridView2.Item(5, i).Value.ToString() = "" Then   '该工序预计单件成本列判断用户输入没有,如果用户没输入
                            MsgBox($"第{i + 1}行请输入该工序预计单件成本！")
                            Return
                        End If
                    End If
                Else
                    If DataGridView2.Item(5, i).Value = Nothing Or DataGridView2.Item(5, i).Value.ToString() = "" Then
                        DataGridView2.Item(5, i).Value = 0
                    End If
                End If
                rs.Close()
            End If
            预计单件成本 = 预计单件成本 + Decimal.Parse(DataGridView2.Item(5, i).Value)
            预计总成本 = 预计单件成本 * Val(TextBox6.Text) '乘以工艺数量
        Next

        Label_预计单件成本.Text = "预计单件成本￥:" + 预计单件成本.ToString()
        Label_预计总成本.Text = "预计总成本￥:" + 预计总成本.ToString()

        '显示当前工件的报价
        SQLstring = "Select UnitP from SFOrderBase where SFOrder=" & "'" & TextBox_生产单号.Text & "' and DrawNo='" + TextBox_已选图号.Text + "'"
        rs.Open(SQLstring, cn, 1, 1)
        Dim price As String
        price = rs.Fields(0).Value.ToString()
        Dim str As String
        If String.IsNullOrEmpty(price) Then
            str = "无报价"
        Else
            str = price
        End If

        If MsgBox($"单件报价¥{str} 预计单件成本¥:{预计单件成本} 工艺数量{Val(TextBox6.Text)} 预计总成本¥:{预计总成本}", MsgBoxStyle.OkCancel, "请确认是否无误") = MsgBoxResult.Ok Then

            SQLstring = "select * From ProcCard order by ID"
            Dim ID As Integer

            rs.Open(SQLstring, cn, 1, 3)
            rs.MoveLast()
            ID = rs.Fields("ID").Value
            For i = 0 To DataGridView2.Rows.Count - 1
                If ((DataGridView2.Item(0, i).Value.ToString = "") Or DataGridView2.Item(1, i).Value.ToString = "") Then
                    Exit For
                Else
                    ID = ID + 1
                    rs.AddNew()
                    rs.Fields("ID").Value = ID
                    rs.Fields("DWGInfo").Value = TextBox4.Text
                    rs.Fields("CustDWG").Value = TextBox3.Text
                    rs.Fields("SFOrder").Value = TextBox_生产单号.Text
                    rs.Fields("DrawNo").Value = TextBox_已选图号.Text
                    rs.Fields("ProcMaker").Value = ChineseName
                    rs.Fields("ProcDate").Value = Now.Date
                    rs.Fields("Qty").Value = Val(TextBox7.Text)
                    rs.Fields("ProcQty").Value = Val(TextBox6.Text) '工艺数量
                    rs.Fields("CustCode").Value = TextBox5.Text
                    rs.Fields("ProcSN").Value = DataGridView2.Item(0, i).Value
                    rs.Fields("ProcName").Value = DataGridView2.Item(1, i).Value
                    rs.Fields("ProcDesc").Value = DataGridView2.Item(2, i).Value
                    rs.Fields("PreTime").Value = DataGridView2.Item(3, i).Value
                    rs.Fields("UnitTime").Value = DataGridView2.Item(4, i).Value
                    rs.Fields("ProcCode").Value = TextBox_生产单号.Text & TextBox_已选图号.Text & DataGridView2.Item(0, i).Value
                    Dim 该工序预计单件工时, 该工序预计单件成本, 工艺数量 As Decimal
                    Decimal.TryParse(DataGridView2.Item(4, i).Value, 该工序预计单件工时)
                    Decimal.TryParse(DataGridView2.Item(5, i).Value, 该工序预计单件成本)
                    Decimal.TryParse(Val(TextBox6.Text), 工艺数量)
                    rs.Fields("该工序预计单件工时").Value = 该工序预计单件工时
                    rs.Fields("该工序预计单件成本").Value = 该工序预计单件成本
                    rs.Fields("该工序预计总工时").Value = 该工序预计单件工时 * 工艺数量
                    rs.Fields("该工序预计总成本").Value = 该工序预计单件成本 * 工艺数量
                    rs.Update()
                End If
            Next
            rs.Close()

            SQLstring = "select * From SFOrderBase where ((SFOrder=" & "'" & TextBox_生产单号.Text & "'" & ")" & "and (DrawNo=" & "'" & TextBox_已选图号.Text & "'" & "))"
            rs.Open(SQLstring, cn, 1, 3)
            rs.MoveFirst()
            rs.Fields("Status").Value = "完成工艺"
            rs.Update()
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing

            DataGridView2.Rows.Clear()
            'ComboBox2.Items.Clear()
            ComboBox2.Items.Remove(TextBox_已选图号.Text)
            ComboBox2.Text = ""

            '*************************** end of combobox2 connection *****************************************************************
            MsgBox(TextBox_生产单号.Text & "    " & TextBox_已选图号.Text & "  工艺卡编写完成！")
            TextBox_已选图号.Clear()
            '888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888
        Else
            Return
        End If


    End Sub

    'lcz20181220更改
    'Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
    '    DataGridView2.Rows.Clear()
    '    DataGridView2.Rows.Add(15)
    '    Dim DataRow, ProcSN As Integer
    '    DataRow = 0
    '    ProcSN = 100
    '    For i = 0 To DataGridView2.Rows.Count - 2
    '        DataGridView2.Item(0, i).Value = ProcSN
    '        ProcSN = ProcSN + 10
    '    Next i

    'End Sub
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()

        '*************************** end of combobox1 connection *****************************************************************
        Dim QueryText As String

        QueryText = "Select * from Resource"
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset

        cn.Open(CNsfmdb)
        rs.Open(QueryText, cn, 1, 1)
        Dim dd As New DataGridViewComboBoxColumn '定义工序名称下拉清单
        Do While rs.EOF = False
            dd.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        '**************************************** 编写工艺列表 ************************************************************************
        DataGridView2.Columns.Add("工序号", "工序号")
        DataGridView2.Columns.Add(dd)
        DataGridView2.Columns.Add("工序内容和注意事项", "工序内容和注意事项")
        DataGridView2.Columns.Add("准备工时h", "准备工时h")
        DataGridView2.Columns.Add("单件工时h", "单件工时h")
        DataGridView2.Columns.Add("该工序预计单件成本¥", "该工序预计单件成本¥")
        For i = 0 To DataGridView2.Columns.Count - 1
            DataGridView2.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        DataGridView2.Columns("工序号").ReadOnly = True
        DataGridView2.Columns(1).HeaderText = "资源名称"
        DataGridView2.Columns(1).Name = "资源名称"

        DataGridView2.Rows.Add(15)
        Dim DataRow, ProcSN As Integer
        DataRow = 0
        ProcSN = 10
        For i = 0 To DataGridView2.Rows.Count - 2
            DataGridView2.Item(0, i).Value = ProcSN
            ProcSN = ProcSN + 5
        Next i

    End Sub

    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If (e.KeyCode = 13) Then
            'shell(SFOrderBaseNetConnect, vbHide)
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            'Dim SQLstring As String
            Dim SQLstring As String
            cn.Open(CNsfmdb)

            '88888888888888888888888888888888 对应图号工艺列表 88888888888888888888888888888888888888888888888888888888888888888888888888
            'SQLstring = "Select DrawNo, SFOrder,CustCode CustDWG, ProcSN, ProcName, ProcDesc, Qty, ProcQty From ProcCard"
            SQLstring = "Select DWGInfo,CustDWG,DrawNo,Qty,CustCode From CardList where DrawNo like " & "'%" & TextBox8.Text & "%'"
            DataGridView1.Columns.Clear()
            DataGridView1.Rows.Clear()
            DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                'shell(SFOrderBaseNetDisconnect, vbHide)
                Exit Sub
            End If
            rs.MoveFirst()

            For i = 1 To rs.Fields.Count
                DataGridView1.Columns.Add(rs.Fields(i - 1).Name, rs.Fields(i - 1).Name)
            Next
            DataGridView1.Rows.Add(rs.RecordCount)

            Dim DataRow As Integer
            DataRow = 0
            Do While rs.EOF = False
                For i = 1 To rs.Fields.Count
                    DataGridView1.Item(i - 1, DataRow).Value = rs.Fields(i - 1).Value
                Next i
                DataRow = DataRow + 1
                rs.MoveNext()
            Loop
            rs.Close()
            '88888888888888888888888888 end of 对应图号工艺列表结束 8888888888888888888888888888888888888888888888888888888888888888888888
            rs = Nothing
            cn = Nothing
            'shell(SFOrderBaseNetDisconnect, vbHide)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DataGridView2.Rows.Add(1)
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        For i = (DataGridView2.SelectedCells(0).RowIndex) To (DataGridView2.SelectedCells(DataGridView2.SelectedCells.Count - 1).RowIndex) Step -1
            DataGridView2.Rows.RemoveAt(i)
        Next
        For i = 0 To DataGridView2.Rows.Count - 1
            DataGridView2.Item("工序号", i).Value = 10 + 5 * i
        Next
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        On Error GoTo ErrorExit
        If (DataGridView2.SelectedCells(0).RowIndex <> 0) Then
            DataGridView2.Rows.Insert(DataGridView2.SelectedCells(0).RowIndex)
            For i = 0 To DataGridView2.Rows.Count - 1
                DataGridView2.Item("工序号", i).Value = 10 + 5 * i
            Next
        End If
ErrorExit:
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_生产单号.KeyDown
        If e.KeyCode = 13 Then
            If TextBox_生产单号.Text = "" Then
                MsgBox("请输入生产单号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            'Dim SQLstring As String
            Dim SQLstring As String
            SQLstring = "Select * From SFOrderBase where SFOrder=" & "'" & TextBox_生产单号.Text & "'"
            TextBox_已选图号.Clear()
            TextBox3.Clear()
            TextBox4.Clear()

            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            rs.MoveFirst()
            TextBox5.Text = rs.Fields("CustCode").Value
            TextBox3.Text = rs.Fields("CustCode").Value + TextBox_已选图号.Text
            TextBox4.Text = rs.Fields("CustCode").Value + TextBox_已选图号.Text + TextBox_生产单号.Text
            rs.Close()
            '********************************** combobox2onnection *****************************************************************
            ComboBox2.Items.Clear()
            SQLstring = "Select DrawNo From SFOrderBase where ((SFOrder=" & "'" & TextBox_生产单号.Text & "'" & ") and (DrawNo not in (Select DrawNo from SFAll.dbo.ProcCard where SFOrder=" & "'" & TextBox_生产单号.Text & "')))"
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                MsgBox("此生产单所有零件均已完成工艺！")
                TextBox_已选图号.Clear()
                TextBox3.Clear()
                TextBox4.Clear()
                rs.Clone()
                cn.Close()
                rs = Nothing
                cn = Nothing
                'shell(SFOrderBaseNetDisconnect, vbHide)
                Exit Sub
            End If
            rs.MoveFirst()

            Do While Not rs.EOF
                ComboBox2.Items.Add(rs("DrawNo").Value)
                rs.MoveNext()
            Loop
            rs.Close()
            '*************************** end of combobox2 connection *****************************************************************
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub DataGridView2_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError
        DataGridView2.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Red
    End Sub
End Class