Public Class Form工艺编写及修改
    Private Sub Form11_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim QueryText As String
        QueryText = "Select * from Resource"
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset

        cn.Open(CNsfmdb)
        rs.Open(QueryText, cn, 1, 1)
        Do While rs.EOF = False
            资源名称.Items.Add(rs(1).Value)
            rs.MoveNext()
        Loop

        DataGridView2.Columns("工序号").Width = 70
        DataGridView2.Columns("资源名称").Width = 130
        DataGridView2.Columns("工序内容和注意事项").AutoSizeMode = DataGridViewAutoSizeColumnsMode.Fill
        DataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        DataGridView2.Columns("准备工时h").Width = 70
        DataGridView2.Columns("单件工时h").Width = 70
        DataGridView2.Columns("该工序预计单件成本").Width = 100

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If (e.RowIndex >= 0) Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            'Dim SQLstring As String
            Dim SQLstring, DWGinfo As String
            DWGinfo = DataGridView1.Item(2, e.RowIndex).Value
            cn.Open(CNsfmdb)
            SQLstring = "Select * From ProcCard where DWGInfo=N'" & DWGinfo & "' order by ProcSN"
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                MsgBox($"DWGInfo= {DWGinfo} 不存在！工艺卡中没有此记录数据！")
                Return
            End If

            DataGridView2.Rows.Clear()
            DataGridView2.Rows.Add(rs.RecordCount)

            Dim DataRow As Integer
            DataRow = 0
            Do While rs.EOF = False
                DataGridView2.Item("工序号", DataRow).Value = rs.Fields("ProcSN").Value
                DataGridView2.Item("资源名称", DataRow).Value = rs.Fields("ProcName").Value
                DataGridView2.Item("工序内容和注意事项", DataRow).Value = rs.Fields("ProcDesc").Value
                DataGridView2.Item("准备工时h", DataRow).Value = rs.Fields("PreTime").Value
                DataGridView2.Item("单件工时h", DataRow).Value = rs.Fields("UnitTime").Value
                DataGridView2.Item("该工序预计单件成本", DataRow).Value = rs.Fields("该工序预计单件成本").Value

                DataRow = DataRow + 1
                rs.MoveNext()
            Loop

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
                是否需要预估成本 = rs.Fields(2).Value
                是否有成本单价 = rs.Fields(3).Value
                成本单价 = rs.Fields(4).Value
                If 是否需要预估成本 = "是" Then
                    If 是否有成本单价 = "是" Then
                        If DataGridView2.Item(4, i).Value = Nothing Or DataGridView2.Item(4, i).Value.ToString() = "" Then  '该工序单件工时判断用户输入没有,如果用户没输入则报警提示
                            MsgBox($"第{i + 1}行[{ProcName}]请输入单件工时！")
                            Return
                        Else
                            DataGridView2.Item(5, i).Value = Decimal.Parse(DataGridView2.Item(4, i).Value) * 成本单价
                            'If DataGridView2.Item(5, i).Value = Nothing Or DataGridView2.Item(5, i).Value.ToString() = "" Then   '该工序预计单件成本列判断用户输入没有,如果用户没输入
                            '    DataGridView2.Item(5, i).Value = Decimal.Parse(DataGridView2.Item(4, i).Value) * 成本单价
                            'End If
                        End If

                    Else '如果没有成本单价
                        If DataGridView2.Item(5, i).Value = Nothing Or DataGridView2.Item(5, i).Value.ToString() = "" Then   '该工序预计单件成本列判断用户输入没有,如果用户没输入
                            MsgBox($"第{i + 1}行[{ProcName}]请输入该工序预计单件成本！")
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

        rs.Close()

        If MsgBox($"单件报价¥{str} 预计单件成本¥:{预计单件成本} 工艺数量{Val(TextBox6.Text)} 预计总成本¥:{预计总成本}", MsgBoxStyle.OkCancel, "请确认是否无误") = MsgBoxResult.Ok Then

            SQLstring = "delete ProcCard where DWGInfo = " & "'" & TextBox_DWGInfo.Text & "'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 3)
            rs.Close()

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
                    rs.Fields("DWGInfo").Value = TextBox_DWGInfo.Text
                    rs.Fields("CustDWG").Value = TextBox3.Text
                    rs.Fields("SFOrder").Value = TextBox_生产单号.Text
                    rs.Fields("DrawNo").Value = TextBox_已选图号.Text
                    rs.Fields("ProcMaker").Value = ChineseName
                    rs.Fields("ProcDate").Value = Now.Date
                    rs.Fields("Qty").Value = Val(TextBox7.Text)
                    rs.Fields("ProcQty").Value = Val(TextBox6.Text) '工艺数量
                    rs.Fields("CustCode").Value = TextBox_CustCode.Text
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

            '*************************** end of combobox2 connection *****************************************************************
            MsgBox(TextBox_生产单号.Text & "    " & TextBox_已选图号.Text & "  工艺卡编写完成！")
            TextBox_已选图号.Clear()

            TextBox生产单号_KeyDown(Nothing, Nothing)
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
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '*************************** end of combobox1 connection *****************************************************************
        DataGridView2.Rows.Clear()
        DataGridView2.Rows.Add(15) '新增15行
        Dim DataRow, ProcSN As Integer
        DataRow = 0
        ProcSN = 10 '第一排工序号为10，后续递增5
        For i = 0 To DataGridView2.Rows.Count - 2
            DataGridView2.Item(0, i).Value = ProcSN
            ProcSN = ProcSN + 5
        Next i

    End Sub

    Private Sub TextBox参考图号_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_参考图号.KeyDown
        If (e.KeyCode = 13) Then
            'shell(SFOrderBaseNetConnect, vbHide)
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            'Dim SQLstring As String
            Dim SQLstring As String
            cn.Open(CNsfmdb)

            '88888888888888888888888888888888 对应图号工艺列表 88888888888888888888888888888888888888888888888888888888888888888888888888
            SQLstring = "Select ProcMaker ,ProcDate, SFOrder,DWGInfo,CustDWG,DrawNo,Qty,CustCode From ProcCard where DrawNo like N'%" & TextBox_参考图号.Text & "%' group by ProcMaker,ProcDate,SFOrder,DWGInfo,CustDWG,DrawNo,Qty,CustCode order by DWGInfo,CustDWG,DrawNo,Qty,CustCode"
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
            cn.Close()
            '88888888888888888888888888 end of 对应图号工艺列表结束 8888888888888888888888888888888888888888888888888888888888888888888888
            rs = Nothing
            cn = Nothing
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

    Private Sub TextBox生产单号_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_生产单号.KeyDown
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
            TextBox_DWGInfo.Clear()

            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            rs.MoveFirst()
            TextBox_CustCode.Text = rs.Fields("CustCode").Value
            TextBox3.Text = rs.Fields("CustCode").Value + TextBox_已选图号.Text
            TextBox_DWGInfo.Text = rs.Fields("CustCode").Value + TextBox_已选图号.Text + TextBox_生产单号.Text
            rs.Close()

            '********************************** combobox2onnection *****************************************************************
            ComboBox_未写工艺图号.Text = ""
            ComboBox_未写工艺图号.Items.Clear()
            SQLstring = "Select DrawNo From SFOrderBase where ((SFOrder=" & "'" & TextBox_生产单号.Text & "'" & ") and (DrawNo not in (Select DrawNo from SFAll.dbo.ProcCard where SFOrder=" & "'" & TextBox_生产单号.Text & "')))"
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                'MsgBox("此生产单所有零件均已完成工艺！")
            Else
                rs.MoveFirst()
                Do While Not rs.EOF
                    ComboBox_未写工艺图号.Items.Add(rs("DrawNo").Value)
                    rs.MoveNext()
                Loop
            End If
            rs.Close()
            '*************************** end of combobox2 connection *****************************************************************

            ComboBox_已写工艺图号.Text = ""
            ComboBox_已写工艺图号.Items.Clear()
            SQLstring = "Select DrawNo From SFOrderBase where ((SFOrder=" & "'" & TextBox_生产单号.Text & "'" & ") and (DrawNo in (Select DrawNo from SFAll.dbo.ProcCard where SFOrder=" & "'" & TextBox_生产单号.Text & "')))"
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                'MsgBox("此生产单已完成工艺的图号记录数为0！")
            Else
                rs.MoveFirst()
                Do While Not rs.EOF
                    ComboBox_已写工艺图号.Items.Add(rs("DrawNo").Value)
                    rs.MoveNext()
                Loop
            End If

            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing

        End If
    End Sub

    Private Sub DataGridView2_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView2.DataError
        DataGridView2.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Red
    End Sub

    Private Sub ComboBox_已写工艺图号_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_已写工艺图号.SelectionChangeCommitted
        ComboBox_未写工艺图号.SelectedItem = Nothing

        If ComboBox_已写工艺图号.SelectedItem IsNot Nothing Then
            TextBox_已选图号.Text = ComboBox_已写工艺图号.SelectedItem.ToString '已选图号
        Else
            If ComboBox_未写工艺图号.SelectedItem IsNot Nothing Then
                TextBox_已选图号.Text = ComboBox_未写工艺图号.SelectedItem.ToString '已选图号
            End If
        End If

        TextBox3.Text = TextBox_CustCode.Text + TextBox_已选图号.Text
        TextBox_DWGInfo.Text = TextBox_CustCode.Text + TextBox_已选图号.Text + TextBox_生产单号.Text

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        'Dim SQLstring As String
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select * From SFOrderBase where ((SFOrder=" & "'" & TextBox_生产单号.Text & "'" & ") and (DrawNo=N'" & TextBox_已选图号.Text & "'" & "))"
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        TextBox6.Text = rs.Fields("Qty").Value
        TextBox7.Text = rs.Fields("Qty").Value
        '*************************** end of combobox2 connection *****************************************************************
        rs.Close()

        '显示当前工件的单件报价
        SQLstring = "Select UnitP from SFOrderBase where SFOrder=" & "'" & TextBox_生产单号.Text & "' and DrawNo=N'" + TextBox_已选图号.Text + "'"
        rs.Open(SQLstring, cn, 1, 1)
        Dim price As String
        price = rs.Fields(0).Value.ToString()
        Dim str As String
        If String.IsNullOrEmpty(price) Then
            str = "无报价"
        Else
            str = price
        End If
        Label_Price.Text = "当前单件报价￥：" + str
        rs.Close()

        '显示当前已写的工艺列表
        SQLstring = "Select * From ProcCard where DWGInfo=N'" & TextBox_DWGInfo.Text & "' order by ProcSN"
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)

        If rs.RecordCount = 0 Then
            MsgBox($"DWGInfo= {TextBox_DWGInfo.Text} 不存在！工艺卡中没有此记录数据！")
            Return
        End If

        DataGridView2.Rows.Clear()
        DataGridView2.Rows.Add(rs.RecordCount)

        Dim DataRow As Integer
        DataRow = 0
        Do While rs.EOF = False
            DataGridView2.Item("工序号", DataRow).Value = rs.Fields("ProcSN").Value
            DataGridView2.Item("资源名称", DataRow).Value = rs.Fields("ProcName").Value
            DataGridView2.Item("工序内容和注意事项", DataRow).Value = rs.Fields("ProcDesc").Value
            DataGridView2.Item("准备工时h", DataRow).Value = rs.Fields("PreTime").Value
            DataGridView2.Item("单件工时h", DataRow).Value = rs.Fields("UnitTime").Value
            DataGridView2.Item("该工序预计单件成本", DataRow).Value = rs.Fields("该工序预计单件成本").Value
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop

        rs.Close()
        '*************************** end of combobox2 connection *****************************************************************

        '88888888888888888888888888888888 对应图号工艺列表 88888888888888888888888888888888888888888888888888888888888888888888888888
        SQLstring = "Select ProcMaker ,ProcDate, SFOrder,DWGInfo,CustDWG,DrawNo,Qty,CustCode From ProcCard where DrawNo=N'" & TextBox_已选图号.Text & "' group by ProcMaker,ProcDate,SFOrder,DWGInfo,CustDWG,DrawNo,Qty,CustCode order by DWGInfo,CustDWG,DrawNo,Qty,CustCode"
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
        DataGridView1.Rows.Add(rs.RecordCount)

        DataRow = 0
        Do While rs.EOF = False
            For i = 1 To rs.Fields.Count
                DataGridView1.Item(i - 1, DataRow).Value = rs.Fields(i - 1).Value
            Next i
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop

        rs.Close()
        cn.Close()
        '88888888888888888888888888 end of 对应图号工艺列表结束 8888888888888888888888888888888888888888888888888888888888888888888888
        rs = Nothing
        cn = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)
    End Sub

    Private Sub ComboBox_未写工艺图号_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox_未写工艺图号.SelectionChangeCommitted
        DataGridView2.Rows.Clear()
        ComboBox_已写工艺图号.SelectedItem = Nothing

        If ComboBox_已写工艺图号.SelectedItem IsNot Nothing Then
            TextBox_已选图号.Text = ComboBox_已写工艺图号.SelectedItem.ToString '已选图号
        Else
            If ComboBox_未写工艺图号.SelectedItem IsNot Nothing Then
                TextBox_已选图号.Text = ComboBox_未写工艺图号.SelectedItem.ToString '已选图号
            End If
        End If

        TextBox3.Text = TextBox_CustCode.Text + TextBox_已选图号.Text
        TextBox_DWGInfo.Text = TextBox_CustCode.Text + TextBox_已选图号.Text + TextBox_生产单号.Text

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        'Dim SQLstring As String
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select * From SFOrderBase where ((SFOrder=" & "'" & TextBox_生产单号.Text & "'" & ") and (DrawNo=N'" & TextBox_已选图号.Text & "'" & "))"
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        TextBox6.Text = rs.Fields("Qty").Value
        TextBox7.Text = rs.Fields("Qty").Value
        '*************************** end of combobox2 connection *****************************************************************
        rs.Close()

        '显示当前工件的单件报价
        SQLstring = "Select UnitP from SFOrderBase where SFOrder=" & "'" & TextBox_生产单号.Text & "' and DrawNo=N'" + TextBox_已选图号.Text + "'"
        rs.Open(SQLstring, cn, 1, 1)
        Dim price As String
        price = rs.Fields(0).Value.ToString()
        Dim str As String
        If String.IsNullOrEmpty(price) Then
            str = "无报价"
        Else
            str = price
        End If
        Label_Price.Text = "当前单件报价￥：" + str
        rs.Close()

        '88888888888888888888888888888888 对应图号工艺列表 88888888888888888888888888888888888888888888888888888888888888888888888888
        SQLstring = "Select ProcMaker ,ProcDate, SFOrder,DWGInfo,CustDWG,DrawNo,Qty,CustCode From ProcCard where DrawNo=N'" & TextBox_已选图号.Text & "' group by ProcMaker,ProcDate,SFOrder,DWGInfo,CustDWG,DrawNo,Qty,CustCode order by DWGInfo,CustDWG,DrawNo,Qty,CustCode"
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
        cn.Close()
        '88888888888888888888888888 end of 对应图号工艺列表结束 8888888888888888888888888888888888888888888888888888888888888888888888
        rs = Nothing
        cn = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)
    End Sub
End Class