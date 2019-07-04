Public Class Form订单查询
    Public QueryText, QueryCase As String
    Private Sub Form4_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Select Case UserType
            Case "admin"
                订单状态查询ToolStripMenuItem.Visible = True
                未完成工艺卡订单查询ToolStripMenuItem.Visible = True
                扫描记录ToolStripMenuItem.Visible = True
            Case "sales"
                订单状态查询ToolStripMenuItem.Visible = True
                未完成工艺卡订单查询ToolStripMenuItem.Visible = True
            Case "process"
                订单状态查询ToolStripMenuItem.Visible = False
                未完成工艺卡订单查询ToolStripMenuItem.Visible = True
            Case "plan"
                订单状态查询ToolStripMenuItem.Visible = True
                未完成工艺卡订单查询ToolStripMenuItem.Visible = True
            Case "manager"
                订单状态查询ToolStripMenuItem.Visible = True
                未完成工艺卡订单查询ToolStripMenuItem.Visible = True
            Case ""
                订单状态查询ToolStripMenuItem.Visible = True
                '未完成工艺卡订单查询ToolStripMenuItem.Visible = True
        End Select
    End Sub

    Private Sub 所有订单查询ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'shell(SFOrderBaseNetConnect, vbHide)
        GroupBox1.Visible = False
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        'Dim SQLstring As String
        Dim SQLstring, CNString As String

        On Error Resume Next
        SQLstring = "Select * from SFOrderBase"
        DataGridView1.DataSource = Nothing
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()

        cn.Open(CNsfmdb)
        'rs.Open(SQLstring, cn, 1, 2)
        rs.Open(SQLstring, cn, 1, 1)
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
        cn.Close()
        rs = Nothing
        cn = Nothing
        Label1.Text = "所有生产单："
        'shell(SFOrderBaseNetDisconnect, vbHide)
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub 订单状态查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 订单状态查询ToolStripMenuItem.Click
        GroupBox1.Visible = False
        Label1.Text = "按生产单号查询零件状况"
        Label2.Text = "请输入生产单号："
        TextBox1.Clear()
        TextBox1.Focus()
        QueryCase = "ByOrderBase"
        Label1.Visible = True
        Label2.Visible = True
        TextBox1.Visible = True
        DataGridView1.DataSource = Nothing
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        Label1.Text = "生产单整体状态查询："
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
    End Sub

    Private Sub 未完成工艺卡订单查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 未完成工艺卡订单查询ToolStripMenuItem.Click
        Dim da As New System.Data.OleDb.OleDbDataAdapter
        Dim ds As New DataSet
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        GroupBox1.Visible = False
        TextBox1.Visible = False
        Label1.Visible = False
        Label2.Visible = False

        QueryText = "Select SFOrder '生产单号',DrawNo '图号',DWGInfo '图号信息',Format(CDate, 'yyyy/MM/dd')  '下单日期',Format(DDate, 'yyyy/MM/dd') '交货日期' from SFOrderbase where DWGInfo not in (select DWGInfo from SFAll.dbo.ProcCard) and Status<>'完成入库'"
        TextBox1.Clear()
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset

        Dim SQLstring, CNString As String
        On Error Resume Next
        Dim CustCode As String
        Dim DDateCol As Integer

        cn.Open(CNsfmdb)
        rs.Open(QueryText, cn, 1, 1)
        If (rs.RecordCount = 0) Then
            MsgBox("没有未完成的工艺卡！")
            Shell(SFOrderBaseNetDisconnect, vbHide)
            Exit Sub
        End If

        da.Fill(ds, rs, "us")
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Refresh()
        cn.Close()
        rs = Nothing
        cn = Nothing
        da = Nothing
        ds = Nothing

        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
        For i = 0 To DataGridView1.RowCount - 1
            If (Val(DataGridView1.Item("交货日期", i).Value) < (Now.Year * 10000 + Now.Month * 100 + Now.Day)) Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
            End If
        Next
        Label1.Text = "未完成工艺卡清单："
        Label1.Visible = True
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim da As New System.Data.OleDb.OleDbDataAdapter
        Dim ds As New DataSet

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset

        If (((DataGridView1.Columns(e.ColumnIndex).Name = "DWGInfo") Or (DataGridView1.Columns(e.ColumnIndex).Name = "图号信息")) And (e.RowIndex >= 0)) Then
            '888888888888888888888888 display the selected process card 888888888888888888888888
            'Dim SQLstring As String
            Dim SQLstring, CNString As String
            'On Error Resume Next
            Dim DWGInfo As String
            DWGInfo = DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
            'SQLstring = "Select DrawNo, SFOrder,CustCode CustDWG, ProcSN, ProcName, ProcDesc, Qty, ProcQty from ProcCard"
            SQLstring = "Select * from ProdRecord where DWGInfo=" & "'" & DWGInfo & "' order by ProcCode"
            Form图纸信息工艺卡.DataGridView1.DataSource = Nothing
            Form图纸信息工艺卡.DataGridView1.Columns.Clear()
            Form图纸信息工艺卡.DataGridView1.Rows.Clear()
            ProcessForm = "form4"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then '未编工艺则退出
                rs.Clone()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            End If

            da.Fill(ds, rs, "us")
            Form图纸信息工艺卡.DataGridView1.DataSource = ds.Tables(0)
            Form图纸信息工艺卡.DataGridView1.Refresh()
            cn.Close()
            rs = Nothing
            cn = Nothing
            da = Nothing
            ds = Nothing
            '8888888888888888888888888 form10 datagridview1 title 8888888888888888888888888888

            Form图纸信息工艺卡.Text = DWGInfo + "  工艺卡"
            Me.Hide()
            Form图纸信息工艺卡.Show()

            Exit Sub
            '88888888888888888888 end of display the selected process card 888888888888888888888
        End If
        If (((DataGridView1.Columns(e.ColumnIndex).Name = "DrawNo") Or (DataGridView1.Columns(e.ColumnIndex).Name = "图号")) And (e.RowIndex >= 0)) Then
            Shell(SFOrderBaseNetConnect, vbHide)
            '88888888888888888888888888 Select DrawNo to display 88888888888888888888888888888
            Dim sF As String
            Dim Path, SearchedFile As String
            Path = DWGPath
            If (DataGridView1.Columns(e.ColumnIndex).Name = "DrawNo") Then
                SearchedFile = DataGridView1.Item("DrawNo", e.RowIndex).Value
            Else
                SearchedFile = DataGridView1.Item("图号", e.RowIndex).Value
            End If
            sF = Dir(Path, vbDirectory) ' 查找目录中第一个文件夹名称
            Do While sF <> "" ' 跳过当前的目录及上层目录
                If (InStr(1, UCase(sF), SearchedFile) > 0) Then
                    System.Diagnostics.Process.Start(Path + sF)
                    Exit Do
                End If
                sF = Dir()
            Loop
            Exit Sub
            '88888888888888888888 end of Select DrawNo to display 8888888888888888888888888888
        End If
    End Sub

    Private Sub 按生产单号查询零件生状况ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 按生产单号查询零件生状况ToolStripMenuItem.Click
        GroupBox1.Visible = False
        Label2.Text = "请输入生产单号："
        QueryText = ""
        TextBox1.Clear()
        TextBox1.Focus()
        QueryCase = "ByOrder"
        Label1.Visible = True
        Label2.Visible = True
        TextBox1.Visible = True
        DataGridView1.DataSource = Nothing
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        Label1.Text = "按生产单查询零件生产状况："
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = 13) Then
            Dim da As New System.Data.OleDb.OleDbDataAdapter
            Dim ds As New DataSet
            DataGridView1.DataSource = Nothing
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            If (TextBox1.Text = "") Then
                MsgBox("请输入查询内容！")
                StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                'shell(SFOrderBaseNetDisconnect, vbHide)
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset

            cn.Open(CNsfmdb)
            'On Error Resume Next
            Select Case QueryCase
                Case "ByOrder"
                    QueryText = "Select DWGInfo '图号信息',DrawNo'图号',OKQty '合格数量',ScrapQty '报废数量',Qty '订单数量',ProcQty '工艺数量',Status '状态',ProcCode '工序代码',convert(varchar(10), DDate,112) '交货日期' from SFOrderQuery where SFOrder=" & "'" & TextBox1.Text & "'"
                    DataGridView1.Columns.Clear()
                    DataGridView1.Rows.Clear()

                    'rs.Open(QueryText, cn, 1, 2)
                    rs.Open(QueryText, cn, 1, 1)
                    If (rs.RecordCount = 0) Then
                        MsgBox("系统中找不到该生产单信息！")
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                        'shell(SFOrderBaseNetDisconnect, vbHide)
                        Exit Sub
                    End If

                    da.Fill(ds, rs, "us")
                    DataGridView1.DataSource = ds.Tables(0)
                    DataGridView1.Refresh()
                    cn.Close()
                    rs = Nothing
                    cn = Nothing
                    da = Nothing
                    ds = Nothing

                    For i = 0 To DataGridView1.RowCount - 1
                        If (Val(DataGridView1.Item("交货日期", i).Value) < (Now.Year * 10000 + Now.Month * 100 + Now.Day)) Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                        End If
                        If (IIf(IsDBNull(DataGridView1.Item("状态", i).Value) = True, "", DataGridView1.Item("状态", i).Value) = "完成入库") Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Green
                        End If
                    Next
                    DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                    Exit Sub
                Case "ByDWG"
                    'QueryText = "Select SFOrder,CustDWG,ProcCode,OKQty,ScrpQty,ProcQty,RecordTime,Status from SFOrderQuery where SFOrder=" & "'" & TextBox1.Text & "'" & " order by ProcCode"

                    'QueryText = "Select * from DWGQuery where DrawNo like " & "'%" & TextBox1.Text & "%'"
                    QueryText = "Select DrawNo'图号',DWGInfo '图号信息',CustDWG '客户图号',ProcCode '工序代码',Operator '操作者',OKQty '合格数量',ScrapQty '报废数量',ProcQty '工艺数量',RecordTime '记录时间',Status '状态' from DWGQuery where DrawNo like " & "'%" & TextBox1.Text & "%' order by RecordTime"
                    DataGridView1.Columns.Clear()
                    DataGridView1.Rows.Clear()
                    'rs.Open(QueryText, cn, 1, 2)
                    rs.Open(QueryText, cn, 1, 1)
                    If (rs.RecordCount = 0) Then
                        MsgBox("系统中找不到该图号信息！")
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                        'shell(SFOrderBaseNetDisconnect, vbHide)
                        Exit Sub
                    End If

                    da.Fill(ds, rs, "us")
                    DataGridView1.DataSource = ds.Tables(0)
                    DataGridView1.Refresh()
                    cn.Close()
                    rs = Nothing
                    cn = Nothing
                    da = Nothing
                    ds = Nothing

                    DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                    Exit Sub
                Case "ByOrderBase"
                    QueryText = "select SFOrder '赛峰单号',DWGInfo '图号信息',DrawNo '图号',CustCode '客户代码',CustOrder '客户单号',PartName '零件名称',Ver '版本',Qty '数量',Status '状态',Creator '下单者',convert(varchar(10), CDate,112) '下单日期',convert(varchar(10), DDate,112) '交货日期',OrderType '订单类型',Comment '备注',MaterialSource '供料来源' from SFOrderBase where SFOrder=" & "'" & TextBox1.Text & "'"
                    'QueryText = "select DWGInfo '图号信息',DrawNo '图号',CustCode '客户代码',CustOrder '客户单号',PartName '零件名称',Ver '版本',Qty '数量',Status '状态',Creator '下单者',CDate '下单日期',DDate '交货日期',OrderType '订单类型',Comment '备注',MaterialSource '供料来源' from SFOrderBase where SFOrder=" & "'" & TextBox1.Text & "'"
                    DataGridView1.Columns.Clear()
                    DataGridView1.Rows.Clear()
                    rs.Open(QueryText, cn, 1, 1)
                    If (rs.RecordCount = 0) Then
                        MsgBox("系统中找不到该生产单信息！")
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                        'shell(SFOrderBaseNetDisconnect, vbHide)
                        Exit Sub
                    End If

                    da.Fill(ds, rs, "us")
                    DataGridView1.DataSource = ds.Tables(0)
                    DataGridView1.Refresh()
                    cn.Close()
                    rs = Nothing
                    cn = Nothing
                    da = Nothing
                    ds = Nothing

                    For i = 0 To DataGridView1.RowCount - 1
                        If (Val(DataGridView1.Item("交货日期", i).Value) < (Now.Year * 10000 + Now.Month * 100 + Now.Day)) Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                        End If
                        If (IIf(IsDBNull(DataGridView1.Item("状态", i).Value) = True, "", DataGridView1.Item("状态", i).Value) = "完成入库") Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Green
                        End If
                    Next
                    DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                    Exit Sub
                Case "ByCustCode"
                    QueryText = "Select DWGInfo '图号信息',DrawNo '图号',SFOrder '赛峰单号',Qty '数量',Status '状态',convert(varchar(10), CDate,112) '下单日期',convert(varchar(10), DDate,112) '交货日期',OrderType '订单类型' from SFOrderBase where CustCode=" & "'" & TextBox1.Text & "'"
                    DataGridView1.Columns.Clear()
                    DataGridView1.Rows.Clear()
                    rs.Open(QueryText, cn, 1, 1)
                    StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count

                    If (rs.RecordCount = 0) Then
                        MsgBox("系统中找不到该客户图号信息！")
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                        'shell(SFOrderBaseNetDisconnect, vbHide)
                        Exit Sub
                    End If

                    da.Fill(ds, rs, "us")
                    DataGridView1.DataSource = ds.Tables(0)
                    DataGridView1.Refresh()
                    cn.Close()
                    rs = Nothing
                    cn = Nothing
                    da = Nothing
                    ds = Nothing

                    For i = 0 To DataGridView1.RowCount - 1
                        If (Val(DataGridView1.Item("交货日期", i).Value) < (Now.Year * 10000 + Now.Month * 100 + Now.Day)) Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                        End If
                        If (IIf(IsDBNull(DataGridView1.Item("状态", i).Value) = True, "", DataGridView1.Item("状态", i).Value) = "完成入库") Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Green
                        End If
                    Next
                    DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                    Exit Sub
                Case "ByDrawNo"
                    QueryText = "Select DrawNo '图号',DWGInfo '图号信息',SFOrder '赛峰单号',Creator '下单者',convert(varchar(10), CDate,112) '下单日期',convert(varchar(10), DDate,112) '交货日期',Qty '订单数量',Status '状态' from SFOrderBase where DrawNo like " & "'%" & TextBox1.Text & "%'"
                    rs.Open(QueryText, cn, 1, 1)
                    If (rs.RecordCount = 0) Then
                        MsgBox("系统中找不到该图号信息！")
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                        Exit Sub
                    End If

                    da.Fill(ds, rs, "us")
                    DataGridView1.DataSource = ds.Tables(0)
                    DataGridView1.Refresh()
                    cn.Close()
                    rs = Nothing
                    cn = Nothing
                    da = Nothing
                    ds = Nothing

                    For i = 0 To DataGridView1.RowCount - 1
                        If (Val(DataGridView1.Item("交货日期", i).Value) < (Now.Year * 10000 + Now.Month * 100 + Now.Day)) Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                        End If
                        If (IIf(IsDBNull(DataGridView1.Item("状态", i).Value) = True, "", DataGridView1.Item("状态", i).Value) = "完成入库") Then
                            DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Green
                        End If
                    Next
                    DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                    StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
                    Exit Sub
            End Select
            StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
        End If
    End Sub
    Private Sub 按客户查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 按客户查询ToolStripMenuItem.Click
        GroupBox1.Visible = False
        Label1.Text = "按客户查询生产进程"
        Label2.Text = "请输入客户代码："
        QueryText = ""
        TextBox1.Clear()
        TextBox1.Focus()
        QueryCase = "ByCustCode"
        Label1.Visible = True
        Label2.Visible = True
        TextBox1.Visible = True
        DataGridView1.DataSource = Nothing
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
    End Sub

    Private Sub 扫描记录ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 扫描记录ToolStripMenuItem.Click
        Dim da As New System.Data.OleDb.OleDbDataAdapter
        Dim ds As New DataSet
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        GroupBox1.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        TextBox1.Visible = False

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String

        SQLstring = "Select top 500 ID as 序号,DWGInfo as 图号信息,Operator as 操作者,OKQty as 合格数量,ScrapQty as 报废数量,Qty as 数量,Format(RecordTime, 'yyyy/MM/dd HH:mm')  as 记录时间,RecordType as 记录类型,Status as 状态,WorkTime as 工时 from ProcRecord order by ID desc"

        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)

        da.Fill(ds, rs, "us")
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Refresh()
        cn.Close()
        rs = Nothing
        cn = Nothing
        da = Nothing
        ds = Nothing

        Label1.Text = "最近500条扫描记录："
        Label1.Visible = True
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
    End Sub
    Private Sub 图号信息查查ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 图号信息查看ToolStripMenuItem.Click
        GroupBox1.Visible = False
        Label1.Visible = True
        Label2.Visible = True
        TextBox1.Visible = True
        Label1.Text = "根据图号查找订单信息"
        Label2.Text = "请输入图号："
        QueryText = ""
        TextBox1.Clear()
        TextBox1.Focus()
        QueryCase = "ByDrawNo"
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
    End Sub

    Private Sub 工序已收料查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 工序已收料查询ToolStripMenuItem.Click
        Dim da As New System.Data.OleDb.OleDbDataAdapter
        Dim ds As New DataSet
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        QueryCase = "QuerybyReceiving"
        Label3.Text = "请选择收料工序："
        GroupBox1.Visible = True
        Label1.Visible = False
        Label2.Visible = False
        TextBox1.Visible = False
        '********************************** combobox1connection *****************************************************************
        ComboBox1.Visible = True
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * from ReceiveList order by '序号'"

        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)

        rs.MoveFirst()
        ComboBox1.Items.Clear()
        Do While Not rs.EOF
            ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        Label1.Text = "工序已收料查询："
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
        '************************* end of t
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim da As New System.Data.OleDb.OleDbDataAdapter
        Dim ds As New DataSet
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        Dim SearChText As String
        If (QueryCase = "QuerybyReceiving") Then
            SearChText = ComboBox1.SelectedItem.ToString & "-收料"
            SQLstring = "Select DWGInfo '图号信息', DrawNo '图号', Qty '订单数量', Status '状态', PartName '零件名称', Format(RecordTime, 'yyyy/MM/dd HH:mm')  '记录时间', Operator '操作者' from ReceivingList where Status like " & "'%" & SearChText & "%' order by RecordTime desc"
        Else
            SearChText = ComboBox1.SelectedItem.ToString & "-工序完成"
            SQLstring = "Select DWGInfo '图号信息', DrawNo '图号', Qty '订单数量', Status '状态', PartName '零件名称', Format(RecordTime, 'yyyy/MM/dd HH:mm') '记录时间', Operator '操作者' from CompletedList where Status like " & "'%" & SearChText & "%' order by RecordTime desc"
        End If

        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)

        da.Fill(ds, rs, "us")
        DataGridView1.DataSource = ds.Tables(0)
        DataGridView1.Refresh()
        cn.Close()
        rs = Nothing
        cn = Nothing
        da = Nothing
        ds = Nothing

        DataGridView1.Focus()
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
        Label1.Text = ComboBox1.SelectedItem.ToString & " 工序收料记录："
    End Sub
    Private Sub 工序已完成查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 工序已完成查询ToolStripMenuItem.Click
        Dim da As New System.Data.OleDb.OleDbDataAdapter
        Dim ds As New DataSet
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        QueryCase = "QuerybyCompleted"
        Label3.Text = "请选择完成工序："
        DataGridView1.Rows.Clear()
        GroupBox1.Visible = True
        Label1.Visible = False
        Label2.Visible = False
        TextBox1.Visible = False
        '********************************** combobox1connection *****************************************************************
        ComboBox1.Visible = True
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * from StatusList order by 'SN'"

        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        ComboBox1.Items.Clear()
        Do While Not rs.EOF
            ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        Label1.Text = "工序已完成查询："
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
        '************************* end of t
    End Sub
    Private Sub 按单号查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 按单号查询ToolStripMenuItem.Click
        Me.Hide()
        Form按生产单号查询未完成零件状况.Show()
        Form按生产单号查询未完成零件状况.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 按客户查询ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 按客户查询ToolStripMenuItem1.Click
        Me.Hide()
        Form按客户查询未完成零件状况.Show()
        Form按客户查询未完成零件状况.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 按图号查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 按图号查询ToolStripMenuItem.Click
        Me.Hide()
        Form按图号查询未完成零件.Show()
        Form按图号查询未完成零件.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub 外发记录查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 外发记录查询ToolStripMenuItem.Click
        SubReordQuery = "NormalQuery"
        Me.Hide()
        Form外发记录查询.Show()
    End Sub

    Private Sub 所有未完成查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 所有未完成查询ToolStripMenuItem.Click
        Me.Hide()
        Form所有未完成生产单.Show()
        Form所有未完成生产单.WindowState = FormWindowState.Maximized
    End Sub
    Private Sub 非嘉睦未完成查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 非嘉睦未完成查询ToolStripMenuItem.Click
        Me.Hide()
        Form非嘉睦未完成订单查询.Show()
    End Sub

    Private Sub 按交期查询ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 按交期查询ToolStripMenuItem.Click
        Me.Hide()
        Form按客户订单号查询生产单.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error Resume Next
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("No drawing can be selected!")
            Exit Sub
        End If

        Dim DrawNo As String
        If Dir(TempDataPath, vbDirectory) = "" Then
            MkDir(TempDataPath)
        Else
            Kill(TempDataPath + "*.*")
        End If

        For i = 0 To DataGridView1.Rows.Count - 1
            DrawNo = TempDataPath & DataGridView1.Item("图号", i).Value & ".pdf"
            If (IO.File.Exists(DrawNo)) Then
                IO.File.Delete(DrawNo)
            End If
            If (IO.File.Exists(DWGPath & DataGridView1.Item("图号", i).Value & ".pdf")) Then
                IO.File.Copy(DWGPath & DataGridView1.Item("图号", i).Value & ".pdf", DrawNo)
            Else
                DataGridView1.Item("版本", i).Value = "缺图"
            End If
        Next
    End Sub
End Class