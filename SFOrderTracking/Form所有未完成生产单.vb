Public Class Form所有未完成生产单
    Private Sub Form29_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form订单查询.Show()
    End Sub

    Private Sub Form29_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("DWGInfo", "图号信息")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("CustOrder", "客户订单号")
        DataGridView1.Columns.Add("Qty", "订单数量")
        DataGridView1.Columns.Add("CDate", "下单日期")
        DataGridView1.Columns.Add("DDate", "交货日期")
        DataGridView1.Columns.Add("OrderType", "订单种类")
        DataGridView1.Columns.Add("PartName", "零件名称")
        DataGridView1.Columns.Add("CustCode", "客户代码")
        DataGridView1.Columns.Add("Status", "状态")
        DataGridView1.Columns.Add("RecordTime", "记录时间")
        DataGridView1.Columns.Add("FollowNote", "跟单备注")
        DataGridView1.Rows.Clear()
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String

        SQLstring = "Select * from ReceiveList order by 序号"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        Do While rs.EOF = False
            ComboBox1.Items.Add(rs.Fields("收料位置").Value)
            rs.MoveNext()
        Loop
        ComboBox1.Items.Add("外发")
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        Me.WindowState = 2
    End Sub
    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        'shell(SFOrderBaseNetConnect, vbHide)
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        If ((DataGridView1.Columns(e.ColumnIndex).Name = "CustCode") And (e.RowIndex >= 0)) Then
            '88888888888888888888888888 Select custcode to display 88888888888888888888888888888

            'Dim SQLstring As String
            Dim SQLstring, CNString As String
            On Error Resume Next
            Dim CustCode As String
            CustCode = DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
            'SQLstring = "Select DrawNo, SFOrder,CustCode CustDWG, ProcSN, ProcName, ProcDesc, Qty, ProcQty from ProcCard"
            SQLstring = "Select * from SFOrderBase where (（isnull(Status,'') <>''）and (CustCode=" & "'" & CustCode & "'" & "))"
            'DataGridView1.Columns.Clear()
            DataGridView1.Rows.Clear()

            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            rs.MoveFirst()
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
            StatusStrip1.Items(0).Text = "未完成工艺卡记录条数： " & DataGridView1.Rows.Count
            cn = Nothing
            'shell(SFOrderBaseNetDisconnect, vbHide)
            Exit Sub
            '88888888888888888888 end of Select custcode to display 8888888888888888888888888888
        End If

        If (((DataGridView1.Columns(e.ColumnIndex).Name = "DWGInfo") Or (DataGridView1.Columns(e.ColumnIndex).Name = "图号信息")) And (e.RowIndex >= 0)) Then
            '888888888888888888888888 display the selected process card 888888888888888888888888
            'Dim SQLstring As String
            Dim SQLstring, CNString As String
            'On Error Resume Next
            Dim DWGInfo As String
            DWGInfo = DataGridView1.Item(e.ColumnIndex, e.RowIndex).Value
            SQLstring = "Select * from ProdRecord where DWGInfo=" & "'" & DWGInfo & "' order by ProcCode"
            Form图纸信息工艺卡.DataGridView1.Columns.Clear()
            Form图纸信息工艺卡.DataGridView1.Rows.Clear()
            ProcessForm = "form29"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then '未编工艺则退出
                rs.Clone()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            End If

            rs.MoveFirst()

            '8888888888888888888888888 form10 datagridview1 title 8888888888888888888888888888
            For i = 1 To rs.Fields.Count
                Form图纸信息工艺卡.DataGridView1.Columns.Add(rs.Fields(i - 1).Name, rs.Fields(i - 1).Name)
                If ((rs.Fields(i - 1).Name = "ID") Or (rs.Fields(i - 1).Name = "CustDWG") Or (rs.Fields(i - 1).Name = "SFOrder") Or (rs.Fields(i - 1).Name = "DrawNo") Or (rs.Fields(i - 1).Name = "PartName") Or (rs.Fields(i - 1).Name = "SPOrder") Or (rs.Fields(i - 1).Name = "ProcMaker") Or (rs.Fields(i - 1).Name = "ProcDate") Or (rs.Fields(i - 1).Name = "CustCode") Or (rs.Fields(i - 1).Name = "ProcCode")) Then
                    Form图纸信息工艺卡.DataGridView1.Columns(i - 1).Visible = False
                End If
            Next
            Form图纸信息工艺卡.Text = DWGInfo + "  工艺卡"
            '8888888888888888888888 end of form10 datagridview1 title 888888888888888888888888
            'StatusStrip1.Items(0).Text = "未完成工艺卡记录条数： " & rs.RecordCount
            Form图纸信息工艺卡.DataGridView1.Rows.Add(rs.RecordCount)
            'MsgBox("Seconds Records in ProcCard is " & rs.RecordCount)
            Dim DataRow As Integer
            DataRow = 0
            Do While rs.EOF = False
                For i = 1 To rs.Fields.Count
                    Form图纸信息工艺卡.DataGridView1.Item(i - 1, DataRow).Value = rs.Fields(i - 1).Value
                Next i
                DataRow = DataRow + 1
                rs.MoveNext()
            Loop

            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Me.Hide()
            Form图纸信息工艺卡.Show()
            'shell(SFOrderBaseNetDisconnect, vbHide)
            Exit Sub
            '88888888888888888888 end of display the selected process card 888888888888888888888
        End If
        'shell(SFOrderBaseNetDisconnect, vbHide)
        'If (((DataGridView1.Columns(e.ColumnIndex).Name = "DrawNo") Or (DataGridView1.Columns(e.ColumnIndex).Name = "图号")) And (e.RowIndex >= 0)) Then
        If ((DataGridView1.Columns(e.ColumnIndex).Name = "图号") And (e.RowIndex >= 0)) Then
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
            Shell(SFOrderBaseNetDisconnect, vbHide)
        End If
    End Sub

    Private Sub Form29_MaximizedBoundsChanged(sender As Object, e As EventArgs) Handles Me.MaximizedBoundsChanged

    End Sub

    Private Sub Form29_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = 2 Then '窗口最大化时的尺寸调整
            DataGridView1.Width = Me.Width - 100
            DataGridView1.Height = Me.Height - 100
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '查询所有未完成生产单
        Dim da As New System.Data.OleDb.OleDbDataAdapter
        Dim ds As New DataSet
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String

        SQLstring = "Select SFOrder '赛峰单号',DWGInfo '图号信息',DrawNo '图号',PartName '零件名称',CustOrder '客户单号',Qty '数量',CDate '下单日期',DDate '交货日期',OrderType '订单类型',CustCode '客户代码',Status '状态',RecordTime '记录时间',FollowNote '跟单备注' from StatusInfo where not((CHARINDEX('完成入库',Status)>0) or (CHARINDEX('取消生产单',Status)>0) or (CHARINDEX('已终检',Status)>0)) or isnull(Status,'')=''"
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

        For i = 0 To DataGridView1.RowCount - 1
            If (Val(DataGridView1.Item("交货日期", i).Value) < (Now.Year * 10000 + Now.Month * 100 + Now.Day)) Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
            End If
        Next
        Me.WindowState = 2
        StatusStrip1.Visible = True
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.Rows.Count
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

        SQLstring = "Select SFOrder '赛峰单号',DWGInfo '图号信息',DrawNo '图号',PartName '零件名称',CustOrder '客户单号',Qty '数量',CDate '下单日期',DDate '交货日期',OrderType '订单类型',CustCode '客户代码',Status '状态',RecordTime '记录时间',FollowNote '跟单备注' from StatusInfo where ((Status not in ('完成入库','取消生产单','已终检') or isnull(Status,'')='') and (Status like " & "'%" & ComboBox1.SelectedItem.ToString & "-%'" & "))"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            StatusStrip1.Items(0).Text = "记录条数： 0 "
            DataGridView1.Focus()
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
        Next
        Me.WindowState = 2
        StatusStrip1.Visible = True
        StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.RowCount
        DataGridView1.Focus()
    End Sub
End Class