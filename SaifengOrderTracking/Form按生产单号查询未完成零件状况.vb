Public Class Form按生产单号查询未完成零件状况

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = 13) Then
            Dim da As New System.Data.OleDb.OleDbDataAdapter
            Dim ds As New DataSet
            DataGridView1.DataSource = Nothing
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()

            If (TextBox1.Text = "") Then
                MsgBox("请输入生产单号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rs1 As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * from SFOrderBase where SFOrder=" & "'" & TextBox1.Text & "'"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有此生产单号，请检查！")
                Exit Sub
            End If
            rs.Close()
            'SQLstring = "Select * from SFOrderBase where (((status<>'完成入库') and (status<>'取消生产单')) and (SFOrder=" & "'" & TextBox1.Text & "'))"
            'SQLstring = "Select * from StatusInfo where ((Status not in (select Status where Status='完成入库' or Status='取消生产单' or Status='已终检')) and (SFOrder=" & "'" & TextBox1.Text & "'))"
            'SQLstring = "Select SFOrder '赛峰单号',DWGInfo '图号信息',DrawNo '图号',CustOrder '客户单号',Qty '数量',CDate '下单日期',DDate '交货日期',OrderType '订单类型',CustCode '客户代码',Status '状态',RecordTime '记录时间',FollowNote '跟单备注' from StatusInfo where not((CHARINDEX('完成入库',Status)>0) or (CHARINDEX('取消生产单',Status)>0) or (CHARINDEX('已终检',Status)>0)) and (SFOrder=" & "'" & TextBox1.Text & "')"
            SQLstring = "Select SFOrder '赛峰单号',DWGInfo '图号信息',DrawNo '图号',PartName '零件名称',CustOrder '客户单号',Qty '数量',CDate '下单日期',DDate '交货日期',OrderType '订单类型',CustCode '客户代码',Status '状态',RecordTime '记录时间',FollowNote '跟单备注' from StatusInfo where ((not((CHARINDEX('完成入库',Status)>0) or (CHARINDEX('取消生产单',Status)>0) or (CHARINDEX('已终检',Status)>0)))) and (SFOrder=" & "'" & TextBox1.Text & "')"
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此生产单号已完全入库，请检查！")
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
            StatusStrip1.Items(0).Text = "记录条数： " & DataGridView1.RowCount
        End If
    End Sub
    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        'shell(SFOrderBaseNetConnect, vbHide)
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
            'SQLstring = "Select * from ProcCard where DWGInfo=" & "'" & DWGInfo & "'"
            SQLstring = "Select * from ProdRecord where DWGInfo=" & "'" & DWGInfo & "' order by ProcCode"
            Form图纸信息工艺卡.DataGridView1.Columns.Clear()
            Form图纸信息工艺卡.DataGridView1.Rows.Clear()
            ProcessForm = "form16"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then '未编工艺则退出
                rs.Clone()
                cn.Close()
                rs = Nothing
                cn = Nothing
                'shell(SFOrderBaseNetDisconnect, vbHide)
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
    Private Sub Form16_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form订单查询.Show()
    End Sub
    Private Sub Form16_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = 2 Then '窗口最大化时的尺寸调整
            DataGridView1.Width = Me.Width - 100
            DataGridView1.Height = Me.Height - 200
        End If
    End Sub
End Class