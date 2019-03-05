Public Class Form工艺流程卡处理
    Public ProcDate
    Private Sub Form2_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.Title = "打开工艺流程卡文件"
        OpenFileDialog1.ShowDialog()
        If (OpenFileDialog1.FileName = "OpenFileDialog1") Then
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        Dim TempFileName As String
        Dim SFOrder, DrawNo, CustCode As String
        Dim Qty, ProcQty, StartNo, EndNo As Integer
        Dim ProcCodeLenghth As Integer '条形码字符数
        ProcCodeLenghth = 0
        TempFileName = OpenFileDialog1.FileName



        'On Error GoTo ErrorOut
        Dim ss As New Microsoft.Office.Interop.Excel.Application
        Dim xlsheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlbook As Microsoft.Office.Interop.Excel.Workbook
        xlbook = ss.Workbooks.Open(TempFileName)
        xlsheet = xlbook.Sheets(1)

        '**************************** 将Excel流程卡内容写入到Datagridview1 ************************************
        If (xlsheet.Cells(2, 1).value = "PJ号") Then
            StartNo = 2
        Else
            If (xlsheet.Cells(3, 1).value = "PJ号") Then
                StartNo = 3
            Else
                MsgBox("工艺卡格式有问题！")
                Exit Sub
            End If
        End If

        SFOrder = xlsheet.Cells(StartNo, 2).value
        DrawNo = xlsheet.Cells(StartNo, 5).value
        ProcDate = xlsheet.Cells(StartNo, 14).value
        Qty = xlsheet.Cells(StartNo, 18).value
        ProcQty = xlsheet.Cells(StartNo, 20).value
        CustCode = xlsheet.Cells(StartNo, 22).value
        TextBox1.Text = DrawNo
        TextBox2.Text = CustCode
        TextBox3.Text = SFOrder
        TextBox4.Text = xlsheet.Cells(StartNo, 18).value
        TextBox5.Text = xlsheet.Cells(StartNo, 20).value
        TextBox6.Text = xlsheet.Cells(StartNo, 12).value
        StartNo = StartNo + 2
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()

        DataGridView1.Columns.Add("工序号", "工序号")
        DataGridView1.Columns.Add("资源名称", "资源名称")
        DataGridView1.Columns.Add("工序内容和注意事项", "工序内容和注意事项")
        DataGridView1.Columns.Add("准备工时", "准备工时")
        DataGridView1.Columns.Add("单件工时", "单件工时")

        For i = StartNo To 300
            If ((Len(xlsheet.Cells(i, 1).value) = 0) Or (Len(xlsheet.Cells(i, 2).value) = 0)) Then
                EndNo = i
                Exit For
            End If
        Next
        DataGridView1.Rows.Add(EndNo - StartNo)
        For i = 0 To (EndNo - StartNo - 1)
            DataGridView1.Item(0, i).Value = xlsheet.Cells(StartNo + i, 1).value
            DataGridView1.Item(1, i).Value = xlsheet.Cells(StartNo + i, 2).value
            DataGridView1.Item(2, i).Value = xlsheet.Cells(StartNo + i, 3).value
            DataGridView1.Item(3, i).Value = xlsheet.Cells(StartNo + i, 10).value
            DataGridView1.Item(4, i).Value = xlsheet.Cells(StartNo + i, 11).value
        Next
        xlbook.Close()
        ss.Application.Quit()
        xlsheet = Nothing
        xlbook = Nothing
        ss = Nothing
        '************************* end of 将Excel流程卡内容写入到Datagridview1 ********************************
        '88888888888888888888888888888 将已存在的数据库ProcCard的数据导入到Datagridview2 888888888888888888888888888
        'shell(SFOrderBaseNetConnect, vbHide)
        SQLstring = "Select * From ProcCard where ((SFOrder=" & "'" & SFOrder & "'" & ") and (DrawNo=" & "'" & DrawNo & "'" & "))"
        DataGridView2.Columns.Clear()
        DataGridView2.Rows.Clear()

        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 2)
        If (rs.RecordCount = 0) Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            'shell(SFOrderBaseNetDisconnect, vbHide)
            MsgBox("找不到对应的生产单或图号！")
            Exit Sub
        End If
        rs.MoveFirst()
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()

        For i = 1 To rs.Fields.Count
            DataGridView2.Columns.Add(rs.Fields(i - 1).Name, rs.Fields(i - 1).Name)
        Next
        DataGridView2.Rows.Add(rs.RecordCount)
        Dim DataRow As Integer
        DataRow = 0
        Do While rs.EOF = False
            For i = 1 To rs.Fields.Count
                DataGridView2.Item(i - 1, DataRow).Value = rs.Fields(i - 1).Value
            Next i
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop
        '8888888888888888888 end of 将已存在的数据库ProcCard的数据导入到Datagridview2 888888888888888888888888888
        '*****************************************************************************
        rs.Close()
        cn.Close()
        For Each p As Process In ProcessWatch
            p.Kill()
        Next
        rs = Nothing
        cn = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'shell(SFOrderBaseNetConnect, vbHide)
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        Dim DWGInfo, SFOrder, DrawNo, CustCode, CustDWG, ProcMaker As String
        SQLstring = "Select * From SFOrderBase where ((SFOrder=" & "'" & TextBox3.Text & "'" & ") and (DrawNo=" & "'" & TextBox1.Text & "'" & "))"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 2)
        If (rs.RecordCount = 0) Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            'shell(SFOrderBaseNetDisconnect, vbHide)
            MsgBox("图号或生产单号不对，请检查！")
            Exit Sub
        End If
        rs.MoveFirst()
        TextBox2.Text = rs.Fields("CustCode").Value
        ProcMaker = TextBox6.Text
        rs.Close()
        '888888888888888888888 删除已有的工艺卡数据并将Datagridview1的数据存入工艺卡数据库 888888888888888888888888888888888888888888888888888888888888
        SQLstring = "Select * From ProcCard where ((SFOrder=" & "'" & TextBox3.Text & "'" & ") and (DrawNo=" & "'" & TextBox1.Text & "'" & "))"
        rs.Open(SQLstring, cn, 1, 2)
        If (rs.RecordCount > 0) Then
            rs.MoveFirst()
            Do While rs.EOF = False
                rs.Delete()
                rs.Update()
                rs.MoveNext()
            Loop
        End If

        DWGInfo = TextBox2.Text & TextBox1.Text & TextBox3.Text
        SFOrder = TextBox3.Text
        DrawNo = TextBox1.Text
        CustDWG = TextBox2.Text & TextBox1.Text
        CustCode = TextBox2.Text
        For i = 0 To (DataGridView1.Rows.Count - 1)
            If ((Len(DataGridView1.Item(0, i).Value) = 0) Or (Len(DataGridView1.Item(1, i).Value) = 0)) Then
                Exit For
            Else
                rs.AddNew()
                rs.Fields("ProcSN").Value = DataGridView1.Item("工序号", i).Value
                rs.Fields("ProcName").Value = DataGridView1.Item("资源名称", i).Value
                rs.Fields("ProcDesc").Value = DataGridView1.Item("工序内容和注意事项", i).Value
                rs.Fields("PreTime").Value = DataGridView1.Item("准备工时", i).Value
                rs.Fields("UnitTime").Value = DataGridView1.Item("单件工时", i).Value
                rs.Fields("SFOrder").Value = SFOrder
                rs.Fields("DrawNo").Value = DrawNo
                rs.Fields("CustCode").Value = CustCode
                rs.Fields("CustDWG").Value = CustDWG
                rs.Fields("DWGInfo").Value = DWGInfo
                'ProcCode = SFOrder & DrawNo & DataGridView1.Item(0, i).Value
                rs.Fields("ProcMaker").Value = ProcMaker
                rs.Fields("Qty").Value = TextBox4.Text
                rs.Fields("ProcQty").Value = TextBox5.Text
                rs.Fields("ProcDate").Value = ProcDate
                rs.Fields("ProcCode").Value = TextBox3.Text & TextBox1.Text & DataGridView1.Item(0, i).Value
                rs.Update()
            End If
        Next
        rs.Close()
        '888888888888888888888 end of 删除已有的工艺卡数据并将Datagridview1的数据存入工艺卡数据库 88888888888888888888888888888888888888888888888888888
        '88888888888888888888888888 更新订单状态 888888888888888888888888888888888888888888888888888
        SQLstring = "Select * From SFOrderBase where ((SFOrder=" & "'" & TextBox3.Text & "'" & ") and (DrawNo=" & "'" & TextBox1.Text & "'" & "))"
        rs.Open(SQLstring, cn, 1, 2)
        rs.MoveFirst()
        rs.Fields("Status").Value = "完成工艺"
        rs.Update()
        '8888888888888888888888888 更新订单状态完成 888888888888888888888888888888888888888888888888
        cn.Close()
        rs = Nothing
        cn = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)
        MsgBox("成功导入Excel工艺卡数据！")
    End Sub
End Class