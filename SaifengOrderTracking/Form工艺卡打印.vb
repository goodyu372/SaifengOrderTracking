Public Class Form工艺卡打印

    Private Sub Form8_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'shell(SFOrderBaseNetConnect, vbHide)
        If (TextBox1.Text = "") Then
            'shell(SFOrderBaseNetDisconnect, vbHide)
            MsgBox("请选择订单号！")
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            'shell(SFOrderBaseNetDisconnect, vbHide)
            MsgBox("请选择图号！")
            Exit Sub
        End If

        '************************************************** 读取工艺卡相关数据 *************************************************
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim RowNo As Integer
        RowNo = 5

        Dim SQLstring, CNString As String
        On Error Resume Next
        Dim CustCode As String
        Dim QueryText As String
        TextBox1.Text = ComboBox1.SelectedItem.ToString
        QueryText = "Select * from ProcCard where(( SFOrder=" & "'" & TextBox1.Text & "'" & ") and (DrawNo=" & "'" & ComboBox2.SelectedItem.ToString & "'" & ")) order by ProcSN"
        cn.Open(CNsfmdb)
        'rs.Open(QueryText, cn, 1, 2)
        rs.Open(QueryText, cn, 1, 1)
        If (rs.RecordCount = 0) Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            'shell(SFOrderBaseNetDisconnect, vbHide)
            MsgBox("该图号没有工艺卡！")
            Exit Sub
        End If

        '**************************************** Copy the Process card sample **********************************************************
        If Dir(TempDataPath, vbDirectory) = "" Then
            MkDir(TempDataPath)
        End If
        Dim PrintCardTemp As String
        PrintCardTemp = TempDataPath & TextBox1.Text & TextBox2.Text & ".xlsx"
        If (IO.File.Exists(PrintCardTemp)) Then
            IO.File.Delete(PrintCardTemp)
        End If
        IO.File.Copy(PrintCard, PrintCardTemp)
        '************************************** end of Copy the Process card sample *****************************************************
        Dim ss As New Microsoft.Office.Interop.Excel.Application
        Dim xlsheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim xlbook As Microsoft.Office.Interop.Excel.Workbook
        xlbook = ss.Workbooks.Open(PrintCardTemp)
        xlsheet = xlbook.Sheets(1)
        rs.MoveFirst()
        xlsheet.Cells(3, 2) = rs.Fields("SFOrder").Value
        xlsheet.Cells(3, 4) = rs.Fields("DrawNo").Value

        xlsheet.Cells(3, 13) = rs.Fields("ProcMaker").Value
        xlsheet.Cells(3, 18) = rs.Fields("Qty").Value
        xlsheet.Cells(3, 20) = rs.Fields("ProcQty").Value
        xlsheet.Cells(3, 22) = Mid(rs.Fields("DWGInfo").Value, 1, 4)
        Do While rs.EOF = False
            'xlsheet.Cells(RowNo, 1) = IIf(rs.Fields("ProcSN").Value >= 100, rs.Fields("ProcSN").Value, IIf(rs.Fields("ProcSN").Value < 10, "'00" & rs.Fields("ProcSN").Value, "'0" & rs.Fields("ProcSN").Value))
            xlsheet.Cells(RowNo, 1) = "'" & Val(rs.Fields("ProcSN").Value)  '     "'" & Format$(Val(rs.Fields("ProcSN").Value), "000")
            xlsheet.Cells(RowNo, 2) = rs.Fields("ProcName").Value
            xlsheet.Cells(RowNo, 3) = rs.Fields("ProcDesc").Value
            xlsheet.Cells(RowNo, 10) = rs.Fields("PreTime").Value
            xlsheet.Cells(RowNo, 11) = rs.Fields("UnitTime").Value
            xlsheet.Cells(RowNo, 12) = (rs.Fields("Qty").Value) * IIf(IsNumeric(rs.Fields("UnitTime").Value), rs.Fields("UnitTime").Value, 0) + IIf(IsNumeric(rs.Fields("PreTime").Value), rs.Fields("PreTime").Value, 0)
            RowNo = RowNo + 1
            rs.MoveNext()
        Loop
        rs.Close()
        QueryText = "Select * from SFOrderBase where(( SFOrder=" & "'" & ComboBox1.SelectedItem.ToString & "'" & ") and (DrawNo=" & "'" & ComboBox2.SelectedItem.ToString & "'" & "))"
        'rs.Open(QueryText, cn, 1, 2)
        rs.Open(QueryText, cn, 1, 1)
        rs.MoveFirst()
        xlsheet.Cells(41, 15) = rs.Fields("OrderType").Value
        xlsheet.Cells(1, 18) = rs.Fields("PartName").Value
        xlsheet.Cells(1, 9) = rs.Fields("CDate").Value
        xlsheet.Cells(1, 14) = rs.Fields("DDate").Value
        For i = 40 To RowNo Step -1
            xlsheet.Rows(i).Delete()
        Next
        xlsheet.Cells(1, 1).select()
        'shell(SFOrderBaseNetDisconnect, vbHide)
        MsgBox("已准备好打印工艺卡，请先打印预览！")
        ss.Visible = True
        '*********************************** end of 读取工艺卡相关数据   *******************************************************
        TextBox2.Clear()
        xlbook.Close()
        xlsheet = Nothing
        xlbook = Nothing
        ss = Nothing

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox2.Text = ComboBox2.SelectedItem.ToString
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            TextBox2.Clear()
            ComboBox2.Items.Clear()
            If TextBox1.Text = "" Then
                MsgBox("请输入生产单号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset

            Dim SQLstring, CNString As String
            On Error Resume Next
            Dim CustCode As String
            Dim QueryText As String
            QueryText = "Select DrawNo,SFOrder from SFOrderBase where SFOrder=" & "'" & TextBox1.Text & "'" & " order by DrawNo"
            cn.Open(CNsfmdb)
            rs.Open(QueryText, cn, 1, 1)
            If rs.RecordCount = 0 Then
                MsgBox("该生产单号不存在！")
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                ComboBox2.Items.Add(rs.Fields(0).Value)
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            ComboBox2.Focus()
        End If
    End Sub
End Class