Public Class Form嘉睦订单录入
    Public ss As New Microsoft.Office.Interop.Excel.Application
    Public xlsheet As Microsoft.Office.Interop.Excel.Worksheet
    Public xlbook As Microsoft.Office.Interop.Excel.Workbook
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ss = CreateObject("Excel.Application")
        xlbook = ss.Workbooks.Add
        If xlbook.Sheets.Count > 1 Then
            For i = xlbook.Sheets.Count To 2
                xlbook.Sheets(i).delete()
            Next
        End If

        ss.Visible = True
        xlsheet = xlbook.Sheets(1)
        xlsheet.Name = "jamu"
        xlsheet.Cells(1, 1) = "下单日期"
        xlsheet.Cells(1, 2) = "图号"
        xlsheet.Cells(1, 3) = "数量"
        xlsheet.Cells(1, 4) = "单价"
        xlsheet.Cells(1, 5) = "金额"
        xlsheet.Cells(1, 6) = "客户订单号"
        xlsheet.Cells(1, 7) = "交期"
        'xlsheet.Cells(1, 8) = "记录批号"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ID, BatchID As Integer

        If (xlsheet.UsedRange.Rows.Count < 2) Then
            MsgBox("没有订单内容，请检查！")
            Exit Sub
        End If
        For i = 2 To xlsheet.UsedRange.Rows.Count
            For j = 3 To 5
                If (IsDBNull(xlsheet.Cells(i, j).value)) Then
                    MsgBox("数量/单价/金额有问题，请检查！")
                    Exit Sub
                Else
                    If (IsNumeric(xlsheet.Cells(i, j).value)) Then
                    Else
                        MsgBox("数量/单价/金额有问题，请检查！")
                        Exit Sub
                    End If
                End If
            Next
            If ((IsDate(xlsheet.Cells(i, 1).value)) And (IsDate(xlsheet.Cells(i, 7).value))) Then
            Else
                MsgBox("下单日期/出货日期有问题，请检查！")
                Exit Sub
            End If
        Next
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        'Dim SQLstring As String
        Dim SQLstring As String
        'On Error Resume Next
        SQLstring = "Select * From JamuOrder order by 序号"

        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 3)
        rs.MoveLast()
        ID = rs.Fields("序号").Value + 1
        BatchID = ID

        For i = 2 To xlsheet.UsedRange.Rows.Count
            rs.AddNew()
            rs.Fields("序号").Value = ID
            rs.Fields("下单日期").Value = xlsheet.Cells(i, 1).value
            rs.Fields("图号").Value = xlsheet.Cells(i, 2).value
            rs.Fields("数量").Value = xlsheet.Cells(i, 3).value
            rs.Fields("单价").Value = xlsheet.Cells(i, 4).value
            rs.Fields("金额").Value = xlsheet.Cells(i, 5).value
            rs.Fields("客户订单号").Value = xlsheet.Cells(i, 6).value
            rs.Fields("交期").Value = xlsheet.Cells(i, 7).value
            rs.Fields("记录批号").Value = Format(BatchID, "000000")
            rs.Update()
            ID = ID + 1
        Next
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
        MsgBox("嘉睦订单录入完成！")
    End Sub

    Private Sub Form52_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
End Class