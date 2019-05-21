Public Class Form嘉睦出货记录
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
        xlsheet.Name = "jamuOutput"
        'xlsheet.Cells(1, 1) = "下单日期"
        xlsheet.Cells(1, 1) = "图号"
        xlsheet.Cells(1, 2) = "数量"
        xlsheet.Cells(1, 3) = "单价"
        xlsheet.Cells(1, 4) = "总价"
        xlsheet.Cells(1, 5) = "客户订单号"
        xlsheet.Cells(1, 6) = "备注"
        xlsheet.Cells(1, 7) = "送货日期"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ID, BatchID As Integer
        On Error Resume Next

        If (xlsheet.UsedRange.Rows.Count < 2) Then
            MsgBox("没有出货内容，请检查！")
            Exit Sub
        End If
        For i = 2 To xlsheet.UsedRange.Rows.Count
            For j = 2 To 4
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
            If IsDate(xlsheet.Cells(i, 7).value) Then
            Else
                MsgBox("出货日期有问题，请检查！")
                Exit Sub
            End If
        Next
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim rs1 As New ADODB.Recordset

        Dim SQLstring, SQLstring1 As String
        Dim CheckErr As Integer
        CheckErr = 0

        SQLstring = "Select * From JamuOutput order by 序号"

        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 3)
        rs.MoveLast()
        ID = rs.Fields("序号").Value + 1
        BatchID = ID
        For i = 2 To xlsheet.UsedRange.Rows.Count
            SQLstring1 = "Select * From JamuOrder where 客户订单号=" & "'" & xlsheet.Cells(i, 5).value & "'" & " and 图号=" & "'" & xlsheet.Cells(i, 1).value & "'"
            rs1.Open(SQLstring1, cn, 1, 1)
            If rs1.RecordCount = 0 Then
                CheckErr = CheckErr + 1
                xlsheet.Cells(i, 8).value = "图号或单号有误"
            End If
            rs1.Close()
        Next
        If CheckErr > 0 Then
            If UserName <> "admin" Then
                MsgBox("出货图号或单号有误，请检查！")
                rs.Close()
                cn.Close()
                rs = Nothing
                rs1 = Nothing
                cn = Nothing
                Exit Sub
            End If
        End If
        For i = 2 To xlsheet.UsedRange.Rows.Count
            rs.AddNew()
            rs.Fields("序号").Value = ID
            rs.Fields("图号").Value = xlsheet.Cells(i, 1).value
            rs.Fields("数量").Value = xlsheet.Cells(i, 2).value
            rs.Fields("单价").Value = xlsheet.Cells(i, 3).value
            rs.Fields("总价").Value = xlsheet.Cells(i, 4).value
            rs.Fields("客户订单号").Value = xlsheet.Cells(i, 5).value
            rs.Fields("备注").Value = xlsheet.Cells(i, 6).value
            rs.Fields("送货日期").Value = xlsheet.Cells(i, 7).value
            rs.Fields("记录批号").Value = Format(BatchID, "000000")
            rs.Update()
            ID = ID + 1
            SQLstring1 = "Select * From JamuOrder where 客户订单号=" & "'" & xlsheet.Cells(i, 5).value & "'" & " and 图号=" & "'" & xlsheet.Cells(i, 1).value & "'"
            rs1.Open(SQLstring1, cn, 1, 3)
            'rs1.Fields("备货标记").Value = vbNull
            rs1.Fields("备货标记").Value = ""
            rs1.Update()
            rs1.Close()
        Next
        rs.Close()
        cn.Close()
        rs = Nothing
        rs1 = Nothing
        cn = Nothing
        ss.DisplayAlerts = False
        xlbook.Close()
        ss.Quit()
        xlsheet = Nothing
        xlbook = Nothing
        ss = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)
        MsgBox("嘉睦出货录入完成！")
    End Sub

    Private Sub Form53_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

End Class