Public Class Form按客户代码及图号输入价格
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
        xlsheet.Name = "PriceUpdate"
        xlsheet.Cells(1, 1) = "图号"
        xlsheet.Cells(1, 2) = "单价"

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("Please input customer code!")
            Exit Sub
        End If
        If xlsheet.UsedRange.Rows.Count < 2 Then
            MsgBox("No data!")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring, CNString As String
        '********************************** Textbox1 connection ***************************************************************

        cn.Open(CNsfmdb)
        'rs.Open(SQLstring, cn, 1, 2)
        For i = 2 To xlsheet.UsedRange.Rows.Count
            If IsDBNull(xlsheet.Cells(i, 1)) Then
            Else
                SQLstring = "Select * From SFOrderBase where DrawNo=" & "'" & xlsheet.Cells(i, 1).value & "'" & " and CustCode=" & "'" & TextBox1.Text & "'"
                rs.Open(SQLstring, cn, 1, 3)
                If rs.RecordCount > 0 Then
                    rs.MoveFirst()
                    Do While rs.EOF = False
                        rs.Fields("UnitP").Value = xlsheet.Cells(i, 2).value
                        rs.Fields("RMBUnitP").Value = xlsheet.Cells(i, 2).value
                        rs.Update()
                        rs.MoveNext()
                    Loop
                End If
                rs.Close()
            End If
        Next

        rs = Nothing
        cn = Nothing
        ss.DisplayAlerts = False
        xlbook.Close()
        ss.Quit()
        xlsheet = Nothing
        xlbook = Nothing
        ss = Nothing
        'shell(SFOrderBaseNetDisconnect, vbHide)
        MsgBox("Price input！")

    End Sub

    Private Sub Form63_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

End Class