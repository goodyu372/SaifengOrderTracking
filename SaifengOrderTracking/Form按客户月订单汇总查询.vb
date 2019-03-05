Public Class Form按客户月订单汇总查询
    Private Sub Form71_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form71_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("CustCode", "客户代码")
        DataGridView1.Columns.Add("MonthlyOrder", "订单总值")
        DataGridView1.Columns.Add("YearMonth", "下单年月")
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            DataGridView1.Rows.Clear()
            If TextBox1.Text = "" Then
                MsgBox("年月不能为空！")
                Exit Sub
            End If
            If IsNumeric(TextBox1.Text) Then
            Else
                MsgBox("年月不能为非数字！")
                Exit Sub
            End If

            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select CustCode,YearMonth,format(sum(Amount),'00') as MonthlyOrder from OrderMoney where YearMonth =" & "'" & TextBox1.Text & "' group by CustCode,YearMonth"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有该年月的订单！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(rs(i).Name, DataGridView1.Rows.Count - 1).Value = rs(i).Value
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            SQLstring = "Select YearMonth,format(sum(amount),'00') as MonthlyOrder from OrderMoney where YearMonth =" & "'" & TextBox1.Text & "' group by YearMonth"
            rs.Open(SQLstring, cn, 1, 1)
            DataGridView1.Rows.Add(1)
            DataGridView1.Item("MonthlyOrder", DataGridView1.Rows.Count - 1).Value = rs("MonthlyOrder").Value
            DataGridView1.Item("CustCode", DataGridView1.Rows.Count - 1).Value = "当月汇总"
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
End Class