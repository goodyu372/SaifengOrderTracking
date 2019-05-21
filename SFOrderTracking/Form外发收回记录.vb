Public Class Form外发收回记录

    Private Sub Form25_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form25_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("OKQty", "数量")
        DataGridView1.Columns.Add("SubCompany", "供应商")
        DataGridView1.Columns.Add("UnitPrice", "外发单价")
        DataGridView1.Columns.Add("Comment", "外发项目备注")
        DataGridView1.Columns.Add("SubDate", "外发日期")
        DataGridView1.Columns.Add("BackDate", "收回日期")
        DataGridView1.Columns.Add("Operator", "经手人")

        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView2.Columns.Add("SFOrder", "赛峰单号")
        DataGridView2.Columns.Add("DrawNo", "图号")
        DataGridView2.Columns.Add("OKQty", "数量")
        DataGridView2.Columns.Add("SubCompany", "供应商")
        DataGridView2.Columns.Add("UnitPrice", "外发单价")
        DataGridView2.Columns.Add("Comment", "外发项目备注")
        DataGridView2.Columns.Add("SubDate", "外发日期")
        DataGridView2.Columns.Add("BackDate", "收回日期")
        DataGridView2.Columns.Add("Operator", "经手人")
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select SubCompany from SubRecord where isnull(BackDate,'')='' group by SubCompany Order by SubCompany"
        rs.Open(SQLstring, cn, 1, 1)
        If (rs.RecordCount = 0) Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If
        rs.MoveFirst()
        Do While rs.EOF = False
            ComboBox1.Items.Add(rs.Fields("SubCompany").Value)
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub
    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        DataGridView2.Rows.Add(1)
        DataGridView2.Item("SFOrder", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("SFOrder", e.RowIndex).Value
        DataGridView2.Item("DrawNo", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("DrawNo", e.RowIndex).Value
        DataGridView2.Item("OKQty", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("OKQty", e.RowIndex).Value
        DataGridView2.Item("SubCompany", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("SubCompany", e.RowIndex).Value
        DataGridView2.Item("UnitPrice", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("UnitPrice", e.RowIndex).Value
        DataGridView2.Item("SubDate", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("SubDate", e.RowIndex).Value
        DataGridView2.Item("BackDate", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("BackDate", e.RowIndex).Value
        DataGridView2.Item("Comment", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("Comment", e.RowIndex).Value
        DataGridView2.Item("Operator", DataGridView2.Rows.Count - 1).Value = DataGridView1.Item("Operator", e.RowIndex).Value
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        Dim CustDWG, DWGInfo, ProcCode, GYS As String
        Dim ID, Qty, OKQty As Integer

        cn.Open(CNsfmdb)
        For i = 0 To DataGridView2.Rows.Count - 1
            SQLstring = "Select * from SFOrderBase where ((SFOrder=" & "'" & (DataGridView2.Item("SFOrder", i).Value) & "') and (DrawNo=" & "'" & (DataGridView2.Item("DrawNo", i).Value) & "')) and status<>'已终检' and status<>'完成入库'"
            rs.Open(SQLstring, cn, 1, 3)
            GYS = DataGridView1.Item("SubCompany", i).Value
            If (rs.RecordCount = 0) Then
                rs.Close()
            Else
                rs.MoveFirst()
                DWGInfo = rs.Fields("DWGInfo").Value
                CustDWG = rs.Fields("CustDWG").Value
                ProcCode = rs.Fields("SFOrder").Value & rs.Fields("DrawNo").Value & "888"
                Qty = rs.Fields("Qty").Value
                OKQty = DataGridView1.Item("OKQty", i).Value
                rs.Fields("Status").Value = "888-外发[" & GYS & "]回"
                rs.Update()
                rs.Close()

                SQLstring = "Select * from ProcRecord order by ID"
                rs.Open(SQLstring, cn, 1, 3)
                rs.MoveLast()
                ID = rs.Fields("ID").Value
                rs.AddNew()
                rs.Fields("ID").Value = ID + 1
                rs.Fields("DWGInfo").Value = DWGInfo
                rs.Fields("CustDWG").Value = CustDWG
                rs.Fields("ProcCode").Value = ProcCode
                rs.Fields("Qty").Value = Qty
                rs.Fields("OKQty").Value = OKQty
                rs.Fields("RecordTime").Value = Now
                rs.Fields("Status").Value = "888-外发[" & GYS & "]回"
                rs.Fields("Operator").Value = ChineseName
                rs.Fields("RecordType").Value = "外发回"
                rs.Update()
                rs.Close()
            End If
            SQLstring = "Select * from SubRecord where ((SFOrder=" & "'" & (DataGridView2.Item("SFOrder", i).Value) & "') and (DrawNo=" & "'" & (DataGridView2.Item("DrawNo", i).Value) & "')) and isnull(BackDate,'')=''"
            rs.Open(SQLstring, cn, 1, 3)
            If (rs.RecordCount <> 0) Then
                rs.MoveFirst()
                rs.Fields("SFOrder").Value = DataGridView2.Item("SFOrder", i).Value
                rs.Fields("DrawNo").Value = DataGridView2.Item("DrawNo", i).Value
                rs.Fields("OKQty").Value = DataGridView2.Item("OKQty", i).Value
                rs.Fields("SubCompany").Value = DataGridView2.Item("SubCompany", i).Value
                rs.Fields("UnitPrice").Value = DataGridView2.Item("UnitPrice", i).Value
                rs.Fields("BackDate").Value = Now
                rs.Fields("Operator").Value = ChineseName
                rs.Fields("Status").Value = "888-外发[" & GYS & "]回"
                rs.Fields("Comment").Value = DataGridView2.Item("Comment", i).Value
                rs.Update()
                rs.Close()
            End If
        Next

        DataGridView1.Rows.Clear()
        SQLstring = "Select * from SubRecord where isnull(BackDate,'')=''"
        rs.Open(SQLstring, cn, 1, 1)
        If (rs.RecordCount = 0) Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            Exit Sub
        End If
        DataGridView1.Rows.Add(rs.RecordCount)
        rs.MoveFirst()
        Dim DataRow As Integer
        DataRow = 0
        Do While rs.EOF = False
            DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
            DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
            DataGridView1.Item("OKQty", DataRow).Value = rs.Fields("OKQty").Value
            DataGridView1.Item("SubCompany", DataRow).Value = rs.Fields("SubCompany").Value
            DataGridView1.Item("UnitPrice", DataRow).Value = rs.Fields("UnitPrice").Value
            DataGridView1.Item("BackDate", DataRow).Value = rs.Fields("BackDate").Value
            DataGridView1.Item("Comment", DataRow).Value = rs.Fields("Comment").Value
            DataGridView1.Item("Operator", DataRow).Value = rs.Fields("Operator").Value
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop
        DataGridView2.Rows.Clear()
        rs.Close()

        rs = Nothing
        cn = Nothing
        MsgBox("确认外发收回记录！", , "外发收回记录输入")
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        DataGridView1.Rows.Clear()
        cn.Open(CNsfmdb)
        SQLstring = "Select * from SubRecord where ((isnull(BackDate,'')='') and (SubCompany=" & "'" & ComboBox1.SelectedItem.ToString & "'))"
        rs.Open(SQLstring, cn, 1, 1)
        DataGridView1.Rows.Add(rs.RecordCount)
        rs.MoveFirst()
        Dim DataRow As Integer
        DataRow = 0
        Do While rs.EOF = False
            DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
            DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
            DataGridView1.Item("OKQty", DataRow).Value = rs.Fields("OKQty").Value
            DataGridView1.Item("SubCompany", DataRow).Value = rs.Fields("SubCompany").Value
            DataGridView1.Item("UnitPrice", DataRow).Value = rs.Fields("UnitPrice").Value
            DataGridView1.Item("SubDate", DataRow).Value = rs.Fields("SubDate").Value
            DataGridView1.Item("BackDate", DataRow).Value = rs.Fields("BackDate").Value
            DataGridView1.Item("Comment", DataRow).Value = rs.Fields("Comment").Value
            DataGridView1.Item("Operator", DataRow).Value = rs.Fields("Operator").Value
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub
    Private Sub MonthCalendar1_DateSelected(sender As Object, e As DateRangeEventArgs)
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        DataGridView1.Rows.Clear()
        cn.Open(CNsfmdb)
        SQLstring = "Select * from SubRecord where ((isnull(BackDate,'')='') and (SubCompany=" & "'" & ComboBox1.SelectedItem.ToString & "'))"
        rs.Open(SQLstring, cn, 1, 1)
        DataGridView1.Rows.Add(rs.RecordCount)
        rs.MoveFirst()
        Dim DataRow As Integer
        DataRow = 0
        Do While rs.EOF = False
            DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
            DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
            DataGridView1.Item("OKQty", DataRow).Value = rs.Fields("OKQty").Value
            DataGridView1.Item("SubCompany", DataRow).Value = rs.Fields("SubCompany").Value
            DataGridView1.Item("UnitPrice", DataRow).Value = rs.Fields("UnitPrice").Value
            DataGridView1.Item("SubDate", DataRow).Value = rs.Fields("SubDate").Value
            DataGridView1.Item("BackDate", DataRow).Value = rs.Fields("BackDate").Value
            DataGridView1.Item("Comment", DataRow).Value = rs.Fields("Comment").Value
            DataGridView1.Item("Operator", DataRow).Value = rs.Fields("Operator").Value
            DataRow = DataRow + 1
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing

        For i = DataGridView1.Rows.Count - 1 To 0 Step -1
            If (DateValue(e.End.ToShortDateString) <> DateValue(DataGridView1.Item("SubDate", i).Value)) Then
                DataGridView1.Rows.RemoveAt(i)
            End If
        Next
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            DataGridView1.Rows.Clear()
            cn.Open(CNsfmdb)
            SQLstring = "Select * from SubRecord where ((isnull(BackDate,'')='') and (SFOrder=" & "'" & TextBox1.Text & "'))"
            rs.Open(SQLstring, cn, 1, 1)

            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有此单号的外发记录！")
                Exit Sub
            End If
            DataGridView1.Rows.Add(rs.RecordCount)
            rs.MoveFirst()
            Dim DataRow As Integer
            DataRow = 0
            Do While rs.EOF = False
                DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("OKQty", DataRow).Value = rs.Fields("OKQty").Value
                DataGridView1.Item("SubCompany", DataRow).Value = rs.Fields("SubCompany").Value
                DataGridView1.Item("UnitPrice", DataRow).Value = rs.Fields("UnitPrice").Value
                DataGridView1.Item("SubDate", DataRow).Value = rs.Fields("SubDate").Value
                DataGridView1.Item("BackDate", DataRow).Value = rs.Fields("BackDate").Value
                DataGridView1.Item("Comment", DataRow).Value = rs.Fields("Comment").Value
                DataGridView1.Item("Operator", DataRow).Value = rs.Fields("Operator").Value
                DataRow = DataRow + 1
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            DataGridView1.Rows.Clear()
            cn.Open(CNsfmdb)
            SQLstring = "Select * from SubRecord where ((isnull(BackDate,'')='') and (DrawNo like " & "'%" & TextBox2.Text & "%'))"
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有此图号的外发记录！")
                Exit Sub
            End If
            DataGridView1.Rows.Add(rs.RecordCount)
            rs.MoveFirst()
            Dim DataRow As Integer
            DataRow = 0
            Do While rs.EOF = False
                DataGridView1.Item("SFOrder", DataRow).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("DrawNo", DataRow).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("OKQty", DataRow).Value = rs.Fields("OKQty").Value
                DataGridView1.Item("SubCompany", DataRow).Value = rs.Fields("SubCompany").Value
                DataGridView1.Item("UnitPrice", DataRow).Value = rs.Fields("UnitPrice").Value
                DataGridView1.Item("SubDate", DataRow).Value = rs.Fields("SubDate").Value
                DataGridView1.Item("BackDate", DataRow).Value = rs.Fields("BackDate").Value
                DataGridView1.Item("Comment", DataRow).Value = rs.Fields("Comment").Value
                DataGridView1.Item("Operator", DataRow).Value = rs.Fields("Operator").Value
                DataRow = DataRow + 1
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If DataGridView2.SelectedCells.Count <> 0 Then
            DataGridView2.Rows.RemoveAt(DataGridView2.SelectedCells(0).RowIndex)
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class