Public Class Form工艺修改

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String

            SQLstring = "Select SFOrder,DrawNo,DWGInfo From ProcCard where SFOrder=" & "'" & TextBox1.Text & "' group by SFOrder,DrawNo,DWGInfo"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此单号无工艺卡！")
                Exit Sub
            End If
            rs.MoveFirst()
            DataGridView1.Rows.Clear()
            Do While Not rs.EOF
                DataGridView1.Rows.Add(1)
                DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("DWGInfo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DWGInfo").Value
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub Form35_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form35_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("DWGInfo", "图号全息")
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * From ProcCard"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        For i = 0 To rs.Fields.Count - 1
            DataGridView2.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
            If ((rs.Fields(i).Name = "ProcSN") Or (rs.Fields(i).Name = "ProcName") Or (rs.Fields(i).Name = "ProcDesc") Or (rs.Fields(i).Name = "ProcQty") Or (rs.Fields(i).Name = "PreTime") Or (rs.Fields(i).Name = "UnitTime") Or (rs.Fields(i).Name = "该工序预计单件工时") Or (rs.Fields(i).Name = "该工序预计单件成本") Or (rs.Fields(i).Name = "该工序预计总工时") Or (rs.Fields(i).Name = "该工序预计总成本")) Then
                DataGridView2.Columns(i).Visible = True
            Else
                DataGridView2.Columns(i).Visible = False
            End If
            If (DataGridView2.Columns(i).Name = "ProcSN") Then
                DataGridView2.Columns(i).ReadOnly = True
            End If
            DataGridView2.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        DataGridView2.Width = Me.Width - DataGridView1.Width - 150
        DataGridView1.Height = Me.Height - 200
        DataGridView2.Height = Me.Height - 200
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String

            SQLstring = "Select SFOrder,DrawNo,DWGInfo From ProcCard where DrawNo like " & "'%" & TextBox2.Text & "%' group by SFOrder,DrawNo,DWGInfo"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此图号无工艺卡！")
                Exit Sub
            End If
            rs.MoveFirst()
            DataGridView1.Rows.Clear()
            Do While Not rs.EOF
                DataGridView1.Rows.Add(1)
                DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("DWGInfo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DWGInfo").Value
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * From ProcCard where DWGInfo = " & "'" & DataGridView1.Item("DWGInfo", e.RowIndex).Value & "' order by ProcSN"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        DataGridView2.Rows.Clear()
        Do While rs.EOF = False
            DataGridView2.Rows.Add(1)
            For i = 0 To rs.Fields.Count - 1
                DataGridView2.Item(i, DataGridView2.Rows.Count - 1).Value = rs.Fields(i).Value
            Next
            rs.MoveNext()
        Loop
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        On Error GoTo ErrorExit
        If (DataGridView2.SelectedCells(0).RowIndex <> 0) Then
            If (DataGridView2.Item("ProcSN", DataGridView2.SelectedCells(0).RowIndex).Value = DataGridView2.Item("ProcSN", DataGridView2.SelectedCells(0).RowIndex - 1).Value + 1) Then
                Exit Sub
            End If
            DataGridView2.Rows.Insert(DataGridView2.SelectedCells(0).RowIndex)
            DataGridView2.Item("ProcSN", DataGridView2.SelectedCells(0).RowIndex - 1).ReadOnly = False
            DataGridView2.Item("ProcSN", DataGridView2.SelectedCells(0).RowIndex - 1).Value = DataGridView2.Item("ProcSN", DataGridView2.SelectedCells(0).RowIndex - 2).Value + 1
            'DataGridView2.Item("ProcSN", DataGridView2.SelectedCells(0).RowIndex - 1).Value = Format((DataGridView2.Item("ProcSN", DataGridView2.SelectedCells(0).RowIndex - 2).Value + DataGridView2.Item("ProcSN", DataGridView2.SelectedCells(0).RowIndex + 1).Value) / 2, "000")
        End If
ErrorExit:
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        On Error GoTo ErrExit
        If (DataGridView2.SelectedCells(0).RowIndex <> 0) Then
            DataGridView2.Rows.RemoveAt(DataGridView2.SelectedCells(0).RowIndex)
        End If
ErrExit:
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        Dim ID As Integer
        SQLstring = "delete ProcCard where DWGInfo = " & "'" & DataGridView2.Item("DWGInfo", 0).Value & "'"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 3)
        SQLstring = "Select * from ProcCard"
        rs.Open(SQLstring, cn, 1, 3)
        rs.MoveLast()
        ID = rs.Fields("ID").Value
        For j = 0 To DataGridView2.Rows.Count - 1
            rs.AddNew()
            ID = ID + 1
            rs(0).Value = ID
            For i = 1 To DataGridView2.Columns.Count - 1
                rs(i).Value = DataGridView2.Item(i, j).Value
            Next
            rs("ProcMaker").Value = ChineseName
            rs("DWGInfo").Value = DataGridView2.Item("DWGInfo", 0).Value
            rs("CustDWG").Value = DataGridView2.Item("CustDWG", 0).Value
            rs("SFOrder").Value = DataGridView2.Item("SFOrder", 0).Value
            rs("DrawNo").Value = DataGridView2.Item("DrawNo", 0).Value
            rs("PartName").Value = DataGridView2.Item("PartName", 0).Value
            rs("Qty").Value = DataGridView2.Item("Qty", 0).Value
            rs("CustCode").Value = DataGridView2.Item("CustCode", 0).Value
            rs("ProcQty").Value = DataGridView2.Item("ProcQty", 0).Value
            rs("ProcDate").Value = Now.Date
            rs("ProcCode").Value = DataGridView2.Item("SFOrder", 0).Value & DataGridView2.Item("DrawNo", 0).Value & DataGridView2.Item("ProcSN", j).Value
            rs.Update()
        Next
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        MsgBox("工艺修改完成！")
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        On Error GoTo ErrExit
        DataGridView2.Rows.Add(1)
        DataGridView2.Item("ProcSN", DataGridView2.Rows.Count - 1).Value = DataGridView2.Item("ProcSN", DataGridView2.Rows.Count - 2).Value + 5
ErrExit:
    End Sub
End Class