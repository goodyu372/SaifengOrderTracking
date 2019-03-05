Public Class Form嘉睦出货记录更改
    Public RowNo As Integer
    Private Sub Form52_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form52_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 200
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim cn As New ADODB.Connection
        Dim rs, rs1 As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select * from JamuOutput where DatePart(Year,送货日期)=" & "'" & DatePart(DateInterval.Year, DateTimePicker1.Value) & "'" & " and DatePart(Month,送货日期)=" & "'" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "'" & " and DatePart(Day,送货日期)=" & "'" & DatePart(DateInterval.Day, DateTimePicker1.Value) & "'"
        'SQLstring = "Select * from JamuOutput 送货日期 between " & "'" & "2018-04-28" & "'" & " and " & "'" & "2018-04-28" & "'"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
        If rs.RecordCount = 0 Then
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            MsgBox("没有该日期的送货记录！")
            Exit Sub
        End If
        rs.MoveFirst()
        For i = 0 To rs.Fields.Count - 1
            DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
        Next
        Do While rs.EOF = False
            DataGridView1.Rows.Add(1)
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = IIf(IsDBNull(rs(i).Value), "", rs(i).Value)
            Next
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            If TextBox2.Text = "" Then
                MsgBox("客户订单号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rs1 As New ADODB.Recordset
            Dim SQLstring As String
            'SQLstring = "Select * from JamuOutput where DatePart(Year,送货日期)=" & "'" & DatePart(DateInterval.Year, DateTimePicker1.Value) & "'" & " and DatePart(Month,送货日期)=" & "'" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "'" & " and DatePart(Day,送货日期)=" & "'" & DatePart(DateInterval.Day, DateTimePicker1.Value) & "'"
            SQLstring = "Select * from JamuOutput where 客户订单号 like " & "'%" & TextBox2.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有该客户订单号的送货记录！")
                Exit Sub
            End If
            rs.MoveFirst()
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
            Next
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = IIf(IsDBNull(rs(i).Value), "", rs(i).Value)
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            If TextBox1.Text = "" Then
                MsgBox("图号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rs1 As New ADODB.Recordset
            Dim SQLstring As String
            'SQLstring = "Select * from JamuOutput where DatePart(Year,送货日期)=" & "'" & DatePart(DateInterval.Year, DateTimePicker1.Value) & "'" & " and DatePart(Month,送货日期)=" & "'" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "'" & " and DatePart(Day,送货日期)=" & "'" & DatePart(DateInterval.Day, DateTimePicker1.Value) & "'"
            SQLstring = "Select * from JamuOutput where 图号 like " & "'%" & TextBox1.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()
            StatusStrip1.Items(0).Text = "记录条数：  " & rs.RecordCount
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有该图号的送货记录！")
                Exit Sub
            End If
            rs.MoveFirst()
            For i = 0 To rs.Fields.Count - 1
                DataGridView1.Columns.Add(rs(i).Name, rs(i).Name)
            Next
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(i, DataGridView1.Rows.Count - 1).Value = IIf(IsDBNull(rs(i).Value), "", rs(i).Value)
                Next
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RowNo = 9999 Then
            MsgBox("Please select the item to be changed!")
            Exit Sub
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        RowNo = e.RowIndex
    End Sub

    Private Sub Form86_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RowNo = 9999
    End Sub
End Class