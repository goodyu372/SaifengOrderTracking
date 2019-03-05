Public Class Form返工记录

    Private Sub Form41_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form41_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("普车")
        ComboBox1.Items.Add("数控车")
        ComboBox1.Items.Add("铣床")
        ComboBox1.Items.Add("磨床")
        ComboBox1.Items.Add("CNC")
        ComboBox1.Items.Add("大水磨")
        ComboBox1.Items.Add("快走丝")
        ComboBox1.Items.Add("火花机")
        ComboBox1.Items.Add("钳工")
        ComboBox1.Items.Add("外发")
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("SFOrder", "单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("ReworkQty", "返工数量")
        DataGridView1.Columns.Add("ReworkDate", "返工日期")
        DataGridView1.Columns.Add("RWFinishDate", "返工完成日期")
        DataGridView1.Columns.Add("ReworkBackQty", "返工完成数量")
        DataGridView1.Columns.Add("Operator", "操作者")
        DataGridView1.Columns.Add("ReworkProcess", "返工工序")
        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).ReadOnly = True
        Next
        DataGridView1.Columns("ReworkQty").ReadOnly = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("没有返工记录！")
            Exit Sub
        End If
        For i = 0 To DataGridView1.Rows.Count - 1
            If (IsNumeric(DataGridView1.Item("ReworkQty", i).Value)) Then
            Else
                MsgBox("请检查返工数量！")
                Exit Sub
            End If
        Next
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        Dim ID As Integer
        cn.Open(CNsfmdb)
        For i = 0 To DataGridView1.Rows.Count - 1
            SQLstring = "Select * from ReworkRecord order by ID"
            rs.Open(SQLstring, cn, 1, 3)
            If rs.RecordCount = 0 Then
                ID = 1
            Else
                rs.MoveLast()
                ID = rs.Fields("ID").Value + 1
            End If
            rs.AddNew()
            rs.Fields("SFOrder").Value = DataGridView1.Item("SFOrder", i).Value
            rs.Fields("DrawNo").Value = DataGridView1.Item("DrawNo", i).Value
            rs.Fields("ReworkQty").Value = DataGridView1.Item("ReworkQty", i).Value
            rs.Fields("ReworkDate").Value = Now
            rs.Fields("Operator").Value = DataGridView1.Item("Operator", i).Value
            rs.Fields("ReworkProcess").Value = DataGridView1.Item("ReworkProcess", i).Value
            rs.Fields("ID").Value = ID
            rs.Update()
            rs.Close()
        Next
        cn.Close()
        rs = Nothing
        cn = Nothing
        DataGridView1.Rows.Clear()
        MsgBox("返工记录输入完成！")
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        '工号输入，对应姓名
        If e.KeyCode = 13 Then
            TextBox3.Clear()
            If TextBox2.Text = "" Then
                MsgBox("请输入工号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SQLstring = "Select 工号,姓名 from SFHR where 工号=" & "'" & TextBox2.Text & "'"
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此工号不存在！")
                Exit Sub
            End If
            rs.MoveFirst()
            TextBox3.Text = rs.Fields("姓名").Value
            Label2.Visible = True
            ComboBox1.Visible = True
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        '选择返工工序
        Label1.Visible = True
        TextBox1.Visible = True
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        'Scan the SFOrder and DrawNo
        If e.KeyCode = 13 Then
            If Len(TextBox1.Text) < 9 Then
                MsgBox("扫描信息有问题,请检查！")
                Exit Sub
            End If
            DataGridView1.Rows.Add(1)
            DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = Mid(TextBox1.Text, 1, 5)
            DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = Mid(TextBox1.Text, 6, Len(TextBox1.Text) - 8)
            DataGridView1.Item("Operator", DataGridView1.Rows.Count - 1).Value = TextBox3.Text
            DataGridView1.Item("ReworkProcess", DataGridView1.Rows.Count - 1).Value = ComboBox1.SelectedItem.ToString
            DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = Mid(TextBox1.Text, 1, 5)
            DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = Mid(TextBox1.Text, 1, 5)
        End If
    End Sub
End Class