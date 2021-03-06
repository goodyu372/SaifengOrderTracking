﻿Public Class Form按单号图号录入终检记录
    Private Sub Form32_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form32_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("CustCode", "客户代码")
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("PartName", "零件名称")
        DataGridView1.Columns.Add("Qty", "下单数量")
        DataGridView1.Columns.Add("DDate", "交期")
        DataGridView1.Columns.Add("OK", "好品数")
        DataGridView1.Columns.Add("NG", "坏品数")
        DataGridView1.Columns.Add("Comment", "备注")
        DataGridView1.Columns("CustCode").ReadOnly = True
        DataGridView1.Columns("SFOrder").ReadOnly = True
        DataGridView1.Columns("DrawNo").ReadOnly = True
        DataGridView1.Columns("PartName").ReadOnly = True
        DataGridView1.Columns("Qty").ReadOnly = True
        DataGridView1.Columns("DDate").ReadOnly = True
        DataGridView1.Columns("OK").DefaultCellStyle.BackColor = Color.Pink
        DataGridView1.Columns("NG").DefaultCellStyle.BackColor = Color.Pink
        DataGridView1.Columns("Comment").DefaultCellStyle.BackColor = Color.Pink

        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView2.Columns.Add("CustCode", "客户代码")
        DataGridView2.Columns.Add("SFOrder", "赛峰单号")
        DataGridView2.Columns.Add("DrawNo", "图号")
        DataGridView2.Columns.Add("PartName", "零件名称")
        DataGridView2.Columns.Add("Qty", "下单数量")
        DataGridView2.Columns.Add("DDate", "交期")
    End Sub

    Private Sub Form32_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView2.Width = Me.Width - 100
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox3.Text = "" Then
            MsgBox("请扫描或输入工号！")
            Exit Sub
        End If
        If DataGridView1.Rows.Count = 0 Then
            MsgBox("无扫描记录！")
            Exit Sub
        End If
        For i = 0 To DataGridView1.Rows.Count - 1
            If Not IsNumeric(DataGridView1.Item("OK", i).Value) Then
                MsgBox("请检查好品数！")
                Exit Sub
            End If
        Next
        For i = 0 To DataGridView1.Rows.Count - 1
            If Not IsNumeric(DataGridView1.Item("NG", i).Value) Then
                MsgBox("请检查坏品数！")
                Exit Sub
            End If
        Next
        Dim cn As New ADODB.Connection
        Dim rs, rs1 As New ADODB.Recordset
        Dim SQLstring As String
        Dim ID As Integer
        SQLstring = "Select * from QARecord order by ID"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 3)
        If rs.RecordCount = 0 Then
            ID = 0
        Else
            rs.MoveLast()
            ID = rs.Fields("ID").Value
        End If
        For i = 0 To DataGridView1.Rows.Count - 1
            rs.AddNew()
            rs.Fields("ID").Value = ID + 1
            rs.Fields("InspectDate").Value = Now
            rs.Fields("Operator").Value = TextBox3.Text
            rs.Fields("SFOrder").Value = DataGridView1.Item("SFOrder", i).Value
            rs.Fields("DrawNo").Value = DataGridView1.Item("DrawNo", i).Value
            rs.Fields("CustCode").Value = DataGridView1.Item("CustCode", i).Value
            rs.Fields("Qty").Value = DataGridView1.Item("Qty", i).Value
            rs.Fields("DDate").Value = DataGridView1.Item("DDate", i).Value
            rs.Fields("PartName").Value = DataGridView1.Item("PartName", i).Value
            rs.Fields("OK").Value = DataGridView1.Item("OK", i).Value
            rs.Fields("NG").Value = DataGridView1.Item("NG", i).Value
            rs.Fields("Comment").Value = DataGridView1.Item("Comment", i).Value
            rs.Update()
            SQLstring = "Select * from SFOrderBase where SFOrder=" & "'" & DataGridView1.Item("SFOrder", i).Value & "' and DrawNo=" & "'" & DataGridView1.Item("DrawNo", i).Value & "'"
            rs1.Open(SQLstring, cn, 1, 3)
            rs1.MoveFirst()
            rs1.Fields("Status").Value = "已终检"
            rs1.Update()
            rs1.Close()
        Next
        rs.Close()
        DataGridView1.Rows.Clear()
        MsgBox("终检录入完成！")
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        DataGridView1.Rows.Add(1)
        DataGridView1.Item("CustCode", DataGridView1.Rows.Count - 1).Value = DataGridView2.Item("CustCode", e.RowIndex).Value
        DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = DataGridView2.Item("SFOrder", e.RowIndex).Value
        DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = DataGridView2.Item("DrawNo", e.RowIndex).Value
        DataGridView1.Item("PartName", DataGridView1.Rows.Count - 1).Value = DataGridView2.Item("PartName", e.RowIndex).Value
        DataGridView1.Item("Qty", DataGridView1.Rows.Count - 1).Value = DataGridView2.Item("Qty", e.RowIndex).Value
        DataGridView1.Item("DDate", DataGridView1.Rows.Count - 1).Value = DataGridView2.Item("DDate", e.RowIndex).Value
    End Sub
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = 13) Then
            DataGridView1.Rows.Clear()
            If (TextBox1.Text = "") Then
                MsgBox("请输入工号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rs1 As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * from SFHR where 工号=" & "'" & TextBox1.Text & "'"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有此工号，请检查！")
                Exit Sub
            End If
            rs.MoveFirst()
            TextBox3.Text = rs.Fields("姓名").Value
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            DataGridView2.Rows.Clear()
            Dim cn As New ADODB.Connection
            Dim rs, rs1 As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * from SFOrderBase where SFOrder=" & "'" & TextBox2.Text & "'"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有此单号，请检查！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView2.Rows.Add(1)
                DataGridView2.Item("SFOrder", DataGridView2.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
                DataGridView2.Item("DrawNo", DataGridView2.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
                DataGridView2.Item("CustCode", DataGridView2.Rows.Count - 1).Value = rs.Fields("CustCode").Value
                DataGridView2.Item("Qty", DataGridView2.Rows.Count - 1).Value = rs.Fields("Qty").Value
                DataGridView2.Item("DDate", DataGridView2.Rows.Count - 1).Value = rs.Fields("DDate").Value
                DataGridView2.Item("PartName", DataGridView2.Rows.Count - 1).Value = rs.Fields("PartName").Value
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = 13 Then
            DataGridView2.Rows.Clear()
            Dim cn As New ADODB.Connection
            Dim rs, rs1 As New ADODB.Recordset
            Dim SQLstring As String
            SQLstring = "Select * from SFOrderBase where DrawNo like " & "'%" & TextBox4.Text & "%'"
            cn.Open(CNsfmdb)
            'rs.Open(SQLstring, cn, 1, 2)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有此图号，请检查！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView2.Rows.Add(1)
                DataGridView2.Item("SFOrder", DataGridView2.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
                DataGridView2.Item("DrawNo", DataGridView2.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
                DataGridView2.Item("CustCode", DataGridView2.Rows.Count - 1).Value = rs.Fields("CustCode").Value
                DataGridView2.Item("Qty", DataGridView2.Rows.Count - 1).Value = rs.Fields("Qty").Value
                DataGridView2.Item("DDate", DataGridView2.Rows.Count - 1).Value = rs.Fields("DDate").Value
                DataGridView2.Item("PartName", DataGridView2.Rows.Count - 1).Value = rs.Fields("PartName").Value
                rs.MoveNext()
            Loop
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class