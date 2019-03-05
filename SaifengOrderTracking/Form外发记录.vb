Public Class Form外发记录
    Public SubType As String
    Private Sub Form24_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form24_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("Qty", "数量")
        DataGridView2.Rows.Clear()
        DataGridView2.Columns.Clear()
        DataGridView2.Columns.Add("供应商名称", "供应商名称")
        DataGridView2.Columns.Add("加工项目", "加工项目")
        DataGridView3.Rows.Clear()
        DataGridView3.Columns.Add("SFOrder", "赛峰单号")
        DataGridView3.Columns.Add("DrawNo", "图号")
        DataGridView3.Columns.Add("OKQty", "数量")
        DataGridView3.Columns.Add("SubCompany", "供应商")
        DataGridView3.Columns.Add("UnitPrice", "外发单价")
        DataGridView3.Columns.Add("Comment", "外发项目备注")
        'shell(SFOrderBaseNetConnect, vbHide)
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        cn.Open(CNsfmdb)
        SQLstring = "Select * from SubList"
        rs.Open(SQLstring, cn, 1, 2)
        rs.MoveFirst()
        Do While rs.EOF = False
            DataGridView2.Rows.Add(1)
            DataGridView2.Item("供应商名称", DataGridView2.Rows.Count - 1).Value = rs.Fields("供应商名称").Value
            DataGridView2.Item("加工项目", DataGridView2.Rows.Count - 1).Value = rs.Fields("加工项目").Value
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        GroupBox2.Visible = True
        SubType = "工序外发"
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (DataGridView3.Rows.Count = 0) Then
            MsgBox("请输入外发零件！")
            Exit Sub
        End If
        For i = 0 To DataGridView3.Rows.Count - 1
            If (DataGridView3.Item("SubCompany", i).Value <> "") Then
                'If (IsNumeric(DataGridView1.Item("OKQty", i).Value)) Then
            Else
                MsgBox("请检查供应商！")
                Exit Sub
            End If
        Next
        For i = 0 To DataGridView3.Rows.Count - 1
            'If (DataGridView1.Item("OKQty", i).Value = "") Then
            If (IsNumeric(DataGridView3.Item("OKQty", i).Value)) Then
            Else
                MsgBox("请检查数量！")
                Exit Sub
            End If
        Next
        Dim ProgramLog As String
        '88888888888888888  modify status for SFOrderbase  888888888888888888888888888888888888888
        Dim cn As New ADODB.Connection
        Dim rs, rs1 As New ADODB.Recordset
        Dim SQLstring As String
        Dim CustDWG, DWGInfo, ProcCode, GYS As String
        Dim ID, Qty, OKQty As Integer

        cn.Open(CNsfmdb)
        For i = 0 To DataGridView3.Rows.Count - 1
            ProgramLog = ""


            'SQLstring = "Select * from SubRecord where ((SFOrder=" & "'" & (DataGridView3.Item("SFOrder", i).Value) & "') and (DrawNo=" & "'" & (DataGridView3.Item("DrawNo", i).Value) & "'))"
            SQLstring = "Select * from SubRecord where ((SFOrder=" & "'" & (DataGridView3.Item("SFOrder", i).Value) & "') and (DrawNo=" & "'" & (DataGridView3.Item("DrawNo", i).Value) & "') and isnull(BackDate,'')='')"
            rs.Open(SQLstring, cn, 1, 3)
            If (rs.RecordCount = 0) Then
                rs.AddNew()
                rs.Fields("SFOrder").Value = DataGridView3.Item("SFOrder", i).Value
                rs.Fields("DrawNo").Value = DataGridView3.Item("DrawNo", i).Value
                rs.Fields("OKQty").Value = DataGridView3.Item("OKQty", i).Value
                rs.Fields("SubCompany").Value = DataGridView3.Item("SubCompany", i).Value
                rs.Fields("UnitPrice").Value = DataGridView3.Item("UnitPrice", i).Value
                rs.Fields("Comment").Value = DataGridView3.Item("Comment", i).Value
                rs.Fields("SubDate").Value = Now
                rs.Fields("Operator").Value = ChineseName
                rs.Fields("SubType").Value = SubType
                rs.Fields("Status").Value = "887-外发"
                rs.Fields("ProgramLog").Value = ProgramLog
                rs.Update()
                '******************************************* Update SFOrderBase and ProcRecord ***********************************************************************
                SQLstring = "Select * from SFOrderBase where ((SFOrder=" & "'" & (DataGridView3.Item("SFOrder", i).Value) & "') and (DrawNo=" & "'" & (DataGridView3.Item("DrawNo", i).Value) & "'))"
                rs1.Open(SQLstring, cn, 1, 3)
                If (rs1.RecordCount = 0) Then
                    ProgramLog = "NoBase"
                    rs1.Close()
                Else
                    rs1.MoveFirst()
                    DWGInfo = rs1.Fields("DWGInfo").Value
                    CustDWG = rs1.Fields("CustDWG").Value
                    ProcCode = rs1.Fields("SFOrder").Value & rs1.Fields("DrawNo").Value & "887"
                    Qty = rs1.Fields("Qty").Value
                    OKQty = DataGridView3.Item("OKQty", i).Value
                    GYS = DataGridView3.Item("SubCompany", i).Value
                    rs1.Fields("Status").Value = "887-外发[" & GYS & "]"
                    rs1.Update()
                    rs1.Close()

                    SQLstring = "Select * from ProcRecord order by ID"
                    rs1.Open(SQLstring, cn, 1, 3)
                    rs1.MoveLast()
                    ID = rs1.Fields("ID").Value
                    rs1.AddNew()
                    rs1.Fields("ID").Value = ID + 1
                    rs1.Fields("DWGInfo").Value = DWGInfo
                    rs1.Fields("CustDWG").Value = CustDWG
                    rs1.Fields("ProcCode").Value = ProcCode
                    rs1.Fields("Qty").Value = Qty
                    rs1.Fields("OKQty").Value = OKQty
                    rs1.Fields("RecordTime").Value = Now
                    rs1.Fields("Status").Value = "887-外发"
                    rs1.Fields("Operator").Value = ChineseName
                    rs1.Fields("RecordType").Value = "外发[" & GYS & "]"
                    rs1.Update()
                    rs1.Close()
                End If
                '******************************************* end of Update SFOrderBase and ProcRecord ***********************************************************************
            End If

            rs.Close()
        Next
        cn.Close()
        rs = Nothing
        cn = Nothing
        DataGridView3.Rows.Clear()
        '8888888888888888888888 end of modify status for SFOrderbase 88888888888888888888888888888
        MsgBox("完成外发记录！", , "外发记录输入")
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        GroupBox2.Visible = True
        SubType = "整体外发"
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        Label3.Visible = True
        Label4.Visible = True
        TextBox2.Visible = True
        TextBox3.Visible = True
        Label2.Visible = True
        TextBox1.Visible = True
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        DataGridView3.Rows.Add(1)
        DataGridView3.Item("SFOrder", DataGridView3.Rows.Count - 1).Value = DataGridView1.Item("SFOrder", e.RowIndex).Value
        DataGridView3.Item("DrawNo", DataGridView3.Rows.Count - 1).Value = DataGridView1.Item("DrawNo", e.RowIndex).Value
        DataGridView3.Item("SubCompany", DataGridView3.Rows.Count - 1).Value = DataGridView2.Item("供应商名称", DataGridView2.SelectedCells(0).RowIndex).Value
        DataGridView3.Item("Comment", DataGridView3.Rows.Count - 1).Value = DataGridView2.Item("加工项目", DataGridView2.SelectedCells(0).RowIndex).Value
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If (e.KeyCode = 13) Then
            Dim SFOrder As String
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SFOrder = TextBox2.Text
            DataGridView1.Rows.Clear()
            SQLstring = "Select * from SFOrderBase where DrawNo like" & "'%" & TextBox3.Text & "%'"
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("该图号不存在！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
                DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("Qty", DataGridView1.Rows.Count - 1).Value = rs.Fields("Qty").Value
                rs.MoveNext()
            Loop
            TextBox3.Text = ""
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If (e.KeyCode = 13) Then
            Dim SFOrder As String
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SFOrder = TextBox2.Text
            DataGridView1.Rows.Clear()
            SQLstring = "Select * from SFOrderBase where SFOrder=" & "'" & SFOrder & "'"
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("该单号不存在！")
                Exit Sub
            End If
            rs.MoveFirst()
            Do While rs.EOF = False
                DataGridView1.Rows.Add(1)
                DataGridView1.Item("SFOrder", DataGridView1.Rows.Count - 1).Value = SFOrder
                DataGridView1.Item("DrawNo", DataGridView1.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
                DataGridView1.Item("Qty", DataGridView1.Rows.Count - 1).Value = rs.Fields("Qty").Value
                rs.MoveNext()
            Loop
            TextBox2.Text = ""
        End If
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DataGridView3.Rows.RemoveAt(DataGridView3.SelectedCells(0).RowIndex)
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            Dim SFOrder As String
            Dim DrawNo As String
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String
            cn.Open(CNsfmdb)
            SFOrder = Mid(TextBox1.Text, 1, 5)
            DrawNo = Mid(TextBox1.Text, 6, Len(TextBox1.Text) - 8)
            DataGridView1.Rows.Clear()
            SQLstring = "Select * from SFOrderBase where SFOrder=" & "'" & SFOrder & "' and DrawNo=" & "'" & DrawNo & "'"
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("该单号图号信息不存在！")
                Exit Sub
            End If
            rs.MoveFirst()
            DataGridView3.Rows.Add(1)
            DataGridView3.Item("SFOrder", DataGridView3.Rows.Count - 1).Value = SFOrder
            DataGridView3.Item("DrawNo", DataGridView3.Rows.Count - 1).Value = DrawNo
            DataGridView3.Item("OKQty", DataGridView3.Rows.Count - 1).Value = rs.Fields("Qty").Value
            DataGridView3.Item("SubCompany", DataGridView3.Rows.Count - 1).Value = DataGridView2.Item("供应商名称", DataGridView2.SelectedCells(0).RowIndex).Value
            DataGridView3.Item("Comment", DataGridView3.Rows.Count - 1).Value = DataGridView2.Item("加工项目", DataGridView2.SelectedCells(0).RowIndex).Value
            TextBox1.Text = ""
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
        End If
    End Sub
End Class