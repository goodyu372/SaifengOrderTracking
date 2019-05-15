Public Class Form主管工序扫描
    Public DWGInfoString As String
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = 13) Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring, CNString As String
            On Error Resume Next

            SQLstring = "Select * From ProcCard where ProcCode=" & "'" & (TextBox1.Text) & "'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                DWGInfoString = "Input"
                TextBox5.Text = "0" '订单数量
                TextBox6.Text = "0" '投产数量
                TextBox7.Text = rs.Fields("ProcDesc").Value
                TextBox3.Text = rs.Fields("ProcQty").Value
                TextBox4.Text = 0
                TextBox9.Text = "XXXX" 'CustCode
                TextBox10.Text = "XXXX" & Mid(TextBox1.Text, 6, 50) 'Drawing Info
                TextBox12.Text = "" '工时

                TextBox_工艺给出该工序单件工时.Text = "0.0"
                TextBox_工艺给出该工序单件成本.Text = "0.0"
                TextBox_工艺数量.Text = "0"
                TextBox_工艺给出该工序总工时.Text = "0.0"
                TextBox_工艺给出该工序总成本.Text = "0.0"

            Else
                DWGInfoString = "ProcessCard"
                TextBox5.Text = rs.Fields("Qty").Value
                TextBox6.Text = rs.Fields("ProcQty").Value '工艺数量
                TextBox7.Text = rs.Fields("ProcDesc").Value
                TextBox3.Text = rs.Fields("ProcQty").Value
                TextBox4.Text = 0
                TextBox9.Text = rs.Fields("CustCode").Value
                TextBox10.Text = rs.Fields("DWGInfo").Value
                TextBox12.Text = IIf(IsNumeric(rs.Fields("PreTime").Value), rs.Fields("PreTime").Value, 0) + rs.Fields("Qty").Value * IIf(IsNumeric(rs.Fields("UnitTime").Value), rs.Fields("UnitTime").Value, 0) '工时 WorkTime

                TextBox_工艺给出该工序单件工时.Text = rs.Fields("该工序预计单件工时").Value.ToString
                TextBox_工艺给出该工序单件成本.Text = rs.Fields("该工序预计单件成本").Value.ToString
                TextBox_工艺数量.Text = rs.Fields("ProcQty").Value.ToString
                TextBox_工艺给出该工序总工时.Text = rs.Fields("该工序预计总工时").Value.ToString
                TextBox_工艺给出该工序总成本.Text = rs.Fields("该工序预计总成本").Value.ToString
            End If

            rs.Close()
            cn.Close()
            TextBox2.Enabled = True
            TextBox2.Focus()
            StatusStrip1.Show()

            StatusStrip1.Text = "要显示的内容"
            ComboBox1.Focus()
        End If
    End Sub
    Private Sub Form5_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Focus()
        ComboBox2.Items.Add("自返修")
        ComboBox2.Items.Add("帮返修")
        ComboBox2.Items.Add("工装夹具治具加工")
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "Select 姓名 From SFHR where 在职标识=" & "'" & ProcessName & "'" & " order by 姓名"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 1)
        rs.MoveFirst()
        Do While Not rs.EOF
            ComboBox1.Items.Add(rs(0).Value)
            rs.MoveNext()
        Loop
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        StatusStrip1.Show()
        StatusStrip1.Items(2).Text = "登录时间：" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss")
        TextBox1.Focus()
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If (e.KeyCode = 13) Then
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring, CNString As String

            On Error Resume Next

            SQLstring = "Select * From SFHR where 工号=" & "'" & (TextBox2.Text) & "'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If (rs.RecordCount = 0) Then
                MsgBox("该工号不存在，请找系统管理员！")
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                Exit Sub
            End If
            TextBox8.Text = rs.Fields("姓名").Value
            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing
            TextBox2.Enabled = False
            TextBox3.Enabled = True
            TextBox3.Focus()
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If (e.KeyCode = 13) Then
            If IsNumeric(TextBox3.Text) Then
                If Val(TextBox5.Text) = 0 Then
                    TextBox12.Text = ""
                Else
                    TextBox12.Text = Val(TextBox3.Text) / Val(TextBox5.Text) * Val(TextBox12.Text)
                End If

            Else
                MsgBox("OK数量格式不对")
                Exit Sub
            End If
            TextBox4.Enabled = True
            TextBox4.Focus()
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ID As Integer
        If (IsNumeric(TextBox12.Text)) Then
        Else
            MsgBox("请输入工时!")
            Exit Sub
        End If
        If (TextBox1.Text = "") Then
            MsgBox("请扫描或输入单号图号!")
            Exit Sub
        End If
        If (TextBox8.Text = "") Then
            MsgBox("请输入员工!")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset

        Dim SQLstring As String

        cn.Open(CNsfmdb)
        SQLstring = "Select * from ProcRecord order by ID"
        rs.Open(SQLstring, cn, 1, 3)
        rs.MoveLast()
        ID = rs.Fields("ID").Value

        rs.AddNew()
        rs.Fields("ID").Value = ID + 1
        rs.Fields("DWGInfo").Value = IIf(Len(TextBox10.Text) < 5, "xxxxTempxxxxx", TextBox10.Text)
        Select Case DWGInfoString
            Case "Input"
                rs.Fields("CustDWG").Value = IIf(Len(TextBox10.Text) < 5, "xxxxTempxxxxx", TextBox10.Text)
            Case "ProcessCard"
                rs.Fields("CustDWG").Value = Mid(TextBox10.Text, 1, Len(TextBox10.Text) - 5)
        End Select
        rs.Fields("ProcCode").Value = TextBox1.Text
        rs.Fields("Operator").Value = TextBox8.Text
        rs.Fields("OKQty").Value = TextBox3.Text
        rs.Fields("ScrapQty").Value = IIf(TextBox4.Text = "", 0, TextBox4.Text)
        rs.Fields("Qty").Value = IIf(TextBox5.Text = "", 0, TextBox5.Text)
        rs.Fields("ProcQty").Value = IIf(TextBox6.Text = "", 0, TextBox6.Text)
        rs.Fields("RecordTime").Value = Now
        rs.Fields("RecordType").Value = IIf(DWGInfoString = "Input", ProcessName & "-工序完成", TextBox11.Text)
        rs.Fields("Status").Value = IIf(DWGInfoString = "Input", ProcessName & "-工序完成", Mid(TextBox1.Text, Len(TextBox1.Text) - 2, 3) & ProcessName & "-工序完成")
        rs.Fields("WorkTime").Value = Val(TextBox12.Text)

        rs.Fields("该工序预计单件工时").Value = Val(TextBox_工艺给出该工序单件工时.Text)
        rs.Fields("该工序预计单件成本").Value = Val(TextBox_工艺给出该工序单件成本.Text)
        rs.Fields("该工序预计总工时").Value = Val(TextBox_工艺给出该工序总工时.Text)
        rs.Fields("该工序预计总成本").Value = Val(TextBox_工艺给出该工序总成本.Text)
        rs.Update()
        rs.Close()
        '888888888888888888888888888888 更新订单状况 status 8888888888888888888888888888888888888888888
        SQLstring = "Select * from SFOrderBase where DWGInfo=N'" & TextBox10.Text & "'"
        rs.Open(SQLstring, cn, 1, 3)
        If rs.RecordCount = 0 Then
        Else
            rs.MoveFirst()
            rs.Fields("Status").Value = Mid(TextBox1.Text, Len(TextBox1.Text) - 2, 3) & ProcessName & "-工序完成"
            rs.Update()
        End If
        rs.Close()
        cn.Close()
        rs = Nothing
        cn = Nothing
        '8888888888888888888888888888 end of 更新订单状况 status 88888888888888888888888888888888888888
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        TextBox7.Clear()
        TextBox8.Clear()
        TextBox9.Clear()
        TextBox10.Clear()
        TextBox11.Clear()
        MsgBox("主管工序扫描完成！")
        TextBox1.Focus()
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox8.Text = ComboBox1.SelectedItem.ToString
        TextBox3.Enabled = True
        TextBox3.Focus()
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox11.Text = ComboBox2.SelectedItem.ToString
    End Sub
End Class