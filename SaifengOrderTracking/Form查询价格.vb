Public Class Form查询价格
    Private Sub Form51_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form51_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("ID", "ID")
        DataGridView1.Columns.Add("CustCode", "客户代码")
        DataGridView1.Columns.Add("CDate", "下单日期")
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("CustOrder", "客户订单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("PartName", "零件名称")
        DataGridView1.Columns.Add("Qty", "数量")
        DataGridView1.Columns.Add("UnitP", "单价")
        DataGridView1.Columns.Add("TotalMoney", "总价")
        DataGridView1.Columns.Add("DDate", "交期")
        For i = 0 To DataGridView1.Columns.Count - 1
            DataGridView1.Columns(i).ReadOnly = True
        Next
        DataGridView1.Columns("UnitP").ReadOnly = False
        If UserName = "admin" Then
            DataGridView1.Columns("Qty").ReadOnly = False
            DataGridView1.Columns("DrawNo").ReadOnly = False
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            On Error Resume Next
            DataGridView1.Rows.Clear()
            If TextBox1.Text = "" Then
                MsgBox("请输入单号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String

            SQLstring = "Select ID,CustOrder,SFOrder,DrawNo,PartName,Qty,UnitP,CustCode,CDate,DDate from SFOrderBase where SFOrder=" & "'" & TextBox1.Text & "'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此单号不存在！")
                Exit Sub
            End If
            rs.MoveFirst()
            Dim TempRow As Integer
            TempRow = 0
            Do While Not rs.EOF
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(rs(i).Name, TempRow).Value = rs.Fields(i).Value
                    DataGridView1.Item("TotalMoney", TempRow).Value = IIf(IsNumeric(rs("UnitP").Value), rs("UnitP").Value, 0) * rs("Qty").Value
                    If (rs(i).Name = "CDate") Or (rs(i).Name = "DDate") Then
                        DataGridView1.Item(rs(i).Name, TempRow).Value = Year(rs.Fields(i).Value) & "/" & Month(rs.Fields(i).Value) & "/" & DatePart(DateInterval.Day, rs.Fields(i).Value)
                    End If
                Next
                TempRow = TempRow + 1
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
            On Error Resume Next
            DataGridView1.Rows.Clear()
            If TextBox2.Text = "" Then
                MsgBox("请输入图号！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String

            SQLstring = "Select ID,CustOrder,SFOrder,DrawNo,PartName,Qty,UnitP,CustCode,CDate,DDate from SFOrderBase where DrawNo like " & "'%" & TextBox2.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此图号不存在！")
                Exit Sub
            End If
            rs.MoveFirst()
            Dim TempRow As Integer
            TempRow = 0
            Do While Not rs.EOF
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(rs(i).Name, TempRow).Value = rs.Fields(i).Value
                    DataGridView1.Item("TotalMoney", TempRow).Value = IIf(IsNumeric(rs("UnitP").Value), rs("UnitP").Value, 0) * rs("Qty").Value
                    If (rs(i).Name = "CDate") Or (rs(i).Name = "DDate") Then
                        DataGridView1.Item(rs(i).Name, TempRow).Value = Year(rs.Fields(i).Value) & "/" & Month(rs.Fields(i).Value) & "/" & DatePart(DateInterval.Day, rs.Fields(i).Value)
                    End If
                Next
                TempRow = TempRow + 1
                rs.MoveNext()
            Loop

            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing

        End If
    End Sub

    Private Sub Form51_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        DataGridView1.Width = Me.Width - 100
        DataGridView1.Height = Me.Height - 150
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = 13 Then
            On Error Resume Next
            DataGridView1.Rows.Clear()
            If TextBox3.Text = "" Then
                MsgBox("请输入零件名称！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            Dim SQLstring As String

            SQLstring = "Select ID,CustOrder,SFOrder,DrawNo,PartName,Qty,UnitP,CustCode,CDate,DDate from SFOrderBase where PartName like " & "'%" & TextBox3.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("此零件名称不存在！")
                Exit Sub
            End If
            rs.MoveFirst()
            Dim TempRow As Integer
            TempRow = 0
            Do While Not rs.EOF
                DataGridView1.Rows.Add(1)
                For i = 0 To rs.Fields.Count - 1
                    DataGridView1.Item(rs(i).Name, TempRow).Value = rs.Fields(i).Value
                    DataGridView1.Item("TotalMoney", TempRow).Value = IIf(IsNumeric(rs("UnitP").Value), rs("UnitP").Value, 0) * rs("Qty").Value
                    If (rs(i).Name = "CDate") Or (rs(i).Name = "DDate") Then
                        DataGridView1.Item(rs(i).Name, TempRow).Value = Year(rs.Fields(i).Value) & "/" & Month(rs.Fields(i).Value) & "/" & DatePart(DateInterval.Day, rs.Fields(i).Value)
                    End If
                Next
                TempRow = TempRow + 1
                rs.MoveNext()
            Loop

            rs.Close()
            cn.Close()
            rs = Nothing
            cn = Nothing

        End If
    End Sub
End Class