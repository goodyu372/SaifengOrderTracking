Public Class Form删除出货记录

    Private Sub Form69_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form49_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form49_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("ID", "序号")
        DataGridView1.Columns.Add("CustCode", "客户代码")
        DataGridView1.Columns.Add("CustOrder", "客户订单号")
        DataGridView1.Columns.Add("SFOrder", "赛峰单号")
        DataGridView1.Columns.Add("DrawNo", "图号")
        DataGridView1.Columns.Add("PartName", "零件名称")
        DataGridView1.Columns.Add("OutputQty", "出货数量")
        DataGridView1.Columns.Add("UnitP", "单价")
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            '按生产单号查询
            On Error Resume Next
            DataGridView1.Rows.Clear()
            'DataGridView1.Columns.Clear()

            If TextBox1.Text = "" Then
                MsgBox("生产单号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring, SQLOutput As String
            Dim ID As Integer
            'SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,Status from SFOrderBase where SFOrder=" & "'" & TextBox1.Text & "' and ID not in (Select ID from TotalOutput where TotalOut>=Qty)"
            SQLstring = "Select * from OutputRecord where SFOrder=" & "'" & TextBox1.Text & "'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("不存在此生产单号的出货！")
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
            cn.Close()
            rs = Nothing
            cn = Nothing

        End If
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        Dim cn As New ADODB.Connection
        Dim rs, rsOutput As New ADODB.Recordset
        Dim SQLstring As String
        SQLstring = "delete from OutputRecord where ID=" & "'" & DataGridView1.Item("ID", e.RowIndex).Value & "'"
        cn.Open(CNsfmdb)
        rs.Open(SQLstring, cn, 1, 3)
        cn.Close()
        rs = Nothing
        cn = Nothing
        MsgBox("Output record deleted !")
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = 13 Then
            '按图号查询
            On Error Resume Next
            DataGridView1.Rows.Clear()

            If TextBox2.Text = "" Then
                MsgBox("图号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring, SQLOutput As String
            Dim ID As Integer
            'SQLstring = "Select CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,Status from SFOrderBase where DrawNo like " & "'%" & TextBox2.Text & "%' and Status<>'完成入库'"
            'SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,Status from SFOrderBase where DrawNo like " & "'" & TextBox2.Text & "' and ID not in (Select ID from TotalOutput where TotalOut>=Qty)"
            'SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,TotalOut from TotalOutput where DrawNo like " & "'%" & TextBox2.Text & "%' and TotalOut>0"
            SQLstring = "Select * from OutputRecord where DrawNo like " & "'%" & TextBox2.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("不存在该图号的出货！")
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
            cn.Close()
            rs = Nothing
            cn = Nothing

        End If
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = 13 Then
            '按客户订单号查询
            On Error Resume Next
            DataGridView1.Rows.Clear()
            If TextBox3.Text = "" Then
                MsgBox("客户订单号不能为空！")
                Exit Sub
            End If
            Dim cn As New ADODB.Connection
            Dim rs, rsOutput As New ADODB.Recordset
            Dim SQLstring, SQLOutput As String
            Dim ID As Integer
            'SQLstring = "Select CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,Status from SFOrderBase where CustOrder like " & "'%" & TextBox3.Text & "%' and Status<>'完成入库'"
            'SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,Status from SFOrderBase where CustOrder like " & "'" & TextBox3.Text & "' and ID not in (Select ID from TotalOutput where TotalOut>=Qty)"
            'SQLstring = "Select ID,CustOrder,CustCode,SFOrder,DrawNo,PartName,Qty,UnitP,TotalOut from TotalOutput where CustOrder like " & "'%" & TextBox3.Text & "%' and TotalOut>0"
            SQLstring = "Select * from OutputRecord where where CustOrder like " & "'%" & TextBox3.Text & "%'"
            cn.Open(CNsfmdb)
            rs.Open(SQLstring, cn, 1, 1)
            If rs.RecordCount = 0 Then
                rs.Close()
                cn.Close()
                rs = Nothing
                cn = Nothing
                MsgBox("没有该客户订单号的出货！")
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
            cn.Close()
            rs = Nothing
            cn = Nothing

        End If
    End Sub
End Class