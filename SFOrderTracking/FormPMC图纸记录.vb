Public Class FormPMC图纸记录
    Public DrawingReceiveStatus As String 'by order, "Y" stands for received, "N" stands for not received
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If (e.KeyCode = 13) Then
            Select Case DrawingReceiveStatus
                Case "Receive"
                    Dim cn As New ADODB.Connection
                    Dim rs, rs1 As New ADODB.Recordset
                    Dim SQLstring, CNString As String
                    On Error Resume Next
                    Dim CustCode As String
                    Dim QueryText As String

                    cn.Open(CNsfmdb)
                    QueryText = "Select SFOrder,DrawNo,CustCode,CDate,DDate,Status from SFOrderBase where (SFOrder=" & "'" & TextBox1.Text & "'" & ") order by DrawNo"
                    'DataGridView1.Visible = False
                    'rs.Open(QueryText, cn, 1, 2)
                    rs.Open(QueryText, cn, 1, 1)
                    StatusStrip1.Items(0).Text = "记录条数： " & rs.RecordCount
                    If (rs.RecordCount = 0) Then
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        MsgBox("该生产单号不存在，请核对！")
                        Exit Sub
                    End If
                    '*********************************************** First check PMC_DrawingRecord records ***************************************************************
                    QueryText = "Select * from PMC_DrawingRecord where (SFOrder=" & "'" & TextBox1.Text & "'" & ") order by DrawNo"
                    'rs1.Open(QueryText, cn, 1, 2)
                    rs1.Open(QueryText, cn, 1, 1)
                    'StatusStrip1.Items(2).Text = "记录条数2： " & rs1.RecordCount
                    If (rs1.RecordCount = 0) Then
                        DataGridView2.Rows.Clear()
                        rs.MoveFirst()
                        'For i = 0 To rs.Fields.Count - 1
                        'DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
                        'Next
                        Do While rs.EOF = False
                            DataGridView2.Rows.Add(1)
                            For i = 0 To rs.Fields.Count - 1
                                DataGridView2.Item("SFOrder", DataGridView2.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
                                DataGridView2.Item("DrawNo", DataGridView2.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
                            Next
                            rs.MoveNext()
                        Loop
                        rs.Close()
                        rs1.Close()
                        cn.Close()
                        rs1 = Nothing
                        rs = Nothing
                        cn = Nothing
                        Exit Sub
                    Else
                        DataGridView2.Rows.Clear()
                        'DataGridView1.Visible = False
                        rs1.MoveFirst()
                        Do While rs1.EOF = False
                            DataGridView2.Rows.Add(1)
                            For i = 0 To rs1.Fields.Count - 1
                                DataGridView2.Item(i, DataGridView2.Rows.Count - 1).Value = rs1.Fields(i).Value
                            Next
                            rs1.MoveNext()
                        Loop
                    End If
                    rs.Close()
                    rs1.Close()
                    cn.Close()
                    rs1 = Nothing
                    rs = Nothing
                    cn = Nothing
                    '*****************************************************************************************************************************************************
                Case "Release"
                    Dim cn As New ADODB.Connection
                    Dim rs, rs1 As New ADODB.Recordset
                    Dim SQLstring, CNString As String
                    On Error Resume Next
                    Dim CustCode As String
                    Dim QueryText As String

                    cn.Open(CNsfmdb)
                    QueryText = "Select SFOrder,DrawNo,CustCode,CDate,DDate,Status from SFOrderBase where (SFOrder=" & "'" & TextBox1.Text & "'" & ") order by DrawNo"
                    'DataGridView1.Visible = False
                    'rs.Open(QueryText, cn, 1, 2)
                    rs.Open(QueryText, cn, 1, 1)
                    If (rs.RecordCount = 0) Then
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        MsgBox("该生产单号不存在，请核对！")
                        Exit Sub
                    End If
                    '*********************************************** First check PMC_DrawingRecord records ***************************************************************
                    QueryText = "Select * from PMC_DrawingRecord where (SFOrder=" & "'" & TextBox1.Text & "'" & ") order by DrawNo"
                    'rs1.Open(QueryText, cn, 1, 2)
                    rs1.Open(QueryText, cn, 1, 1)
                    If (rs1.RecordCount = 0) Then
                        DataGridView2.Rows.Clear()
                        rs.Close()
                        rs1.Close()
                        cn.Close()
                        rs1 = Nothing
                        rs = Nothing
                        cn = Nothing
                        MsgBox("该单号尚未收图，请核查！")
                        Exit Sub
                    Else
                        DataGridView2.Rows.Clear()
                        'DataGridView1.Visible = False
                        rs1.MoveFirst()
                        Do While rs1.EOF = False
                            DataGridView2.Rows.Add(1)
                            For i = 0 To rs1.Fields.Count - 1
                                DataGridView2.Item(i, DataGridView2.Rows.Count - 1).Value = rs1.Fields(i).Value
                            Next
                            rs1.MoveNext()
                        Loop
                    End If
                    rs.Close()
                    rs1.Close()
                    cn.Close()
                    rs1 = Nothing
                    rs = Nothing
                    cn = Nothing
                    '*****************************************************************************************************************************************************
                Case "Return"
                    Dim cn As New ADODB.Connection
                    Dim rs, rs1 As New ADODB.Recordset
                    Dim SQLstring, CNString As String
                    On Error Resume Next
                    Dim CustCode As String
                    Dim QueryText As String

                    cn.Open(CNsfmdb)
                    QueryText = "Select SFOrder,DrawNo,CustCode,CDate,DDate,Status from SFOrderBase where (SFOrder=" & "'" & TextBox1.Text & "'" & ") order by DrawNo"
                    'DataGridView1.Visible = False
                    'rs.Open(QueryText, cn, 1, 2)
                    rs.Open(QueryText, cn, 1, 1)
                    If (rs.RecordCount = 0) Then
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        MsgBox("该生产单号不存在，请核对！")
                        Exit Sub
                    End If
                    '*********************************************** First check PMC_DrawingRecord records ***************************************************************
                    QueryText = "Select * from PMC_DrawingRecord where (SFOrder=" & "'" & TextBox1.Text & "'" & ") order by DrawNo"
                    'rs1.Open(QueryText, cn, 1, 2)
                    rs1.Open(QueryText, cn, 1, 1)
                    If (rs1.RecordCount = 0) Then
                        DataGridView2.Rows.Clear()
                        rs.Close()
                        rs1.Close()
                        cn.Close()
                        rs1 = Nothing
                        rs = Nothing
                        cn = Nothing
                        MsgBox("该单号尚未收图，请核查！")
                        Exit Sub
                    Else
                        DataGridView2.Rows.Clear()
                        'DataGridView1.Visible = False
                        rs1.MoveFirst()
                        Do While rs1.EOF = False
                            DataGridView2.Rows.Add(1)
                            For i = 0 To rs1.Fields.Count - 1
                                DataGridView2.Item(i, DataGridView2.Rows.Count - 1).Value = rs1.Fields(i).Value
                            Next
                            rs1.MoveNext()
                        Loop
                    End If
                    rs.Close()
                    rs1.Close()
                    cn.Close()
                    rs1 = Nothing
                    rs = Nothing
                    cn = Nothing
                    '*****************************************************************************************************************************************************
            End Select
        End If
    End Sub

    Private Sub Form15_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form15_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView2.Columns.Add("SFOrder", "赛峰单号")
        DataGridView2.Columns.Add("DrawNo", "图号")
        DataGridView2.Columns.Add("ReceiveDate", "收图日期")
        DataGridView2.Columns.Add("ReleaseDate", "给工艺日期")
        DataGridView2.Columns.Add("ReturnDate", "工艺交回日期")
        DataGridView2.Columns.Add("ProcMaker", "编工艺者")
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentDoubleClick
        On Error Resume Next
        Select Case DrawingReceiveStatus
            Case "Receive"
                If (IsDBNull(DataGridView2.Item("ReceiveDate", e.RowIndex).Value)) Then
                    DataGridView2.Item("ReceiveDate", e.RowIndex).Value = "V"
                End If
                If (DataGridView2.Item("ReceiveDate", e.RowIndex).Value = Nothing) Then
                    DataGridView2.Item("ReceiveDate", e.RowIndex).Value = "V"
                End If
            Case "Release"
                If (IsDBNull(DataGridView2.Item("ProcMaker", e.RowIndex).Value)) Then
                    DataGridView2.Item("ProcMaker", e.RowIndex).Value = ComboBox1.SelectedItem.ToString
                End If
                If (DataGridView2.Item("ProcMaker", e.RowIndex).Value = Nothing) Then
                    DataGridView2.Item("ProcMaker", e.RowIndex).Value = ComboBox1.SelectedItem.ToString
                End If
            Case "Return"
                If (IsDBNull(DataGridView2.Item("ReturnDate", e.RowIndex).Value)) Then
                    DataGridView2.Item("ReturnDate", e.RowIndex).Value = "V"
                End If
                If (DataGridView2.Item("ReturnDate", e.RowIndex).Value = Nothing) Then
                    DataGridView2.Item("ReturnDate", e.RowIndex).Value = "V"
                End If
        End Select
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If (DataGridView2.Rows.Count = 0) Then
            MsgBox("没有收图信息！")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring, CNString As String
        On Error Resume Next
        Dim CustCode As String
        Dim QueryText As String
        QueryText = "Select * from PMC_DrawingRecord where (SFOrder=" & "'" & TextBox1.Text & "'" & ") order by DrawNo"
        cn.Open(CNsfmdb)
        'rs.Open(QueryText, cn, 1, 2)
        rs.Open(QueryText, cn, 1, 3)
        If (rs.RecordCount = 0) Then
            For i = 0 To DataGridView2.Rows.Count - 1
                rs.AddNew()
                rs.Fields("SFOrder").Value = DataGridView2.Item("SFOrder", i).Value
                rs.Fields("DrawNo").Value = DataGridView2.Item("DrawNo", i).Value
                rs.Fields("ReceiveDate").Value = IIf(DataGridView2.Item("ReceiveDate", i).Value = "V", Now.Date, DataGridView2.Item("ReceiveDate", i).Value)
                rs.Update()
            Next
        Else
            For i = 0 To DataGridView2.Rows.Count - 1
                'If ((Not (IsDBNull(DataGridView2.Item("ReceiveDate", i).Value))) And (DataGridView2.Item("ReceiveDate", i).Value = "V")) Then
                If (DataGridView2.Item("ReceiveDate", i).Value = Nothing) Then

                Else
                    If (DataGridView2.Item("ReceiveDate", i).Value = "V") Then
                        rs.Close()
                        QueryText = "Select * from PMC_DrawingRecord where ((SFOrder=" & "'" & (DataGridView2.Item("SFOrder", i).Value) & "'" & ") and (DrawNo=" & "'" & (DataGridView2.Item("DrawNo", i).Value) & "'))"
                        'rs.Open(QueryText, cn, 1, 2)
                        rs.Open(QueryText, cn, 1, 3)
                        rs.MoveFirst()
                        rs.Fields("ReceiveDate").Value = Now
                        rs.Update()
                    End If
                End If
            Next
        End If

        rs.Close()
        cn.Close()
        DataGridView2.Rows.Clear()
        rs = Nothing
        cn = Nothing
        MsgBox("完成收图记录！")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (DataGridView2.Rows.Count = 0) Then
            MsgBox("没有发图信息！")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring, CNString As String
        On Error Resume Next
        Dim CustCode As String
        Dim QueryText As String
        QueryText = "Select * from PMC_DrawingRecord where (SFOrder=" & "'" & TextBox1.Text & "'" & ") order by DrawNo"
        cn.Open(CNsfmdb)
        'rs.Open(QueryText, cn, 1, 2)
        rs.Open(QueryText, cn, 1, 3)
        For i = 0 To DataGridView2.Rows.Count - 1
            If (IsDBNull(DataGridView2.Item("ProcMaker", i).Value)) Then
            Else
                If (Len(DataGridView2.Item("ProcMaker", i).Value) <> 0) Then
                    rs.Close()
                    QueryText = "Select * from PMC_DrawingRecord where ((SFOrder=" & "'" & (DataGridView2.Item("SFOrder", i).Value) & "'" & ") and (DrawNo=" & "'" & (DataGridView2.Item("DrawNo", i).Value) & "'))"
                    'rs.Open(QueryText, cn, 1, 2)
                    rs.Open(QueryText, cn, 1, 3)
                    rs.MoveFirst()
                    rs.Fields("ProcMaker").Value = DataGridView2.Item("ProcMaker", i).Value
                    rs.Fields("ReleaseDate").Value = Now
                    rs.Update()
                End If
            End If
        Next

        rs.Close()
        cn.Close()
        DataGridView2.Rows.Clear()
        rs = Nothing
        cn = Nothing
        MsgBox("完成发图记录！")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If (DataGridView2.Rows.Count = 0) Then
            MsgBox("没有回图信息！")
            Exit Sub
        End If
        Dim cn As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim SQLstring, CNString As String
        On Error Resume Next
        Dim CustCode As String
        Dim QueryText As String
        QueryText = "Select * from PMC_DrawingRecord where (SFOrder=" & "'" & TextBox1.Text & "'" & ") order by DrawNo"
        cn.Open(CNsfmdb)
        'rs.Open(QueryText, cn, 1, 2)
        rs.Open(QueryText, cn, 1, 3)
        For i = 0 To DataGridView2.Rows.Count - 1
            If (IsDBNull(DataGridView2.Item("ReturnDate", i).Value)) Then
            Else
                If (Len(DataGridView2.Item("ReturnDate", i).Value) <> 0) Then
                    rs.Close()
                    QueryText = "Select * from PMC_DrawingRecord where ((SFOrder=" & "'" & (DataGridView2.Item("SFOrder", i).Value) & "'" & ") and (DrawNo=" & "'" & (DataGridView2.Item("DrawNo", i).Value) & "'))"
                    'rs.Open(QueryText, cn, 1, 2)
                    rs.Open(QueryText, cn, 1, 3)
                    rs.MoveFirst()
                    rs.Fields("ReturnDate").Value = Now
                    rs.Update()
                End If
            End If
        Next

        rs.Close()
        cn.Close()
        DataGridView2.Rows.Clear()
        rs = Nothing
        cn = Nothing
        MsgBox("完成回图记录！")
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If (e.KeyCode = 13) Then
            Select Case DrawingReceiveStatus
                Case "Receive"
                    Dim cn As New ADODB.Connection
                    Dim rs, rs1 As New ADODB.Recordset
                    Dim SQLstring, CNString As String
                    'On Error Resume Next
                    Dim CustCode As String
                    Dim QueryText As String

                    cn.Open(CNsfmdb)
                    QueryText = "Select SFOrder,DrawNo,CustCode,CDate,DDate,Status from SFOrderBase where (DrawNo like " & "'%" & TextBox2.Text & "%'" & ")"
                    'DataGridView1.Visible = False
                    'rs.Open(QueryText, cn, 1, 2)
                    rs.Open(QueryText, cn, 1, 1)
                    If (rs.RecordCount = 0) Then
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        MsgBox("该图号不存在，请核对！")
                        Exit Sub
                    End If
                    '*********************************************** First check PMC_DrawingRecord records ***************************************************************
                    QueryText = "Select * from PMC_DrawingRecord where (DrawNo like " & "'%" & TextBox2.Text & "%')"
                    'rs1.Open(QueryText, cn, 1, 2)
                    rs1.Open(QueryText, cn, 1, 1)
                    StatusStrip1.Items(0).Text = "记录条数： " & rs1.RecordCount
                    If (rs1.RecordCount = 0) Then
                        DataGridView2.Rows.Clear()
                        rs.MoveFirst()
                        'For i = 0 To rs.Fields.Count - 1
                        'DataGridView1.Columns.Add(rs.Fields(i).Name, rs.Fields(i).Name)
                        'Next
                        Do While rs.EOF = False
                            DataGridView2.Rows.Add(1)
                            For i = 0 To rs.Fields.Count - 1
                                DataGridView2.Item("SFOrder", DataGridView2.Rows.Count - 1).Value = rs.Fields("SFOrder").Value
                                DataGridView2.Item("DrawNo", DataGridView2.Rows.Count - 1).Value = rs.Fields("DrawNo").Value
                            Next
                            rs.MoveNext()
                        Loop
                        rs.Close()
                        rs1.Close()
                        cn.Close()
                        rs1 = Nothing
                        rs = Nothing
                        cn = Nothing
                        Exit Sub
                    Else
                        DataGridView2.Rows.Clear()
                        'DataGridView1.Visible = False
                        rs1.MoveFirst()
                        Do While rs1.EOF = False
                            DataGridView2.Rows.Add(1)
                            For i = 0 To rs1.Fields.Count - 1
                                DataGridView2.Item(i, DataGridView2.Rows.Count - 1).Value = rs1.Fields(i).Value
                            Next
                            rs1.MoveNext()
                        Loop
                    End If
                    rs.Close()
                    rs1.Close()
                    cn.Close()
                    rs1 = Nothing
                    rs = Nothing
                    cn = Nothing
                    '*****************************************************************************************************************************************************
                Case "Release"
                    Dim cn As New ADODB.Connection
                    Dim rs, rs1 As New ADODB.Recordset
                    Dim SQLstring, CNString As String
                    On Error Resume Next
                    Dim CustCode As String
                    Dim QueryText As String

                    cn.Open(CNsfmdb)
                    QueryText = "Select SFOrder,DrawNo,CustCode,CDate,DDate,Status from SFOrderBase where (DrawNo like " & "'%" & TextBox2.Text & "%')"
                    'DataGridView1.Visible = False
                    'rs.Open(QueryText, cn, 1, 2)
                    rs.Open(QueryText, cn, 1, 1)
                    If (rs.RecordCount = 0) Then
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        MsgBox("该图号不存在，请核对！")
                        Exit Sub
                    End If
                    '*********************************************** First check PMC_DrawingRecord records ***************************************************************
                    QueryText = "Select * from PMC_DrawingRecord where (DrawNo like " & "'%" & TextBox2.Text & "%')"
                    'rs1.Open(QueryText, cn, 1, 2)
                    rs1.Open(QueryText, cn, 1, 1)
                    StatusStrip1.Items(0).Text = "记录条数： " & rs1.RecordCount
                    If (rs1.RecordCount = 0) Then
                        DataGridView2.Rows.Clear()
                        rs.Close()
                        rs1.Close()
                        cn.Close()
                        rs1 = Nothing
                        rs = Nothing
                        cn = Nothing
                        MsgBox("该图号尚未收图，请核查！")
                        Exit Sub
                    Else
                        DataGridView2.Rows.Clear()
                        'DataGridView1.Visible = False
                        rs1.MoveFirst()
                        Do While rs1.EOF = False
                            DataGridView2.Rows.Add(1)
                            For i = 0 To rs1.Fields.Count - 1
                                DataGridView2.Item(i, DataGridView2.Rows.Count - 1).Value = rs1.Fields(i).Value
                            Next
                            rs1.MoveNext()
                        Loop
                    End If
                    rs.Close()
                    rs1.Close()
                    cn.Close()
                    rs1 = Nothing
                    rs = Nothing
                    cn = Nothing
                    '*****************************************************************************************************************************************************
                Case "Return"
                    Dim cn As New ADODB.Connection
                    Dim rs, rs1 As New ADODB.Recordset
                    Dim SQLstring, CNString As String
                    On Error Resume Next
                    Dim CustCode As String
                    Dim QueryText As String

                    cn.Open(CNsfmdb)
                    QueryText = "Select SFOrder,DrawNo,CustCode,CDate,DDate,Status from SFOrderBase where (DrawNo like " & "'%" & TextBox2.Text & "%')"
                    'DataGridView1.Visible = False
                    'rs.Open(QueryText, cn, 1, 2)
                    rs.Open(QueryText, cn, 1, 1)
                    If (rs.RecordCount = 0) Then
                        rs.Close()
                        cn.Close()
                        rs = Nothing
                        cn = Nothing
                        MsgBox("该图号不存在，请核对！")
                        Exit Sub
                    End If
                    '*********************************************** First check PMC_DrawingRecord records ***************************************************************
                    QueryText = "Select * from PMC_DrawingRecord where (DrawNo like " & "'%" & TextBox2.Text & "%')"
                    'rs1.Open(QueryText, cn, 1, 2)
                    rs1.Open(QueryText, cn, 1, 1)
                    StatusStrip1.Items(0).Text = "记录条数： " & rs1.RecordCount
                    If (rs1.RecordCount = 0) Then
                        DataGridView2.Rows.Clear()
                        rs.Close()
                        rs1.Close()
                        cn.Close()
                        rs1 = Nothing
                        rs = Nothing
                        cn = Nothing
                        MsgBox("该图号尚未收图，请核查！")
                        Exit Sub
                    Else
                        DataGridView2.Rows.Clear()
                        'DataGridView1.Visible = False
                        rs1.MoveFirst()
                        Do While rs1.EOF = False
                            DataGridView2.Rows.Add(1)
                            For i = 0 To rs1.Fields.Count - 1
                                DataGridView2.Item(i, DataGridView2.Rows.Count - 1).Value = rs1.Fields(i).Value
                            Next
                            rs1.MoveNext()
                        Loop
                    End If
                    rs.Close()
                    rs1.Close()
                    cn.Close()
                    rs1 = Nothing
                    rs = Nothing
                    cn = Nothing
                    '*****************************************************************************************************************************************************
            End Select
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        On Error Resume Next
        Select Case DrawingReceiveStatus
            Case "Receive"
                For i = 0 To DataGridView2.Rows.Count - 1
                    DataGridView2.Item("ReceiveDate", i).Value = "V"
                Next
        End Select
    End Sub
End Class