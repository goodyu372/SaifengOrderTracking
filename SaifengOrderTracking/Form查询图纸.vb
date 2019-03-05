Public Class Form查询图纸

    Private Sub Form73_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form73_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("FileName", "文件名")
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = 13 Then
            If TextBox1.Text = "" Then
                MsgBox("请输入图号！")
                Exit Sub
            End If
            DataGridView1.Rows.Clear()
            Dim cn As New ADODB.Connection
            Dim rs As New ADODB.Recordset
            cn.Open(CNsfmdb)
            Shell(SFOrderBaseNetConnect, vbHide)
            '88888888888888888888888888 Select DrawNo to display 88888888888888888888888888888
            Dim sF As String
            Dim Path, SearchedFile As String
            Path = DWGPath

            SearchedFile = TextBox1.Text

            sF = Dir(Path, vbDirectory) ' 查找目录中第一个文件夹名称
            Do While sF <> "" ' 跳过当前的目录及上层目录
                If (InStr(1, UCase(sF), UCase(SearchedFile)) > 0) Then
                    DataGridView1.Rows.Add(1)
                    DataGridView1.Item(0, DataGridView1.Rows.Count - 1).Value = sF
                End If
                sF = Dir()
            Loop
        End If
    End Sub
    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex >= 0 Then
            Dim sF As String
            Dim Path As String
            Path = DWGPath
            sF = DataGridView1.Item(0, e.RowIndex).Value
            System.Diagnostics.Process.Start(Path + sF)
        End If
    End Sub
End Class