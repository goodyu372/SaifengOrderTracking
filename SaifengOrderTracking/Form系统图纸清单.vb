Public Class Form系统图纸清单

    Private Sub Form62_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form62_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '**************************** 将Excel流程卡内容写入到Datagridview1 ************************************
        Dim FileType As String
        Dim sPath As String
        Dim Folder_name As String
        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("FileName", "FileName")
        Dim a As Integer
        Folder_name = DWGPath
        FileType = "*.*"
        sPath = Dir(Folder_name & FileType) '查找第一个文件  

        Do While Len(sPath) '循环到没有文件为止  
            a = a + 1
            sPath = Dir() '查找下一个文件  
            DataGridView1.Rows.Add(1)
            DataGridView1.Item(0, DataGridView1.Rows.Count - 1).Value = Mid(sPath, 1, Len(sPath) - 4)
        Loop
    End Sub
End Class