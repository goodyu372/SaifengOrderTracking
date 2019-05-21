Public Class Form37

    Private Sub Form37_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form主窗体.Show()
    End Sub

    Private Sub Form37_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.Filter = "bmp files (*.bmp)|*.bmp|pdf files (*.pdf)|*.pdf"
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        ListBox1.Items.Clear()
        For i = 0 To OpenFileDialog1.FileNames.Count - 1
            ListBox1.Items.Add(System.IO.Path.GetFullPath(OpenFileDialog1.FileNames(i)))
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox1.Items.Count = 0 Then
            MsgBox("没有选择要录入的图纸！")
            Exit Sub
        End If
        On Error Resume Next
        Dim CopyPath As String
        CopyPath = DWGPath
        For i = 0 To ListBox1.Items.Count - 1
            If My.Computer.FileSystem.FileExists(CopyPath & System.IO.Path.GetFileName(ListBox1.Items(i).ToString)) Then
                IO.File.Delete(CopyPath & System.IO.Path.GetFileName(ListBox1.Items(i).ToString))
            End If
            IO.File.Copy(ListBox1.Items(i).ToString, CopyPath & System.IO.Path.GetFileName(ListBox1.Items(i).ToString))
            ListBox1.Items(i).ToString()
        Next
        ListBox1.Items.Clear()
        MsgBox("成功录入图纸！")
    End Sub
End Class