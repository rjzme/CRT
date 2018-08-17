Imports System.IO

Public Class FrmCRT
    Sub main()
        Connection()
        MdlBasic.main()
    End Sub

    Private Sub BtnStart_Click(sender As Object, e As EventArgs) Handles BtnStart.Click
        Try
            main()
        Catch ex As Exception
            MsgBox(ex.StackTrace)
        End Try
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles BtnClear.Click

        DataGridView1.DataSource = Nothing

    End Sub

    Private Sub FolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FolderToolStripMenuItem.Click
        Dim OpenFileDialog As New OpenFileDialog
        Dim excelFile As String
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "All Files (*.*)|*.*|Excel files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|XLS Files (*.xls)|*xls"

        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then

            Dim fi As New FileInfo(OpenFileDialog.FileName)
            Dim FileName As String = OpenFileDialog.FileName


            excelFile = fi.FullName
        Else

            Exit Sub

        End If

        ImportExcel(excelFile)
    End Sub


End Class



