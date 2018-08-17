
Imports System.IO
Imports System
Imports System.Threading

Public Class MainForm

    Private Sub MetroLabelClose_Click(sender As Object, e As EventArgs) Handles MetroLabelClose.Click
        End
    End Sub

    Sub main()
        Connection()

        MdlBasic.main()
    End Sub

    Private Sub MetroButtonStart_Click(sender As Object, e As EventArgs) Handles MetroButtonStart.Click

        Try

            main()
        Catch ex As Exception
            MsgBox(ex.StackTrace)
        End Try
    End Sub

    Private Sub ImportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportToolStripMenuItem.Click
        Dim OpenFileDialog As New OpenFileDialog

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
        MetroTabControlMain.SelectedTab = MetroTabPageMyWork

    End Sub



    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        End
    End Sub

    Private Sub MetroLabelMax_Click(sender As Object, e As EventArgs) Handles MetroLabelMax.Click
        If Me.WindowState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Normal
        Else
            Me.WindowState = FormWindowState.Maximized
        End If
    End Sub

    Private Sub MetroLabelMin_Click(sender As Object, e As EventArgs) Handles MetroLabelMin.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub MetroLabelClose_MouseEnter(sender As Object, e As EventArgs) Handles MetroLabelClose.MouseEnter
        Me.MetroLabelClose.BackColor = Color.DimGray
    End Sub

    Private Sub MetroLabelMax_MouseEnter(sender As Object, e As EventArgs) Handles MetroLabelMax.MouseEnter
        Me.MetroLabelMax.BackColor = Color.DimGray
    End Sub

    Private Sub MetroLabelMin_MouseEnter(sender As Object, e As EventArgs) Handles MetroLabelMin.MouseEnter
        Me.MetroLabelMin.BackColor = Color.DimGray
    End Sub

    Private Sub MetroLabelMin_MouseLeave(sender As Object, e As EventArgs) Handles MetroLabelMin.MouseLeave
        Me.MetroLabelMin.BackColor = Color.Transparent
    End Sub

    Private Sub MetroLabelMax_MouseLeave(sender As Object, e As EventArgs) Handles MetroLabelMax.MouseLeave
        Me.MetroLabelMax.BackColor = Color.Transparent
    End Sub

    Private Sub MetroLabelClose_MouseLeave(sender As Object, e As EventArgs) Handles MetroLabelClose.MouseLeave
        Me.MetroLabelClose.BackColor = Color.Transparent
    End Sub

    Private IsFormBeingDragged As Boolean = False

    Private MouseDownX As Integer

    Private MouseDownY As Integer



    Private Sub frmDashBoard_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MetroPanelNavBar.MouseDown



        If e.Button = MouseButtons.Left Then

            IsFormBeingDragged = True

            MouseDownX = e.X

            MouseDownY = e.Y

        End If

    End Sub



    Private Sub frmDashBoard_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MetroPanelNavBar.MouseUp



        If e.Button = MouseButtons.Left Then

            IsFormBeingDragged = False

        End If

    End Sub



    Private Sub frmDashBoard_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MetroPanelNavBar.MouseMove



        If IsFormBeingDragged Then

            Dim temp As Point = New Point()



            temp.X = Me.Location.X + (e.X - MouseDownX)

            temp.Y = Me.Location.Y + (e.Y - MouseDownY)

            Me.Location = temp

            temp = Nothing

        End If

    End Sub

    Private Sub MetroTabPageUnetInfo_Click(sender As Object, e As EventArgs) Handles MetroTabPageUnetInfo.Click

    End Sub

    Private Sub SplitContainerControlPanel_Panel2_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainerControlPanel.Panel2.Paint

    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub MetroPanelRightTop_Paint(sender As Object, e As PaintEventArgs) Handles MetroPanelRightTop.Paint

    End Sub




End Class
