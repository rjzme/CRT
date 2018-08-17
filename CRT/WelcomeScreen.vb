Public NotInheritable Class WelcomeScreen

    'TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
    '  of the Project Designer ("Properties" under the "Project" menu).


    Private Sub WelcomeScreen_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set up the dialog text at runtime according to the application's assembly information.  

        'TODO: Customize the application's assembly information in the "Application" pane of the project 
        '  properties dialog (under the "Project" menu).



        'Format the version information using the text set into the Version control at design time as the
        '  formatting string.  This allows for effective localization if desired.
        '  Build and revision information could be included by using the following code and changing the 
        '  Version control's designtime text to "Version {0}.{1:00}.{2}.{3}" or something similar.  See
        '  String.Format() in Help for more information.
        '
        '    Version.Text = System.String.Format(Version.Text, My.Application.Info.Version.Major, My.Application.Info.Version.Minor, My.Application.Info.Version.Build, My.Application.Info.Version.Revision)
        Timer1.Enabled = True

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        MetroProgressBar1.Visible = True
        If MetroProgressBar1.Value <= 97 Then

            MetroProgressBar1.Value = MetroProgressBar1.Value + 3
            If MetroProgressBar1.Value = 12 Then
                lblProgessInfo.Text = "Initialising ..."
            End If

            If MetroProgressBar1.Value = 60 Then
                lblProgessInfo.Text = "Loading data ..."
            End If

        End If
    End Sub
End Class
