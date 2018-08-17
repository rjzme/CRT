<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WelcomeScreen
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.MetroPanel1 = New MetroFramework.Controls.MetroPanel()
        Me.MetroPanel3 = New MetroFramework.Controls.MetroPanel()
        Me.MetroProgressBar1 = New MetroFramework.Controls.MetroProgressBar()
        Me.lblProgessInfo = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MetroPanel1.SuspendLayout()
        Me.MetroPanel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'MetroPanel1
        '
        Me.MetroPanel1.BackColor = System.Drawing.Color.Black
        Me.MetroPanel1.Controls.Add(Me.lblProgessInfo)
        Me.MetroPanel1.Controls.Add(Me.MetroPanel3)
        Me.MetroPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MetroPanel1.HorizontalScrollbarBarColor = True
        Me.MetroPanel1.HorizontalScrollbarHighlightOnWheel = False
        Me.MetroPanel1.HorizontalScrollbarSize = 10
        Me.MetroPanel1.Location = New System.Drawing.Point(0, 0)
        Me.MetroPanel1.Name = "MetroPanel1"
        Me.MetroPanel1.Size = New System.Drawing.Size(315, 76)
        Me.MetroPanel1.TabIndex = 0
        Me.MetroPanel1.UseCustomBackColor = True
        Me.MetroPanel1.VerticalScrollbarBarColor = True
        Me.MetroPanel1.VerticalScrollbarHighlightOnWheel = False
        Me.MetroPanel1.VerticalScrollbarSize = 10
        '
        'MetroPanel3
        '
        Me.MetroPanel3.Controls.Add(Me.MetroProgressBar1)
        Me.MetroPanel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.MetroPanel3.HorizontalScrollbarBarColor = True
        Me.MetroPanel3.HorizontalScrollbarHighlightOnWheel = False
        Me.MetroPanel3.HorizontalScrollbarSize = 10
        Me.MetroPanel3.Location = New System.Drawing.Point(0, 43)
        Me.MetroPanel3.Name = "MetroPanel3"
        Me.MetroPanel3.Size = New System.Drawing.Size(315, 33)
        Me.MetroPanel3.TabIndex = 3
        Me.MetroPanel3.VerticalScrollbarBarColor = True
        Me.MetroPanel3.VerticalScrollbarHighlightOnWheel = False
        Me.MetroPanel3.VerticalScrollbarSize = 10
        '
        'MetroProgressBar1
        '
        Me.MetroProgressBar1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MetroProgressBar1.Location = New System.Drawing.Point(0, 0)
        Me.MetroProgressBar1.Name = "MetroProgressBar1"
        Me.MetroProgressBar1.Size = New System.Drawing.Size(315, 33)
        Me.MetroProgressBar1.TabIndex = 2
        '
        'lblProgessInfo
        '
        Me.lblProgessInfo.AutoSize = True
        Me.lblProgessInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgessInfo.ForeColor = System.Drawing.Color.Lime
        Me.lblProgessInfo.Location = New System.Drawing.Point(109, 11)
        Me.lblProgessInfo.Name = "lblProgessInfo"
        Me.lblProgessInfo.Size = New System.Drawing.Size(24, 18)
        Me.lblProgessInfo.TabIndex = 4
        Me.lblProgessInfo.Text = " ..."
        '
        'Timer1
        '
        '
        'WelcomeScreen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(315, 76)
        Me.ControlBox = False
        Me.Controls.Add(Me.MetroPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "WelcomeScreen"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.MetroPanel1.ResumeLayout(False)
        Me.MetroPanel1.PerformLayout()
        Me.MetroPanel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents MetroPanel1 As MetroFramework.Controls.MetroPanel
    Friend WithEvents lblProgessInfo As Label
    Friend WithEvents MetroPanel3 As MetroFramework.Controls.MetroPanel
    Friend WithEvents MetroProgressBar1 As MetroFramework.Controls.MetroProgressBar
    Friend WithEvents Timer1 As Timer
End Class
