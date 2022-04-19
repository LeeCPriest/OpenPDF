<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ConfigSelect
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
        Me.ConfigListBox = New System.Windows.Forms.ListBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.InfoText = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ConfigListBox
        '
        Me.ConfigListBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ConfigListBox.FormattingEnabled = True
        Me.ConfigListBox.Location = New System.Drawing.Point(12, 25)
        Me.ConfigListBox.Name = "ConfigListBox"
        Me.ConfigListBox.Size = New System.Drawing.Size(360, 134)
        Me.ConfigListBox.TabIndex = 0
        '
        'OKButton
        '
        Me.OKButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.OKButton.Location = New System.Drawing.Point(160, 169)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(75, 23)
        Me.OKButton.TabIndex = 1
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'InfoText
        '
        Me.InfoText.AutoSize = True
        Me.InfoText.Location = New System.Drawing.Point(12, 8)
        Me.InfoText.Name = "InfoText"
        Me.InfoText.Size = New System.Drawing.Size(39, 13)
        Me.InfoText.TabIndex = 2
        Me.InfoText.Text = "Label1"
        '
        'ConfigSelect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(384, 201)
        Me.Controls.Add(Me.InfoText)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.ConfigListBox)
        Me.MinimumSize = New System.Drawing.Size(100, 50)
        Me.Name = "ConfigSelect"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select Configuration"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ConfigListBox As Windows.Forms.ListBox
    Friend WithEvents OKButton As Windows.Forms.Button
    Friend WithEvents InfoText As Windows.Forms.Label
End Class
