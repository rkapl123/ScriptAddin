<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ScriptOutput
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ScriptOutput))
        Me.ScriptOutputTextbox = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'ScriptOutputTextbox
        '
        Me.ScriptOutputTextbox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ScriptOutputTextbox.BackColor = System.Drawing.SystemColors.WindowText
        Me.ScriptOutputTextbox.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ScriptOutputTextbox.ForeColor = System.Drawing.SystemColors.Window
        Me.ScriptOutputTextbox.Location = New System.Drawing.Point(2, 1)
        Me.ScriptOutputTextbox.Name = "ScriptOutputTextbox"
        Me.ScriptOutputTextbox.Size = New System.Drawing.Size(796, 447)
        Me.ScriptOutputTextbox.TabIndex = 0
        Me.ScriptOutputTextbox.Text = ""
        '
        'ScriptOutput
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.ScriptOutputTextbox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "ScriptOutput"
        Me.Text = "Script Output"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ScriptOutputTextbox As Windows.Forms.RichTextBox
End Class
