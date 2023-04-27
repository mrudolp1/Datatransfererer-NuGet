Namespace UnitTesting
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class frmSummary
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
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSummary))
            Me.tcData = New System.Windows.Forms.TabControl()
            Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
            Me.ToolStripSplitButton1 = New System.Windows.Forms.ToolStripDropDownButton()
            Me.StatusStrip1.SuspendLayout()
            Me.SuspendLayout()
            '
            'tcData
            '
            Me.tcData.Dock = System.Windows.Forms.DockStyle.Fill
            Me.tcData.Location = New System.Drawing.Point(0, 0)
            Me.tcData.Name = "tcData"
            Me.tcData.SelectedIndex = 0
            Me.tcData.Size = New System.Drawing.Size(1162, 634)
            Me.tcData.TabIndex = 0
            '
            'StatusStrip1
            '
            Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSplitButton1})
            Me.StatusStrip1.Location = New System.Drawing.Point(0, 612)
            Me.StatusStrip1.Name = "StatusStrip1"
            Me.StatusStrip1.Size = New System.Drawing.Size(1162, 22)
            Me.StatusStrip1.TabIndex = 1
            Me.StatusStrip1.Text = "StatusStrip1"
            '
            'ToolStripSplitButton1
            '
            Me.ToolStripSplitButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
            Me.ToolStripSplitButton1.Image = CType(resources.GetObject("ToolStripSplitButton1.Image"), System.Drawing.Image)
            Me.ToolStripSplitButton1.ImageTransparentColor = System.Drawing.Color.Magenta
            Me.ToolStripSplitButton1.Name = "ToolStripSplitButton1"
            Me.ToolStripSplitButton1.ShowDropDownArrow = False
            Me.ToolStripSplitButton1.Size = New System.Drawing.Size(45, 20)
            Me.ToolStripSplitButton1.Text = "Export"
            '
            'frmSummary
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(1162, 634)
            Me.Controls.Add(Me.StatusStrip1)
            Me.Controls.Add(Me.tcData)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            Me.Name = "frmSummary"
            Me.Text = "Summary"
            Me.StatusStrip1.ResumeLayout(False)
            Me.StatusStrip1.PerformLayout()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

        Friend WithEvents tcData As TabControl
        Friend WithEvents StatusStrip1 As StatusStrip
        Friend WithEvents ToolStripSplitButton1 As ToolStripDropDownButton
    End Class

End Namespace
