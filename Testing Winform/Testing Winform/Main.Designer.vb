Partial Public Class frmMain
    Inherits DevExpress.XtraEditors.XtraForm

    ''' <summary>
    ''' Required designer variable.
    ''' </summary>
    Private components As System.ComponentModel.IContainer = Nothing

    ''' <summary>
    ''' Clean up any resources being used.
    ''' </summary>
    ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso (components IsNot Nothing) Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

#Region "Windows Form Designer generated code"

    ''' <summary>
    ''' Required method for Designer support - do not modify
    ''' the contents of this method with the code editor.
    ''' </summary>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.pgStructure = New System.Windows.Forms.TabPage()
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.btnSaveFndToEDS = New System.Windows.Forms.Button()
        Me.btnImportStrcFiles = New System.Windows.Forms.Button()
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.propgridFndXL = New System.Windows.Forms.PropertyGrid()
        Me.SplitContainer5 = New System.Windows.Forms.SplitContainer()
        Me.propgridFndEDS = New System.Windows.Forms.PropertyGrid()
        Me.btnLoadFndFromEDS = New System.Windows.Forms.Button()
        Me.btnExportStrcFiles = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnCompareStrc = New System.Windows.Forms.Button()
        Me.txtFndStrc = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtFndBU = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.pgStructure.SuspendLayout()
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        CType(Me.SplitContainer5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer5.Panel1.SuspendLayout()
        Me.SplitContainer5.Panel2.SuspendLayout()
        Me.SplitContainer5.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'pgStructure
        '
        Me.pgStructure.Controls.Add(Me.SplitContainer3)
        Me.pgStructure.Controls.Add(Me.Panel2)
        Me.pgStructure.Location = New System.Drawing.Point(4, 22)
        Me.pgStructure.Name = "pgStructure"
        Me.pgStructure.Padding = New System.Windows.Forms.Padding(3)
        Me.pgStructure.Size = New System.Drawing.Size(883, 461)
        Me.pgStructure.TabIndex = 4
        Me.pgStructure.Text = "Structure"
        Me.pgStructure.UseVisualStyleBackColor = True
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(3, 48)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.btnSaveFndToEDS)
        Me.SplitContainer3.Panel1.Controls.Add(Me.btnImportStrcFiles)
        Me.SplitContainer3.Panel1.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer3.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.No
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer5)
        Me.SplitContainer3.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer3.Size = New System.Drawing.Size(877, 410)
        Me.SplitContainer3.SplitterDistance = 438
        Me.SplitContainer3.TabIndex = 7
        '
        'btnSaveFndToEDS
        '
        Me.btnSaveFndToEDS.Location = New System.Drawing.Point(3, 40)
        Me.btnSaveFndToEDS.Name = "btnSaveFndToEDS"
        Me.btnSaveFndToEDS.Size = New System.Drawing.Size(160, 21)
        Me.btnSaveFndToEDS.TabIndex = 3
        Me.btnSaveFndToEDS.Text = "Save to EDS"
        Me.btnSaveFndToEDS.UseVisualStyleBackColor = True
        '
        'btnImportStrcFiles
        '
        Me.btnImportStrcFiles.Location = New System.Drawing.Point(3, 10)
        Me.btnImportStrcFiles.Name = "btnImportStrcFiles"
        Me.btnImportStrcFiles.Size = New System.Drawing.Size(160, 21)
        Me.btnImportStrcFiles.TabIndex = 1
        Me.btnImportStrcFiles.Text = "Import Structure Files"
        Me.btnImportStrcFiles.UseVisualStyleBackColor = True
        '
        'SplitContainer4
        '
        Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer4.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer4.Name = "SplitContainer4"
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.propgridFndXL)
        Me.SplitContainer4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer4.Size = New System.Drawing.Size(438, 410)
        Me.SplitContainer4.SplitterDistance = 164
        Me.SplitContainer4.TabIndex = 4
        '
        'propgridFndXL
        '
        Me.propgridFndXL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.propgridFndXL.Location = New System.Drawing.Point(0, 0)
        Me.propgridFndXL.Name = "propgridFndXL"
        Me.propgridFndXL.Size = New System.Drawing.Size(270, 410)
        Me.propgridFndXL.TabIndex = 4
        '
        'SplitContainer5
        '
        Me.SplitContainer5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer5.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer5.Name = "SplitContainer5"
        '
        'SplitContainer5.Panel1
        '
        Me.SplitContainer5.Panel1.Controls.Add(Me.propgridFndEDS)
        '
        'SplitContainer5.Panel2
        '
        Me.SplitContainer5.Panel2.Controls.Add(Me.btnLoadFndFromEDS)
        Me.SplitContainer5.Panel2.Controls.Add(Me.btnExportStrcFiles)
        Me.SplitContainer5.Size = New System.Drawing.Size(435, 410)
        Me.SplitContainer5.SplitterDistance = 275
        Me.SplitContainer5.TabIndex = 0
        '
        'propgridFndEDS
        '
        Me.propgridFndEDS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.propgridFndEDS.Location = New System.Drawing.Point(0, 0)
        Me.propgridFndEDS.Name = "propgridFndEDS"
        Me.propgridFndEDS.Size = New System.Drawing.Size(275, 410)
        Me.propgridFndEDS.TabIndex = 4
        '
        'btnLoadFndFromEDS
        '
        Me.btnLoadFndFromEDS.Location = New System.Drawing.Point(-1, 10)
        Me.btnLoadFndFromEDS.Name = "btnLoadFndFromEDS"
        Me.btnLoadFndFromEDS.Size = New System.Drawing.Size(160, 21)
        Me.btnLoadFndFromEDS.TabIndex = 6
        Me.btnLoadFndFromEDS.Text = "Load From EDS"
        Me.btnLoadFndFromEDS.UseVisualStyleBackColor = True
        '
        'btnExportStrcFiles
        '
        Me.btnExportStrcFiles.Location = New System.Drawing.Point(-1, 38)
        Me.btnExportStrcFiles.Name = "btnExportStrcFiles"
        Me.btnExportStrcFiles.Size = New System.Drawing.Size(160, 21)
        Me.btnExportStrcFiles.TabIndex = 5
        Me.btnExportStrcFiles.Text = "Export Structure Files"
        Me.btnExportStrcFiles.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnCompareStrc)
        Me.Panel2.Controls.Add(Me.txtFndStrc)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.txtFndBU)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(3, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(877, 45)
        Me.Panel2.TabIndex = 6
        '
        'btnCompareStrc
        '
        Me.btnCompareStrc.Location = New System.Drawing.Point(358, 12)
        Me.btnCompareStrc.Name = "btnCompareStrc"
        Me.btnCompareStrc.Size = New System.Drawing.Size(160, 21)
        Me.btnCompareStrc.TabIndex = 9
        Me.btnCompareStrc.Text = "Compare"
        Me.btnCompareStrc.UseVisualStyleBackColor = True
        '
        'txtFndStrc
        '
        Me.txtFndStrc.Location = New System.Drawing.Point(205, 12)
        Me.txtFndStrc.Name = "txtFndStrc"
        Me.txtFndStrc.Size = New System.Drawing.Size(100, 21)
        Me.txtFndStrc.TabIndex = 8
        Me.txtFndStrc.Text = "A"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(164, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Strc:"
        '
        'txtFndBU
        '
        Me.txtFndBU.Location = New System.Drawing.Point(49, 12)
        Me.txtFndBU.Name = "txtFndBU"
        Me.txtFndBU.Size = New System.Drawing.Size(100, 21)
        Me.txtFndBU.TabIndex = 6
        Me.txtFndBU.Text = "800000"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(8, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(24, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "BU:"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.pgStructure)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(891, 487)
        Me.TabControl1.TabIndex = 3
        '
        'frmMain
        '
        Me.Appearance.BackColor = System.Drawing.Color.White
        Me.Appearance.Options.UseBackColor = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(891, 487)
        Me.Controls.Add(Me.TabControl1)
        Me.IconOptions.Image = CType(resources.GetObject("frmMain.IconOptions.Image"), System.Drawing.Image)
        Me.Name = "frmMain"
        Me.Text = "EDS & Excel Testing"
        Me.pgStructure.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.ResumeLayout(False)
        Me.SplitContainer5.Panel1.ResumeLayout(False)
        Me.SplitContainer5.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer5.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pgStructure As TabPage
    Friend WithEvents SplitContainer3 As SplitContainer
    Friend WithEvents btnSaveFndToEDS As Button
    Friend WithEvents btnImportStrcFiles As Button
    Friend WithEvents SplitContainer4 As SplitContainer
    Friend WithEvents propgridFndXL As PropertyGrid
    Friend WithEvents SplitContainer5 As SplitContainer
    Friend WithEvents propgridFndEDS As PropertyGrid
    Friend WithEvents btnLoadFndFromEDS As Button
    Friend WithEvents btnExportStrcFiles As Button
    Friend WithEvents btnCompareStrc As Button
    Friend WithEvents txtFndStrc As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtFndBU As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents Panel2 As Panel

#End Region

End Class
