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
        Me.sqltoexcel = New System.Windows.Forms.Button()
        Me.exceltosql = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.pgSQL = New System.Windows.Forms.TabPage()
        Me.pgTNX = New System.Windows.Forms.TabPage()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.btnImportTNX = New System.Windows.Forms.Button()
        Me.btnExportTNX = New System.Windows.Forms.Button()
        Me.propgridTNXObject = New System.Windows.Forms.PropertyGrid()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.pgSQL.SuspendLayout()
        Me.pgTNX.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'sqltoexcel
        '
        Me.sqltoexcel.Location = New System.Drawing.Point(21, 21)
        Me.sqltoexcel.Name = "sqltoexcel"
        Me.sqltoexcel.Size = New System.Drawing.Size(160, 52)
        Me.sqltoexcel.TabIndex = 0
        Me.sqltoexcel.Text = "Load from SQL / Save to Excel"
        Me.sqltoexcel.UseVisualStyleBackColor = True
        '
        'exceltosql
        '
        Me.exceltosql.Location = New System.Drawing.Point(21, 111)
        Me.exceltosql.Name = "exceltosql"
        Me.exceltosql.Size = New System.Drawing.Size(160, 52)
        Me.exceltosql.TabIndex = 1
        Me.exceltosql.Text = "Load from Excel / Save to SQL"
        Me.exceltosql.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(245, 23)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(214, 140)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.pgSQL)
        Me.TabControl1.Controls.Add(Me.pgTNX)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(507, 233)
        Me.TabControl1.TabIndex = 3
        '
        'pgSQL
        '
        Me.pgSQL.Controls.Add(Me.sqltoexcel)
        Me.pgSQL.Controls.Add(Me.PictureBox1)
        Me.pgSQL.Controls.Add(Me.exceltosql)
        Me.pgSQL.Location = New System.Drawing.Point(4, 22)
        Me.pgSQL.Name = "pgSQL"
        Me.pgSQL.Padding = New System.Windows.Forms.Padding(3)
        Me.pgSQL.Size = New System.Drawing.Size(499, 207)
        Me.pgSQL.TabIndex = 0
        Me.pgSQL.Text = "SQL"
        Me.pgSQL.UseVisualStyleBackColor = True
        '
        'pgTNX
        '
        Me.pgTNX.Controls.Add(Me.SplitContainer1)
        Me.pgTNX.Location = New System.Drawing.Point(4, 22)
        Me.pgTNX.Name = "pgTNX"
        Me.pgTNX.Padding = New System.Windows.Forms.Padding(3)
        Me.pgTNX.Size = New System.Drawing.Size(499, 207)
        Me.pgTNX.TabIndex = 1
        Me.pgTNX.Text = "TNX"
        Me.pgTNX.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 3)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnImportTNX)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnExportTNX)
        Me.SplitContainer1.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.No
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.propgridTNXObject)
        Me.SplitContainer1.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer1.Size = New System.Drawing.Size(493, 201)
        Me.SplitContainer1.SplitterDistance = 164
        Me.SplitContainer1.TabIndex = 4
        '
        'btnImportTNX
        '
        Me.btnImportTNX.Location = New System.Drawing.Point(3, 3)
        Me.btnImportTNX.Name = "btnImportTNX"
        Me.btnImportTNX.Size = New System.Drawing.Size(160, 52)
        Me.btnImportTNX.TabIndex = 1
        Me.btnImportTNX.Text = "Import TNX"
        Me.btnImportTNX.UseVisualStyleBackColor = True
        '
        'btnExportTNX
        '
        Me.btnExportTNX.Location = New System.Drawing.Point(3, 81)
        Me.btnExportTNX.Name = "btnExportTNX"
        Me.btnExportTNX.Size = New System.Drawing.Size(160, 52)
        Me.btnExportTNX.TabIndex = 2
        Me.btnExportTNX.Text = "Export TNX"
        Me.btnExportTNX.UseVisualStyleBackColor = True
        '
        'propgridTNXObject
        '
        Me.propgridTNXObject.Dock = System.Windows.Forms.DockStyle.Fill
        Me.propgridTNXObject.Location = New System.Drawing.Point(0, 0)
        Me.propgridTNXObject.Name = "propgridTNXObject"
        Me.propgridTNXObject.Size = New System.Drawing.Size(325, 201)
        Me.propgridTNXObject.TabIndex = 3
        '
        'frmMain
        '
        Me.Appearance.BackColor = System.Drawing.Color.White
        Me.Appearance.Options.UseBackColor = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(507, 233)
        Me.Controls.Add(Me.TabControl1)
        Me.IconOptions.Image = CType(resources.GetObject("frmMain.IconOptions.Image"), System.Drawing.Image)
        Me.Name = "frmMain"
        Me.Text = "EDS & Excel Testing"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.pgSQL.ResumeLayout(False)
        Me.pgTNX.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents sqltoexcel As Button
    Friend WithEvents exceltosql As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents pgSQL As TabPage
    Friend WithEvents pgTNX As TabPage
    Friend WithEvents btnExportTNX As Button
    Friend WithEvents btnImportTNX As Button
    Friend WithEvents propgridTNXObject As PropertyGrid
    Friend WithEvents SplitContainer1 As SplitContainer

#End Region

End Class
