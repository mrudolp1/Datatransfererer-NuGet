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
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.PropertyGrid1 = New System.Windows.Forms.PropertyGrid()
        Me.pgTNX = New System.Windows.Forms.TabPage()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.btnSavetoEDS = New System.Windows.Forms.Button()
        Me.btnImportERI = New System.Windows.Forms.Button()
        Me.scFromERI = New System.Windows.Forms.SplitContainer()
        Me.propgridTNXERI = New System.Windows.Forms.PropertyGrid()
        Me.scFromEDS = New System.Windows.Forms.SplitContainer()
        Me.propgridTNXEDS = New System.Windows.Forms.PropertyGrid()
        Me.btnLoadfromEDS = New System.Windows.Forms.Button()
        Me.btnExportERI = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnCompare = New System.Windows.Forms.Button()
        Me.txtStrc = New System.Windows.Forms.TextBox()
        Me.lblStrc = New System.Windows.Forms.Label()
        Me.txtBU = New System.Windows.Forms.TextBox()
        Me.lblBU = New System.Windows.Forms.Label()
        Me.pgSQL = New System.Windows.Forms.TabPage()
        Me.sqltoexcel = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.exceltosql = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.btnTest = New System.Windows.Forms.Button()
        Me.TabPage2.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.pgTNX.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.scFromERI, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.scFromERI.Panel2.SuspendLayout()
        Me.scFromERI.SuspendLayout()
        CType(Me.scFromEDS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.scFromEDS.Panel1.SuspendLayout()
        Me.scFromEDS.Panel2.SuspendLayout()
        Me.scFromEDS.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.pgSQL.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.SplitContainer2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(883, 461)
        Me.TabPage2.TabIndex = 3
        Me.TabPage2.Text = "TNXBackup"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.Location = New System.Drawing.Point(3, 3)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.TextBox1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TextBox2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Button1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Button2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Button3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Button4)
        Me.SplitContainer2.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.No
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.PropertyGrid1)
        Me.SplitContainer2.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer2.Size = New System.Drawing.Size(877, 455)
        Me.SplitContainer2.SplitterDistance = 164
        Me.SplitContainer2.TabIndex = 4
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(51, 39)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 21)
        Me.TextBox1.TabIndex = 8
        Me.TextBox1.Text = "A"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(10, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(30, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Strc:"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(51, 12)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(100, 21)
        Me.TextBox2.TabIndex = 6
        Me.TextBox2.Text = "800000"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "BU:"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(3, 128)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(160, 21)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Load From EDS"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(3, 100)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(160, 21)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Save to EDS"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(3, 72)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(160, 21)
        Me.Button3.TabIndex = 1
        Me.Button3.Text = "Import TNX"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(3, 156)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(160, 21)
        Me.Button4.TabIndex = 2
        Me.Button4.Text = "Export TNX"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'PropertyGrid1
        '
        Me.PropertyGrid1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PropertyGrid1.Location = New System.Drawing.Point(0, 0)
        Me.PropertyGrid1.Name = "PropertyGrid1"
        Me.PropertyGrid1.Size = New System.Drawing.Size(709, 455)
        Me.PropertyGrid1.TabIndex = 3
        '
        'pgTNX
        '
        Me.pgTNX.Controls.Add(Me.SplitContainer1)
        Me.pgTNX.Controls.Add(Me.Panel1)
        Me.pgTNX.Location = New System.Drawing.Point(4, 22)
        Me.pgTNX.Name = "pgTNX"
        Me.pgTNX.Padding = New System.Windows.Forms.Padding(3)
        Me.pgTNX.Size = New System.Drawing.Size(883, 461)
        Me.pgTNX.TabIndex = 1
        Me.pgTNX.Text = "TNX"
        Me.pgTNX.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 48)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnSavetoEDS)
        Me.SplitContainer1.Panel1.Controls.Add(Me.btnImportERI)
        Me.SplitContainer1.Panel1.Controls.Add(Me.scFromERI)
        Me.SplitContainer1.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.No
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.scFromEDS)
        Me.SplitContainer1.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer1.Size = New System.Drawing.Size(877, 410)
        Me.SplitContainer1.SplitterDistance = 438
        Me.SplitContainer1.TabIndex = 6
        '
        'btnSavetoEDS
        '
        Me.btnSavetoEDS.Location = New System.Drawing.Point(3, 40)
        Me.btnSavetoEDS.Name = "btnSavetoEDS"
        Me.btnSavetoEDS.Size = New System.Drawing.Size(160, 21)
        Me.btnSavetoEDS.TabIndex = 3
        Me.btnSavetoEDS.Text = "Save to EDS"
        Me.btnSavetoEDS.UseVisualStyleBackColor = True
        '
        'btnImportERI
        '
        Me.btnImportERI.Location = New System.Drawing.Point(3, 10)
        Me.btnImportERI.Name = "btnImportERI"
        Me.btnImportERI.Size = New System.Drawing.Size(160, 21)
        Me.btnImportERI.TabIndex = 1
        Me.btnImportERI.Text = "Import ERI"
        Me.btnImportERI.UseVisualStyleBackColor = True
        '
        'scFromERI
        '
        Me.scFromERI.Dock = System.Windows.Forms.DockStyle.Fill
        Me.scFromERI.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.scFromERI.Location = New System.Drawing.Point(0, 0)
        Me.scFromERI.Name = "scFromERI"
        '
        'scFromERI.Panel2
        '
        Me.scFromERI.Panel2.Controls.Add(Me.propgridTNXERI)
        Me.scFromERI.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.scFromERI.Size = New System.Drawing.Size(438, 410)
        Me.scFromERI.SplitterDistance = 164
        Me.scFromERI.TabIndex = 4
        '
        'propgridTNXERI
        '
        Me.propgridTNXERI.Dock = System.Windows.Forms.DockStyle.Fill
        Me.propgridTNXERI.Location = New System.Drawing.Point(0, 0)
        Me.propgridTNXERI.Name = "propgridTNXERI"
        Me.propgridTNXERI.Size = New System.Drawing.Size(270, 410)
        Me.propgridTNXERI.TabIndex = 4
        '
        'scFromEDS
        '
        Me.scFromEDS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.scFromEDS.Location = New System.Drawing.Point(0, 0)
        Me.scFromEDS.Name = "scFromEDS"
        '
        'scFromEDS.Panel1
        '
        Me.scFromEDS.Panel1.Controls.Add(Me.propgridTNXEDS)
        '
        'scFromEDS.Panel2
        '
        Me.scFromEDS.Panel2.Controls.Add(Me.btnLoadfromEDS)
        Me.scFromEDS.Panel2.Controls.Add(Me.btnExportERI)
        Me.scFromEDS.Size = New System.Drawing.Size(435, 410)
        Me.scFromEDS.SplitterDistance = 275
        Me.scFromEDS.TabIndex = 0
        '
        'propgridTNXEDS
        '
        Me.propgridTNXEDS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.propgridTNXEDS.Location = New System.Drawing.Point(0, 0)
        Me.propgridTNXEDS.Name = "propgridTNXEDS"
        Me.propgridTNXEDS.Size = New System.Drawing.Size(275, 410)
        Me.propgridTNXEDS.TabIndex = 4
        '
        'btnLoadfromEDS
        '
        Me.btnLoadfromEDS.Location = New System.Drawing.Point(-1, 10)
        Me.btnLoadfromEDS.Name = "btnLoadfromEDS"
        Me.btnLoadfromEDS.Size = New System.Drawing.Size(160, 21)
        Me.btnLoadfromEDS.TabIndex = 6
        Me.btnLoadfromEDS.Text = "Load From EDS"
        Me.btnLoadfromEDS.UseVisualStyleBackColor = True
        '
        'btnExportERI
        '
        Me.btnExportERI.Location = New System.Drawing.Point(-1, 38)
        Me.btnExportERI.Name = "btnExportERI"
        Me.btnExportERI.Size = New System.Drawing.Size(160, 21)
        Me.btnExportERI.TabIndex = 5
        Me.btnExportERI.Text = "Export ERI"
        Me.btnExportERI.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnTest)
        Me.Panel1.Controls.Add(Me.btnCompare)
        Me.Panel1.Controls.Add(Me.txtStrc)
        Me.Panel1.Controls.Add(Me.lblStrc)
        Me.Panel1.Controls.Add(Me.txtBU)
        Me.Panel1.Controls.Add(Me.lblBU)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(877, 45)
        Me.Panel1.TabIndex = 5
        '
        'btnCompare
        '
        Me.btnCompare.Location = New System.Drawing.Point(358, 12)
        Me.btnCompare.Name = "btnCompare"
        Me.btnCompare.Size = New System.Drawing.Size(160, 21)
        Me.btnCompare.TabIndex = 9
        Me.btnCompare.Text = "Compare"
        Me.btnCompare.UseVisualStyleBackColor = True
        '
        'txtStrc
        '
        Me.txtStrc.Location = New System.Drawing.Point(205, 12)
        Me.txtStrc.Name = "txtStrc"
        Me.txtStrc.Size = New System.Drawing.Size(100, 21)
        Me.txtStrc.TabIndex = 8
        Me.txtStrc.Text = "A"
        '
        'lblStrc
        '
        Me.lblStrc.AutoSize = True
        Me.lblStrc.Location = New System.Drawing.Point(164, 16)
        Me.lblStrc.Name = "lblStrc"
        Me.lblStrc.Size = New System.Drawing.Size(30, 13)
        Me.lblStrc.TabIndex = 7
        Me.lblStrc.Text = "Strc:"
        '
        'txtBU
        '
        Me.txtBU.Location = New System.Drawing.Point(49, 12)
        Me.txtBU.Name = "txtBU"
        Me.txtBU.Size = New System.Drawing.Size(100, 21)
        Me.txtBU.TabIndex = 6
        Me.txtBU.Text = "800000"
        '
        'lblBU
        '
        Me.lblBU.AutoSize = True
        Me.lblBU.Location = New System.Drawing.Point(8, 16)
        Me.lblBU.Name = "lblBU"
        Me.lblBU.Size = New System.Drawing.Size(24, 13)
        Me.lblBU.TabIndex = 5
        Me.lblBU.Text = "BU:"
        '
        'pgSQL
        '
        Me.pgSQL.Controls.Add(Me.sqltoexcel)
        Me.pgSQL.Controls.Add(Me.PictureBox1)
        Me.pgSQL.Controls.Add(Me.exceltosql)
        Me.pgSQL.Location = New System.Drawing.Point(4, 22)
        Me.pgSQL.Name = "pgSQL"
        Me.pgSQL.Padding = New System.Windows.Forms.Padding(3)
        Me.pgSQL.Size = New System.Drawing.Size(883, 461)
        Me.pgSQL.TabIndex = 0
        Me.pgSQL.Text = "SQL"
        Me.pgSQL.UseVisualStyleBackColor = True
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
        'exceltosql
        '
        Me.exceltosql.Location = New System.Drawing.Point(21, 111)
        Me.exceltosql.Name = "exceltosql"
        Me.exceltosql.Size = New System.Drawing.Size(160, 52)
        Me.exceltosql.TabIndex = 1
        Me.exceltosql.Text = "Load from Excel / Save to SQL"
        Me.exceltosql.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.pgSQL)
        Me.TabControl1.Controls.Add(Me.pgTNX)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(891, 487)
        Me.TabControl1.TabIndex = 3
        '
        'btnTest
        '
        Me.btnTest.Location = New System.Drawing.Point(557, 11)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.Size = New System.Drawing.Size(160, 21)
        Me.btnTest.TabIndex = 5
        Me.btnTest.Text = "Test"
        Me.btnTest.UseVisualStyleBackColor = True
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
        Me.TabPage2.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.pgTNX.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.scFromERI.Panel2.ResumeLayout(False)
        CType(Me.scFromERI, System.ComponentModel.ISupportInitialize).EndInit()
        Me.scFromERI.ResumeLayout(False)
        Me.scFromEDS.Panel1.ResumeLayout(False)
        Me.scFromEDS.Panel2.ResumeLayout(False)
        CType(Me.scFromEDS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.scFromEDS.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.pgSQL.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents SplitContainer2 As SplitContainer
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents PropertyGrid1 As PropertyGrid
    Friend WithEvents pgTNX As TabPage
    Friend WithEvents scFromERI As SplitContainer
    Friend WithEvents txtStrc As TextBox
    Friend WithEvents lblStrc As Label
    Friend WithEvents txtBU As TextBox
    Friend WithEvents lblBU As Label
    Friend WithEvents btnSavetoEDS As Button
    Friend WithEvents btnImportERI As Button
    Friend WithEvents pgSQL As TabPage
    Friend WithEvents sqltoexcel As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents exceltosql As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents scFromEDS As SplitContainer
    Friend WithEvents propgridTNXEDS As PropertyGrid
    Friend WithEvents btnLoadfromEDS As Button
    Friend WithEvents btnExportERI As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents btnCompare As Button
    Friend WithEvents propgridTNXERI As PropertyGrid
    Friend WithEvents btnTest As Button

#End Region

End Class
