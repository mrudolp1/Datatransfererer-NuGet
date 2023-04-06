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
        Me.btnImportStrcFiles = New System.Windows.Forms.Button()
        Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
        Me.btnSaveFndToEDS = New System.Windows.Forms.Button()
        Me.propgridFndXL = New System.Windows.Forms.PropertyGrid()
        Me.SplitContainer5 = New System.Windows.Forms.SplitContainer()
        Me.propgridFndEDS = New System.Windows.Forms.PropertyGrid()
        Me.btnLoadFndFromEDS = New System.Windows.Forms.Button()
        Me.btnExportStrcFiles = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtDirectory = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtFndWO = New System.Windows.Forms.TextBox()
        Me.btnCompareStrc = New System.Windows.Forms.Button()
        Me.txtFndStrc = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtFndBU = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
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
        Me.btnTest = New System.Windows.Forms.Button()
        Me.btnCompare = New System.Windows.Forms.Button()
        Me.txtStrc = New System.Windows.Forms.TextBox()
        Me.lblStrc = New System.Windows.Forms.Label()
        Me.txtBU = New System.Windows.Forms.TextBox()
        Me.lblBU = New System.Windows.Forms.Label()
        Me.pgSQLBackUp = New System.Windows.Forms.TabPage()
        Me.txtSQLStrc = New System.Windows.Forms.TextBox()
        Me.txtSQLBU = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.sqltoexcel = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.exceltosql = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.pgStructure.SuspendLayout()
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer4.Panel1.SuspendLayout()
        Me.SplitContainer4.Panel2.SuspendLayout()
        Me.SplitContainer4.SuspendLayout()
        CType(Me.SplitContainer5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer5.Panel1.SuspendLayout()
        Me.SplitContainer5.Panel2.SuspendLayout()
        Me.SplitContainer5.SuspendLayout()
        Me.Panel2.SuspendLayout()
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
        Me.pgSQLBackUp.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.SplitContainer3.Location = New System.Drawing.Point(3, 80)
        Me.SplitContainer3.Name = "SplitContainer3"
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.btnImportStrcFiles)
        Me.SplitContainer3.Panel1.Controls.Add(Me.SplitContainer4)
        Me.SplitContainer3.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.No
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.SplitContainer5)
        Me.SplitContainer3.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer3.Size = New System.Drawing.Size(877, 378)
        Me.SplitContainer3.SplitterDistance = 438
        Me.SplitContainer3.TabIndex = 7
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
        'SplitContainer4.Panel1
        '
        Me.SplitContainer4.Panel1.Controls.Add(Me.btnSaveFndToEDS)
        '
        'SplitContainer4.Panel2
        '
        Me.SplitContainer4.Panel2.Controls.Add(Me.propgridFndXL)
        Me.SplitContainer4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.SplitContainer4.Size = New System.Drawing.Size(438, 378)
        Me.SplitContainer4.SplitterDistance = 164
        Me.SplitContainer4.TabIndex = 4
        '
        'btnSaveFndToEDS
        '
        Me.btnSaveFndToEDS.Location = New System.Drawing.Point(3, 38)
        Me.btnSaveFndToEDS.Name = "btnSaveFndToEDS"
        Me.btnSaveFndToEDS.Size = New System.Drawing.Size(160, 21)
        Me.btnSaveFndToEDS.TabIndex = 3
        Me.btnSaveFndToEDS.Text = "Save to EDS"
        Me.btnSaveFndToEDS.UseVisualStyleBackColor = True
        '
        'propgridFndXL
        '
        Me.propgridFndXL.Dock = System.Windows.Forms.DockStyle.Fill
        Me.propgridFndXL.Location = New System.Drawing.Point(0, 0)
        Me.propgridFndXL.Name = "propgridFndXL"
        Me.propgridFndXL.Size = New System.Drawing.Size(270, 378)
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
        Me.SplitContainer5.Size = New System.Drawing.Size(435, 378)
        Me.SplitContainer5.SplitterDistance = 275
        Me.SplitContainer5.TabIndex = 0
        '
        'propgridFndEDS
        '
        Me.propgridFndEDS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.propgridFndEDS.Location = New System.Drawing.Point(0, 0)
        Me.propgridFndEDS.Name = "propgridFndEDS"
        Me.propgridFndEDS.Size = New System.Drawing.Size(275, 378)
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
        Me.Panel2.Controls.Add(Me.btnBrowse)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.txtDirectory)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.txtFndWO)
        Me.Panel2.Controls.Add(Me.btnCompareStrc)
        Me.Panel2.Controls.Add(Me.txtFndStrc)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.txtFndBU)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(3, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(877, 77)
        Me.Panel2.TabIndex = 6
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(723, 40)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(79, 21)
        Me.btnBrowse.TabIndex = 14
        Me.btnBrowse.Text = "Browse..."
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(11, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(97, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Working Directory:"
        '
        'txtDirectory
        '
        Me.txtDirectory.Location = New System.Drawing.Point(114, 41)
        Me.txtDirectory.Name = "txtDirectory"
        Me.txtDirectory.Size = New System.Drawing.Size(603, 21)
        Me.txtDirectory.TabIndex = 12
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(582, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(29, 13)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "WO:"
        '
        'txtFndWO
        '
        Me.txtFndWO.Location = New System.Drawing.Point(617, 12)
        Me.txtFndWO.Name = "txtFndWO"
        Me.txtFndWO.Size = New System.Drawing.Size(100, 21)
        Me.txtFndWO.TabIndex = 10
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
        'btnTest
        '
        Me.btnTest.Location = New System.Drawing.Point(557, 11)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.Size = New System.Drawing.Size(160, 21)
        Me.btnTest.TabIndex = 5
        Me.btnTest.Text = "Test"
        Me.btnTest.UseVisualStyleBackColor = True
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
        'pgSQLBackUp
        '
        Me.pgSQLBackUp.Controls.Add(Me.txtSQLStrc)
        Me.pgSQLBackUp.Controls.Add(Me.txtSQLBU)
        Me.pgSQLBackUp.Controls.Add(Me.Label3)
        Me.pgSQLBackUp.Controls.Add(Me.Label4)
        Me.pgSQLBackUp.Controls.Add(Me.sqltoexcel)
        Me.pgSQLBackUp.Controls.Add(Me.PictureBox1)
        Me.pgSQLBackUp.Controls.Add(Me.exceltosql)
        Me.pgSQLBackUp.Location = New System.Drawing.Point(4, 22)
        Me.pgSQLBackUp.Name = "pgSQLBackUp"
        Me.pgSQLBackUp.Padding = New System.Windows.Forms.Padding(3)
        Me.pgSQLBackUp.Size = New System.Drawing.Size(883, 461)
        Me.pgSQLBackUp.TabIndex = 0
        Me.pgSQLBackUp.Text = "SQL"
        Me.pgSQLBackUp.UseVisualStyleBackColor = True
        '
        'txtSQLStrc
        '
        Me.txtSQLStrc.Location = New System.Drawing.Point(215, 14)
        Me.txtSQLStrc.Name = "txtSQLStrc"
        Me.txtSQLStrc.Size = New System.Drawing.Size(100, 21)
        Me.txtSQLStrc.TabIndex = 12
        Me.txtSQLStrc.Text = "A"
        '
        'txtSQLBU
        '
        Me.txtSQLBU.Location = New System.Drawing.Point(59, 14)
        Me.txtSQLBU.Name = "txtSQLBU"
        Me.txtSQLBU.Size = New System.Drawing.Size(100, 21)
        Me.txtSQLBU.TabIndex = 10
        Me.txtSQLBU.Text = "800000"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(174, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(30, 13)
        Me.Label3.TabIndex = 11
        Me.Label3.Text = "Strc:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(18, 18)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 13)
        Me.Label4.TabIndex = 9
        Me.Label4.Text = "BU:"
        '
        'sqltoexcel
        '
        Me.sqltoexcel.Location = New System.Drawing.Point(21, 43)
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
        Me.PictureBox1.Location = New System.Drawing.Point(245, 45)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(214, 140)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'exceltosql
        '
        Me.exceltosql.Location = New System.Drawing.Point(21, 133)
        Me.exceltosql.Name = "exceltosql"
        Me.exceltosql.Size = New System.Drawing.Size(160, 52)
        Me.exceltosql.TabIndex = 1
        Me.exceltosql.Text = "Load from Excel / Save to SQL"
        Me.exceltosql.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.pgStructure)
        Me.TabControl1.Controls.Add(Me.pgSQLBackUp)
        Me.TabControl1.Controls.Add(Me.pgTNX)
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
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "EDS & Excel Testing"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pgStructure.ResumeLayout(False)
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.ResumeLayout(False)
        Me.SplitContainer4.Panel1.ResumeLayout(False)
        Me.SplitContainer4.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer4.ResumeLayout(False)
        Me.SplitContainer5.Panel1.ResumeLayout(False)
        Me.SplitContainer5.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer5.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
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
        Me.pgSQLBackUp.ResumeLayout(False)
        Me.pgSQLBackUp.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents Panel2 As Panel
    Friend WithEvents btnCompareStrc As Button
    Friend WithEvents txtFndStrc As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents txtFndBU As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents pgTNX As TabPage
    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents btnSavetoEDS As Button
    Friend WithEvents btnImportERI As Button
    Friend WithEvents scFromERI As SplitContainer
    Friend WithEvents propgridTNXERI As PropertyGrid
    Friend WithEvents scFromEDS As SplitContainer
    Friend WithEvents propgridTNXEDS As PropertyGrid
    Friend WithEvents btnLoadfromEDS As Button
    Friend WithEvents btnExportERI As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents btnTest As Button
    Friend WithEvents btnCompare As Button
    Friend WithEvents txtStrc As TextBox
    Friend WithEvents lblStrc As Label
    Friend WithEvents txtBU As TextBox
    Friend WithEvents lblBU As Label
    Friend WithEvents pgSQLBackUp As TabPage
    Friend WithEvents txtSQLStrc As TextBox
    Friend WithEvents txtSQLBU As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents sqltoexcel As Button
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents exceltosql As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents Label7 As Label
    Friend WithEvents txtFndWO As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtDirectory As TextBox
    Friend WithEvents btnBrowse As Button

#End Region

End Class
