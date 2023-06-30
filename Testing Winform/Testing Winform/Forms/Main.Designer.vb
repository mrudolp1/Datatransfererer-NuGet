Namespace UnitTesting
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
            Me.btnLoopThroughERI = New System.Windows.Forms.Button()
            Me.Button1 = New System.Windows.Forms.Button()
            Me.btnConduct = New System.Windows.Forms.Button()
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
            Me.pgUnitTesting = New System.Windows.Forms.TabPage()
            Me.sccMain = New DevExpress.XtraEditors.SplitContainerControl()
            Me.sccTesting = New DevExpress.XtraEditors.SplitContainerControl()
            Me.SplitContainerControl1 = New DevExpress.XtraEditors.SplitContainerControl()
            Me.XtraTabControl1 = New DevExpress.XtraTab.XtraTabControl()
            Me.XtraTabPage1 = New DevExpress.XtraTab.XtraTabPage()
            Me.seNetwork = New Testing_Winform.UnitTesting.SimpleExplorer()
            Me.XtraTabPage2 = New DevExpress.XtraTab.XtraTabPage()
            Me.seSA = New Testing_Winform.UnitTesting.SimpleExplorer()
            Me.XtraTabControl2 = New DevExpress.XtraTab.XtraTabControl()
            Me.XtraTabPage3 = New DevExpress.XtraTab.XtraTabPage()
            Me.seLocal = New Testing_Winform.UnitTesting.SimpleExplorer()
            Me.SplitContainerControl2 = New DevExpress.XtraEditors.SplitContainerControl()
            Me.rtbNotes = New System.Windows.Forms.RichTextBox()
            Me.LabelControl14 = New DevExpress.XtraEditors.LabelControl()
            Me.mainLogViewer = New Testing_Winform.UnitTesting.LogViewer()
            Me.rtfactivityLog = New System.Windows.Forms.RichTextBox()
            Me.gcViewer = New DevExpress.XtraGrid.GridControl()
            Me.GridView1 = New DevExpress.XtraGrid.Views.Grid.GridView()
            Me.SplitterControl1 = New DevExpress.XtraEditors.SplitterControl()
            Me.pgcUnitTesting = New System.Windows.Forms.PropertyGrid()
            Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
            Me.CheckEditAutoReport = New DevExpress.XtraEditors.CheckEdit()
            Me.LabelControl12 = New DevExpress.XtraEditors.LabelControl()
            Me.btnProcess23 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnChecking = New DevExpress.XtraEditors.SimpleButton()
            Me.btnAuto = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess21 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess22 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess20 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess19 = New DevExpress.XtraEditors.SimpleButton()
            Me.CheckEditExcelVisibleII = New DevExpress.XtraEditors.CheckEdit()
            Me.LabelControl11 = New DevExpress.XtraEditors.LabelControl()
            Me.LabelConductOptions = New DevExpress.XtraEditors.LabelControl()
            Me.CheckEditExcelVisible = New DevExpress.XtraEditors.CheckEdit()
            Me.CheckEditDevMode = New DevExpress.XtraEditors.CheckEdit()
            Me.btnProcess18 = New DevExpress.XtraEditors.SimpleButton()
            Me.chkWorkLocal = New DevExpress.XtraEditors.CheckEdit()
            Me.btnProcess17 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess16 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess15 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess14 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess13 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnCheckout = New DevExpress.XtraEditors.SimpleButton()
            Me.testPull = New DevExpress.XtraEditors.SimpleButton()
            Me.testPush = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess12 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess11 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess9 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess10 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnClose = New DevExpress.XtraEditors.SimpleButton()
            Me.testBugFile = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess8 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess7 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess6 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess5 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess4 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess3 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess2 = New DevExpress.XtraEditors.SimpleButton()
            Me.btnProcess1 = New DevExpress.XtraEditors.SimpleButton()
            Me.testLocalWorkarea = New DevExpress.XtraEditors.TextEdit()
            Me.LabelControl10 = New DevExpress.XtraEditors.LabelControl()
            Me.testComb = New DevExpress.XtraEditors.TextEdit()
            Me.LabelControl9 = New DevExpress.XtraEditors.LabelControl()
            Me.testNextIteration = New DevExpress.XtraEditors.TextEdit()
            Me.LabelControl8 = New DevExpress.XtraEditors.LabelControl()
            Me.testID = New DevExpress.XtraEditors.ComboBoxEdit()
            Me.LabelControl7 = New DevExpress.XtraEditors.LabelControl()
            Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
            Me.testSaFolder = New DevExpress.XtraEditors.ButtonEdit()
            Me.testFolder = New DevExpress.XtraEditors.TextEdit()
            Me.testBu = New DevExpress.XtraEditors.TextEdit()
            Me.testIteration = New DevExpress.XtraEditors.TextEdit()
            Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
            Me.LabelControl6 = New DevExpress.XtraEditors.LabelControl()
            Me.testSid = New DevExpress.XtraEditors.TextEdit()
            Me.LabelControl5 = New DevExpress.XtraEditors.LabelControl()
            Me.testWo = New DevExpress.XtraEditors.TextEdit()
            Me.LabelControl4 = New DevExpress.XtraEditors.LabelControl()
            Me.LabelControl3 = New DevExpress.XtraEditors.LabelControl()
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
            Me.pgUnitTesting.SuspendLayout()
            CType(Me.sccMain, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.sccMain.Panel1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sccMain.Panel1.SuspendLayout()
            CType(Me.sccMain.Panel2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sccMain.Panel2.SuspendLayout()
            Me.sccMain.SuspendLayout()
            CType(Me.sccTesting, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.sccTesting.Panel1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sccTesting.Panel1.SuspendLayout()
            CType(Me.sccTesting.Panel2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sccTesting.Panel2.SuspendLayout()
            Me.sccTesting.SuspendLayout()
            CType(Me.SplitContainerControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.SplitContainerControl1.Panel1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SplitContainerControl1.Panel1.SuspendLayout()
            CType(Me.SplitContainerControl1.Panel2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SplitContainerControl1.Panel2.SuspendLayout()
            Me.SplitContainerControl1.SuspendLayout()
            CType(Me.XtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.XtraTabControl1.SuspendLayout()
            Me.XtraTabPage1.SuspendLayout()
            Me.XtraTabPage2.SuspendLayout()
            CType(Me.XtraTabControl2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.XtraTabControl2.SuspendLayout()
            Me.XtraTabPage3.SuspendLayout()
            CType(Me.SplitContainerControl2, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.SplitContainerControl2.Panel1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SplitContainerControl2.Panel1.SuspendLayout()
            CType(Me.SplitContainerControl2.Panel2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SplitContainerControl2.Panel2.SuspendLayout()
            Me.SplitContainerControl2.SuspendLayout()
            CType(Me.gcViewer, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.GridView1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.PanelControl1.SuspendLayout()
            CType(Me.CheckEditAutoReport.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.CheckEditExcelVisibleII.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.CheckEditExcelVisible.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.CheckEditDevMode.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.chkWorkLocal.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.testLocalWorkarea.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.testComb.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.testNextIteration.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.testID.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.testSaFolder.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.testFolder.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.testBu.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.testIteration.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.testSid.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.testWo.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'pgStructure
            '
            Me.pgStructure.Controls.Add(Me.SplitContainer3)
            Me.pgStructure.Controls.Add(Me.Panel2)
            Me.pgStructure.Location = New System.Drawing.Point(4, 22)
            Me.pgStructure.Name = "pgStructure"
            Me.pgStructure.Padding = New System.Windows.Forms.Padding(3)
            Me.pgStructure.Size = New System.Drawing.Size(1959, 742)
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
            Me.SplitContainer3.Size = New System.Drawing.Size(1953, 659)
            Me.SplitContainer3.SplitterDistance = 966
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
            Me.SplitContainer4.Panel1.Controls.Add(Me.btnLoopThroughERI)
            Me.SplitContainer4.Panel1.Controls.Add(Me.Button1)
            Me.SplitContainer4.Panel1.Controls.Add(Me.btnConduct)
            Me.SplitContainer4.Panel1.Controls.Add(Me.btnSaveFndToEDS)
            '
            'SplitContainer4.Panel2
            '
            Me.SplitContainer4.Panel2.Controls.Add(Me.propgridFndXL)
            Me.SplitContainer4.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.SplitContainer4.Size = New System.Drawing.Size(966, 659)
            Me.SplitContainer4.SplitterDistance = 164
            Me.SplitContainer4.TabIndex = 4
            '
            'btnLoopThroughERI
            '
            Me.btnLoopThroughERI.Location = New System.Drawing.Point(2, 319)
            Me.btnLoopThroughERI.Name = "btnLoopThroughERI"
            Me.btnLoopThroughERI.Size = New System.Drawing.Size(160, 21)
            Me.btnLoopThroughERI.TabIndex = 5
            Me.btnLoopThroughERI.Text = "Loop Through ERI"
            Me.btnLoopThroughERI.UseVisualStyleBackColor = True
            '
            'Button1
            '
            Me.Button1.Location = New System.Drawing.Point(2, 319)
            Me.Button1.Name = "Button1"
            Me.Button1.Size = New System.Drawing.Size(160, 21)
            Me.Button1.TabIndex = 5
            Me.Button1.Text = "tnx Loop Conductor"
            Me.Button1.UseVisualStyleBackColor = True
            '
            'btnConduct
            '
            Me.btnConduct.Location = New System.Drawing.Point(2, 179)
            Me.btnConduct.Name = "btnConduct"
            Me.btnConduct.Size = New System.Drawing.Size(160, 21)
            Me.btnConduct.TabIndex = 4
            Me.btnConduct.Text = "Conduct"
            Me.btnConduct.UseVisualStyleBackColor = True
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
            Me.propgridFndXL.Size = New System.Drawing.Size(798, 659)
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
            Me.SplitContainer5.Size = New System.Drawing.Size(983, 659)
            Me.SplitContainer5.SplitterDistance = 614
            Me.SplitContainer5.TabIndex = 0
            '
            'propgridFndEDS
            '
            Me.propgridFndEDS.Dock = System.Windows.Forms.DockStyle.Fill
            Me.propgridFndEDS.Location = New System.Drawing.Point(0, 0)
            Me.propgridFndEDS.Name = "propgridFndEDS"
            Me.propgridFndEDS.Size = New System.Drawing.Size(614, 659)
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
            Me.Panel2.Size = New System.Drawing.Size(1953, 77)
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
            Me.pgTNX.Size = New System.Drawing.Size(1959, 742)
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
            Me.SplitContainer1.Size = New System.Drawing.Size(1953, 691)
            Me.SplitContainer1.SplitterDistance = 966
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
            Me.scFromERI.Size = New System.Drawing.Size(966, 691)
            Me.scFromERI.SplitterDistance = 164
            Me.scFromERI.TabIndex = 4
            '
            'propgridTNXERI
            '
            Me.propgridTNXERI.Dock = System.Windows.Forms.DockStyle.Fill
            Me.propgridTNXERI.Location = New System.Drawing.Point(0, 0)
            Me.propgridTNXERI.Name = "propgridTNXERI"
            Me.propgridTNXERI.Size = New System.Drawing.Size(798, 691)
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
            Me.scFromEDS.Size = New System.Drawing.Size(983, 691)
            Me.scFromEDS.SplitterDistance = 614
            Me.scFromEDS.TabIndex = 0
            '
            'propgridTNXEDS
            '
            Me.propgridTNXEDS.Dock = System.Windows.Forms.DockStyle.Fill
            Me.propgridTNXEDS.Location = New System.Drawing.Point(0, 0)
            Me.propgridTNXEDS.Name = "propgridTNXEDS"
            Me.propgridTNXEDS.Size = New System.Drawing.Size(614, 691)
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
            Me.Panel1.Size = New System.Drawing.Size(1953, 45)
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
            Me.pgSQLBackUp.Size = New System.Drawing.Size(1959, 742)
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
            Me.TabControl1.Controls.Add(Me.pgUnitTesting)
            Me.TabControl1.Controls.Add(Me.pgStructure)
            Me.TabControl1.Controls.Add(Me.pgSQLBackUp)
            Me.TabControl1.Controls.Add(Me.pgTNX)
            Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.TabControl1.Location = New System.Drawing.Point(0, 0)
            Me.TabControl1.Name = "TabControl1"
            Me.TabControl1.SelectedIndex = 0
            Me.TabControl1.Size = New System.Drawing.Size(1967, 768)
            Me.TabControl1.TabIndex = 0
            '
            'pgUnitTesting
            '
            Me.pgUnitTesting.Controls.Add(Me.sccMain)
            Me.pgUnitTesting.Controls.Add(Me.PanelControl1)
            Me.pgUnitTesting.Location = New System.Drawing.Point(4, 22)
            Me.pgUnitTesting.Name = "pgUnitTesting"
            Me.pgUnitTesting.Padding = New System.Windows.Forms.Padding(3)
            Me.pgUnitTesting.Size = New System.Drawing.Size(1959, 742)
            Me.pgUnitTesting.TabIndex = 5
            Me.pgUnitTesting.Text = "Testing"
            Me.pgUnitTesting.UseVisualStyleBackColor = True
            '
            'sccMain
            '
            Me.sccMain.Dock = System.Windows.Forms.DockStyle.Fill
            Me.sccMain.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2
            Me.sccMain.Location = New System.Drawing.Point(3, 127)
            Me.sccMain.Name = "sccMain"
            '
            'sccMain.Panel1
            '
            Me.sccMain.Panel1.Controls.Add(Me.sccTesting)
            Me.sccMain.Panel1.Text = "Panel1"
            '
            'sccMain.Panel2
            '
            Me.sccMain.Panel2.Controls.Add(Me.gcViewer)
            Me.sccMain.Panel2.Controls.Add(Me.SplitterControl1)
            Me.sccMain.Panel2.Controls.Add(Me.pgcUnitTesting)
            Me.sccMain.Panel2.Text = "Panel2"
            Me.sccMain.Size = New System.Drawing.Size(1953, 612)
            Me.sccMain.SplitterPosition = 459
            Me.sccMain.TabIndex = 21
            '
            'sccTesting
            '
            Me.sccTesting.Dock = System.Windows.Forms.DockStyle.Fill
            Me.sccTesting.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2
            Me.sccTesting.Horizontal = False
            Me.sccTesting.Location = New System.Drawing.Point(0, 0)
            Me.sccTesting.Name = "sccTesting"
            '
            'sccTesting.Panel1
            '
            Me.sccTesting.Panel1.Controls.Add(Me.SplitContainerControl1)
            Me.sccTesting.Panel1.Text = "Panel1"
            '
            'sccTesting.Panel2
            '
            Me.sccTesting.Panel2.Controls.Add(Me.SplitContainerControl2)
            Me.sccTesting.Panel2.Text = "Panel2"
            Me.sccTesting.Size = New System.Drawing.Size(1484, 612)
            Me.sccTesting.SplitterPosition = 206
            Me.sccTesting.TabIndex = 18
            '
            'SplitContainerControl1
            '
            Me.SplitContainerControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.SplitContainerControl1.Location = New System.Drawing.Point(0, 0)
            Me.SplitContainerControl1.Name = "SplitContainerControl1"
            '
            'SplitContainerControl1.Panel1
            '
            Me.SplitContainerControl1.Panel1.Controls.Add(Me.XtraTabControl1)
            Me.SplitContainerControl1.Panel1.Text = "Panel1"
            '
            'SplitContainerControl1.Panel2
            '
            Me.SplitContainerControl1.Panel2.Controls.Add(Me.XtraTabControl2)
            Me.SplitContainerControl1.Panel2.Text = "Panel2"
            Me.SplitContainerControl1.Size = New System.Drawing.Size(1484, 396)
            Me.SplitContainerControl1.SplitterPosition = 435
            Me.SplitContainerControl1.TabIndex = 15
            '
            'XtraTabControl1
            '
            Me.XtraTabControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.XtraTabControl1.Enabled = False
            Me.XtraTabControl1.HeaderOrientation = DevExpress.XtraTab.TabOrientation.Horizontal
            Me.XtraTabControl1.Location = New System.Drawing.Point(0, 0)
            Me.XtraTabControl1.Name = "XtraTabControl1"
            Me.XtraTabControl1.SelectedTabPage = Me.XtraTabPage1
            Me.XtraTabControl1.ShowTabHeader = DevExpress.Utils.DefaultBoolean.[True]
            Me.XtraTabControl1.Size = New System.Drawing.Size(435, 396)
            Me.XtraTabControl1.TabIndex = 16
            Me.XtraTabControl1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPage1, Me.XtraTabPage2})
            '
            'XtraTabPage1
            '
            Me.XtraTabPage1.Controls.Add(Me.seNetwork)
            Me.XtraTabPage1.Name = "XtraTabPage1"
            Me.XtraTabPage1.Size = New System.Drawing.Size(433, 371)
            Me.XtraTabPage1.Text = "Network Test Folder"
            '
            'seNetwork
            '
            Me.seNetwork.Dock = System.Windows.Forms.DockStyle.Fill
            Me.seNetwork.Location = New System.Drawing.Point(0, 0)
            Me.seNetwork.Margin = New System.Windows.Forms.Padding(0)
            Me.seNetwork.Name = "seNetwork"
            Me.seNetwork.SelectedFile = Nothing
            Me.seNetwork.Size = New System.Drawing.Size(433, 371)
            Me.seNetwork.TabIndex = 16
            '
            'XtraTabPage2
            '
            Me.XtraTabPage2.Controls.Add(Me.seSA)
            Me.XtraTabPage2.Name = "XtraTabPage2"
            Me.XtraTabPage2.Size = New System.Drawing.Size(433, 371)
            Me.XtraTabPage2.Text = "SA Reference Folder"
            '
            'seSA
            '
            Me.seSA.Dock = System.Windows.Forms.DockStyle.Fill
            Me.seSA.Location = New System.Drawing.Point(0, 0)
            Me.seSA.Margin = New System.Windows.Forms.Padding(0)
            Me.seSA.Name = "seSA"
            Me.seSA.SelectedFile = Nothing
            Me.seSA.Size = New System.Drawing.Size(433, 371)
            Me.seSA.TabIndex = 17
            '
            'XtraTabControl2
            '
            Me.XtraTabControl2.Dock = System.Windows.Forms.DockStyle.Fill
            Me.XtraTabControl2.Location = New System.Drawing.Point(0, 0)
            Me.XtraTabControl2.Name = "XtraTabControl2"
            Me.XtraTabControl2.SelectedTabPage = Me.XtraTabPage3
            Me.XtraTabControl2.Size = New System.Drawing.Size(1039, 396)
            Me.XtraTabControl2.TabIndex = 17
            Me.XtraTabControl2.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.XtraTabPage3})
            '
            'XtraTabPage3
            '
            Me.XtraTabPage3.Controls.Add(Me.seLocal)
            Me.XtraTabPage3.Name = "XtraTabPage3"
            Me.XtraTabPage3.Size = New System.Drawing.Size(1037, 371)
            Me.XtraTabPage3.Text = "Local Testing Folder"
            '
            'seLocal
            '
            Me.seLocal.Dock = System.Windows.Forms.DockStyle.Fill
            Me.seLocal.Location = New System.Drawing.Point(0, 0)
            Me.seLocal.Margin = New System.Windows.Forms.Padding(0)
            Me.seLocal.Name = "seLocal"
            Me.seLocal.SelectedFile = Nothing
            Me.seLocal.Size = New System.Drawing.Size(1037, 371)
            Me.seLocal.TabIndex = 16
            '
            'SplitContainerControl2
            '
            Me.SplitContainerControl2.Dock = System.Windows.Forms.DockStyle.Fill
            Me.SplitContainerControl2.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2
            Me.SplitContainerControl2.Location = New System.Drawing.Point(0, 0)
            Me.SplitContainerControl2.Name = "SplitContainerControl2"
            '
            'SplitContainerControl2.Panel1
            '
            Me.SplitContainerControl2.Panel1.Controls.Add(Me.rtbNotes)
            Me.SplitContainerControl2.Panel1.Controls.Add(Me.LabelControl14)
            Me.SplitContainerControl2.Panel1.Text = "Panel1"
            '
            'SplitContainerControl2.Panel2
            '
            Me.SplitContainerControl2.Panel2.Controls.Add(Me.mainLogViewer)
            Me.SplitContainerControl2.Panel2.Controls.Add(Me.rtfactivityLog)
            Me.SplitContainerControl2.Panel2.Text = "Panel2"
            Me.SplitContainerControl2.Size = New System.Drawing.Size(1484, 206)
            Me.SplitContainerControl2.SplitterPosition = 764
            Me.SplitContainerControl2.TabIndex = 22
            '
            'rtbNotes
            '
            Me.rtbNotes.Dock = System.Windows.Forms.DockStyle.Fill
            Me.rtbNotes.Font = New System.Drawing.Font("Tahoma", 9.0!)
            Me.rtbNotes.Location = New System.Drawing.Point(0, 17)
            Me.rtbNotes.Name = "rtbNotes"
            Me.rtbNotes.Size = New System.Drawing.Size(710, 189)
            Me.rtbNotes.TabIndex = 21
            Me.rtbNotes.TabStop = False
            Me.rtbNotes.Text = ""
            Me.rtbNotes.WordWrap = False
            '
            'LabelControl14
            '
            Me.LabelControl14.Appearance.Font = New System.Drawing.Font("Tahoma", 10.25!)
            Me.LabelControl14.Appearance.Options.UseFont = True
            Me.LabelControl14.Dock = System.Windows.Forms.DockStyle.Top
            Me.LabelControl14.Location = New System.Drawing.Point(0, 0)
            Me.LabelControl14.Name = "LabelControl14"
            Me.LabelControl14.Size = New System.Drawing.Size(79, 17)
            Me.LabelControl14.TabIndex = 22
            Me.LabelControl14.Text = "TEST NOTES"
            '
            'mainLogViewer
            '
            Me.mainLogViewer.Dock = System.Windows.Forms.DockStyle.Fill
            Me.mainLogViewer.Enabled = False
            Me.mainLogViewer.Location = New System.Drawing.Point(0, 0)
            Me.mainLogViewer.Margin = New System.Windows.Forms.Padding(0)
            Me.mainLogViewer.Name = "mainLogViewer"
            Me.mainLogViewer.Size = New System.Drawing.Size(764, 206)
            Me.mainLogViewer.TabIndex = 23
            '
            'rtfactivityLog
            '
            Me.rtfactivityLog.Dock = System.Windows.Forms.DockStyle.Fill
            Me.rtfactivityLog.Location = New System.Drawing.Point(0, 0)
            Me.rtfactivityLog.Name = "rtfactivityLog"
            Me.rtfactivityLog.ReadOnly = True
            Me.rtfactivityLog.Size = New System.Drawing.Size(764, 206)
            Me.rtfactivityLog.TabIndex = 22
            Me.rtfactivityLog.Text = ""
            Me.rtfactivityLog.Visible = False
            Me.rtfactivityLog.WordWrap = False
            '
            'gcViewer
            '
            Me.gcViewer.Dock = System.Windows.Forms.DockStyle.Fill
            Me.gcViewer.Location = New System.Drawing.Point(0, 452)
            Me.gcViewer.MainView = Me.GridView1
            Me.gcViewer.Name = "gcViewer"
            Me.gcViewer.Size = New System.Drawing.Size(459, 160)
            Me.gcViewer.TabIndex = 20
            Me.gcViewer.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView1})
            '
            'GridView1
            '
            Me.GridView1.GridControl = Me.gcViewer
            Me.GridView1.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always
            Me.GridView1.Name = "GridView1"
            Me.GridView1.OptionsFilter.AllowAutoFilterConditionChange = DevExpress.Utils.DefaultBoolean.[False]
            Me.GridView1.OptionsFilter.AllowFilterEditor = False
            Me.GridView1.OptionsFilter.InHeaderSearchMode = DevExpress.XtraGrid.Views.Grid.GridInHeaderSearchMode.Disabled
            Me.GridView1.OptionsLayout.Columns.AddNewColumns = False
            Me.GridView1.OptionsLayout.Columns.RemoveOldColumns = False
            Me.GridView1.OptionsView.BestFitMode = DevExpress.XtraGrid.Views.Grid.GridBestFitMode.Full
            Me.GridView1.OptionsView.ColumnAutoWidth = False
            Me.GridView1.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.[True]
            Me.GridView1.OptionsView.ShowGroupPanel = False
            Me.GridView1.OptionsView.ShowIndicator = False
            '
            'SplitterControl1
            '
            Me.SplitterControl1.Cursor = System.Windows.Forms.Cursors.HSplit
            Me.SplitterControl1.Dock = System.Windows.Forms.DockStyle.Top
            Me.SplitterControl1.Location = New System.Drawing.Point(0, 442)
            Me.SplitterControl1.Name = "SplitterControl1"
            Me.SplitterControl1.Size = New System.Drawing.Size(459, 10)
            Me.SplitterControl1.TabIndex = 21
            Me.SplitterControl1.TabStop = False
            '
            'pgcUnitTesting
            '
            Me.pgcUnitTesting.Dock = System.Windows.Forms.DockStyle.Top
            Me.pgcUnitTesting.HelpVisible = False
            Me.pgcUnitTesting.Location = New System.Drawing.Point(0, 0)
            Me.pgcUnitTesting.Name = "pgcUnitTesting"
            Me.pgcUnitTesting.Size = New System.Drawing.Size(459, 442)
            Me.pgcUnitTesting.TabIndex = 0
            '
            'PanelControl1
            '
            Me.PanelControl1.Controls.Add(Me.CheckEditAutoReport)
            Me.PanelControl1.Controls.Add(Me.LabelControl12)
            Me.PanelControl1.Controls.Add(Me.btnProcess23)
            Me.PanelControl1.Controls.Add(Me.btnChecking)
            Me.PanelControl1.Controls.Add(Me.btnAuto)
            Me.PanelControl1.Controls.Add(Me.btnProcess21)
            Me.PanelControl1.Controls.Add(Me.btnProcess22)
            Me.PanelControl1.Controls.Add(Me.btnProcess20)
            Me.PanelControl1.Controls.Add(Me.btnProcess19)
            Me.PanelControl1.Controls.Add(Me.CheckEditExcelVisibleII)
            Me.PanelControl1.Controls.Add(Me.LabelControl11)
            Me.PanelControl1.Controls.Add(Me.LabelConductOptions)
            Me.PanelControl1.Controls.Add(Me.CheckEditExcelVisible)
            Me.PanelControl1.Controls.Add(Me.CheckEditDevMode)
            Me.PanelControl1.Controls.Add(Me.btnProcess18)
            Me.PanelControl1.Controls.Add(Me.chkWorkLocal)
            Me.PanelControl1.Controls.Add(Me.btnProcess17)
            Me.PanelControl1.Controls.Add(Me.btnProcess16)
            Me.PanelControl1.Controls.Add(Me.btnProcess15)
            Me.PanelControl1.Controls.Add(Me.btnProcess14)
            Me.PanelControl1.Controls.Add(Me.btnProcess13)
            Me.PanelControl1.Controls.Add(Me.btnCheckout)
            Me.PanelControl1.Controls.Add(Me.testPull)
            Me.PanelControl1.Controls.Add(Me.testPush)
            Me.PanelControl1.Controls.Add(Me.btnProcess12)
            Me.PanelControl1.Controls.Add(Me.btnProcess11)
            Me.PanelControl1.Controls.Add(Me.btnProcess9)
            Me.PanelControl1.Controls.Add(Me.btnProcess10)
            Me.PanelControl1.Controls.Add(Me.btnClose)
            Me.PanelControl1.Controls.Add(Me.testBugFile)
            Me.PanelControl1.Controls.Add(Me.btnProcess8)
            Me.PanelControl1.Controls.Add(Me.btnProcess7)
            Me.PanelControl1.Controls.Add(Me.btnProcess6)
            Me.PanelControl1.Controls.Add(Me.btnProcess5)
            Me.PanelControl1.Controls.Add(Me.btnProcess4)
            Me.PanelControl1.Controls.Add(Me.btnProcess3)
            Me.PanelControl1.Controls.Add(Me.btnProcess2)
            Me.PanelControl1.Controls.Add(Me.btnProcess1)
            Me.PanelControl1.Controls.Add(Me.testLocalWorkarea)
            Me.PanelControl1.Controls.Add(Me.LabelControl10)
            Me.PanelControl1.Controls.Add(Me.testComb)
            Me.PanelControl1.Controls.Add(Me.LabelControl9)
            Me.PanelControl1.Controls.Add(Me.testNextIteration)
            Me.PanelControl1.Controls.Add(Me.LabelControl8)
            Me.PanelControl1.Controls.Add(Me.testID)
            Me.PanelControl1.Controls.Add(Me.LabelControl7)
            Me.PanelControl1.Controls.Add(Me.LabelControl1)
            Me.PanelControl1.Controls.Add(Me.testSaFolder)
            Me.PanelControl1.Controls.Add(Me.testFolder)
            Me.PanelControl1.Controls.Add(Me.testBu)
            Me.PanelControl1.Controls.Add(Me.testIteration)
            Me.PanelControl1.Controls.Add(Me.LabelControl2)
            Me.PanelControl1.Controls.Add(Me.LabelControl6)
            Me.PanelControl1.Controls.Add(Me.testSid)
            Me.PanelControl1.Controls.Add(Me.LabelControl5)
            Me.PanelControl1.Controls.Add(Me.testWo)
            Me.PanelControl1.Controls.Add(Me.LabelControl4)
            Me.PanelControl1.Controls.Add(Me.LabelControl3)
            Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Top
            Me.PanelControl1.Location = New System.Drawing.Point(3, 3)
            Me.PanelControl1.Name = "PanelControl1"
            Me.PanelControl1.Size = New System.Drawing.Size(1953, 124)
            Me.PanelControl1.TabIndex = 16
            '
            'CheckEditAutoReport
            '
            Me.CheckEditAutoReport.Enabled = False
            Me.CheckEditAutoReport.Location = New System.Drawing.Point(1625, 97)
            Me.CheckEditAutoReport.Name = "CheckEditAutoReport"
            Me.CheckEditAutoReport.Properties.Caption = "Auto Mode"
            Me.CheckEditAutoReport.Size = New System.Drawing.Size(75, 20)
            Me.CheckEditAutoReport.TabIndex = 79
            '
            'LabelControl12
            '
            Me.LabelControl12.Appearance.Font = New System.Drawing.Font("Tahoma", 7.5!, System.Drawing.FontStyle.Bold)
            Me.LabelControl12.Appearance.Options.UseFont = True
            Me.LabelControl12.Location = New System.Drawing.Point(1625, 82)
            Me.LabelControl12.Name = "LabelControl12"
            Me.LabelControl12.Size = New System.Drawing.Size(78, 12)
            Me.LabelControl12.TabIndex = 78
            Me.LabelControl12.Text = "Report Options:"
            '
            'btnProcess23
            '
            Me.btnProcess23.Appearance.Options.UseTextOptions = True
            Me.btnProcess23.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess23.Enabled = False
            Me.btnProcess23.Location = New System.Drawing.Point(1322, 92)
            Me.btnProcess23.Name = "btnProcess23"
            Me.btnProcess23.Size = New System.Drawing.Size(90, 23)
            Me.btnProcess23.TabIndex = 77
            Me.btnProcess23.Tag = "STEP8|Create a report of the files in the selected folder."
            Me.btnProcess23.Text = "8. Gen. Report"
            Me.btnProcess23.ToolTip = "GENERATE REPORT"
            '
            'btnChecking
            '
            Me.btnChecking.Appearance.Options.UseTextOptions = True
            Me.btnChecking.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            Me.btnChecking.Enabled = False
            Me.btnChecking.Location = New System.Drawing.Point(1771, 12)
            Me.btnChecking.Name = "btnChecking"
            Me.btnChecking.Size = New System.Drawing.Size(145, 23)
            Me.btnChecking.TabIndex = 76
            Me.btnChecking.Tag = "Checking out it all"
            Me.btnChecking.Text = "Checkout Everything"
            Me.btnChecking.ToolTip = "CHECKOUT ALL TEST CASES"
            Me.btnChecking.Visible = False
            '
            'btnAuto
            '
            Me.btnAuto.Appearance.Options.UseTextOptions = True
            Me.btnAuto.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            Me.btnAuto.Enabled = False
            Me.btnAuto.Location = New System.Drawing.Point(1771, 40)
            Me.btnAuto.Name = "btnAuto"
            Me.btnAuto.Size = New System.Drawing.Size(145, 23)
            Me.btnAuto.TabIndex = 75
            Me.btnAuto.Tag = "Auto Test It"
            Me.btnAuto.Text = """Automated"" Unit Testing"
            Me.btnAuto.ToolTip = "AUTO TEST"
            Me.btnAuto.Visible = False
            '
            'btnProcess21
            '
            Me.btnProcess21.Appearance.Options.UseTextOptions = True
            Me.btnProcess21.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess21.Enabled = False
            Me.btnProcess21.Location = New System.Drawing.Point(1322, 64)
            Me.btnProcess21.Name = "btnProcess21"
            Me.btnProcess21.Size = New System.Drawing.Size(90, 23)
            Me.btnProcess21.TabIndex = 74
            Me.btnProcess21.Tag = "STEP7A|Generate and compare the results for all excel tools. "
            Me.btnProcess21.Text = "7a. Excel Results"
            Me.btnProcess21.ToolTip = "CREATE RESULTS (Excel)"
            '
            'btnProcess22
            '
            Me.btnProcess22.Appearance.Options.UseTextOptions = True
            Me.btnProcess22.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess22.Enabled = False
            Me.btnProcess22.Location = New System.Drawing.Point(1418, 64)
            Me.btnProcess22.Name = "btnProcess22"
            Me.btnProcess22.Size = New System.Drawing.Size(90, 23)
            Me.btnProcess22.TabIndex = 73
            Me.btnProcess22.Tag = "STEP7B|Generate and compare the results for all eri files. "
            Me.btnProcess22.Text = "7b. tnx Results"
            Me.btnProcess22.ToolTip = "CREATE RESULTS (TNX)"
            '
            'btnProcess20
            '
            Me.btnProcess20.Appearance.Options.UseTextOptions = True
            Me.btnProcess20.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess20.Enabled = False
            Me.btnProcess20.Location = New System.Drawing.Point(1625, 7)
            Me.btnProcess20.Name = "btnProcess20"
            Me.btnProcess20.Size = New System.Drawing.Size(91, 23)
            Me.btnProcess20.TabIndex = 72
            Me.btnProcess20.Tag = "STEP11|Conduct files created from EDS Data"
            Me.btnProcess20.Text = "11. Conduct EDS "
            Me.btnProcess20.ToolTip = "CONDUCT EDS FILES"
            '
            'btnProcess19
            '
            Me.btnProcess19.Appearance.Options.UseTextOptions = True
            Me.btnProcess19.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess19.Enabled = False
            Me.btnProcess19.Location = New System.Drawing.Point(1528, 7)
            Me.btnProcess19.Name = "btnProcess19"
            Me.btnProcess19.Size = New System.Drawing.Size(91, 23)
            Me.btnProcess19.TabIndex = 71
            Me.btnProcess19.Tag = "STEP10|Load the data saved to EDS into Excel sheets"
            Me.btnProcess19.Text = "10. Load EDS"
            Me.btnProcess19.ToolTip = "LOAD FROM EDS"
            '
            'CheckEditExcelVisibleII
            '
            Me.CheckEditExcelVisibleII.EditValue = True
            Me.CheckEditExcelVisibleII.Enabled = False
            Me.CheckEditExcelVisibleII.Location = New System.Drawing.Point(1625, 56)
            Me.CheckEditExcelVisibleII.Name = "CheckEditExcelVisibleII"
            Me.CheckEditExcelVisibleII.Properties.Caption = "Excel Visible"
            Me.CheckEditExcelVisibleII.Size = New System.Drawing.Size(84, 20)
            Me.CheckEditExcelVisibleII.TabIndex = 70
            '
            'LabelControl11
            '
            Me.LabelControl11.Appearance.Font = New System.Drawing.Font("Tahoma", 7.5!, System.Drawing.FontStyle.Bold)
            Me.LabelControl11.Appearance.Options.UseFont = True
            Me.LabelControl11.Location = New System.Drawing.Point(1625, 40)
            Me.LabelControl11.Name = "LabelControl11"
            Me.LabelControl11.Size = New System.Drawing.Size(114, 12)
            Me.LabelControl11.TabIndex = 69
            Me.LabelControl11.Text = "Import Inputs Options:"
            '
            'LabelConductOptions
            '
            Me.LabelConductOptions.Appearance.Font = New System.Drawing.Font("Tahoma", 7.5!, System.Drawing.FontStyle.Bold)
            Me.LabelConductOptions.Appearance.Options.UseFont = True
            Me.LabelConductOptions.Location = New System.Drawing.Point(1528, 38)
            Me.LabelConductOptions.Name = "LabelConductOptions"
            Me.LabelConductOptions.Size = New System.Drawing.Size(84, 12)
            Me.LabelConductOptions.TabIndex = 68
            Me.LabelConductOptions.Text = "Conduct Options:"
            '
            'CheckEditExcelVisible
            '
            Me.CheckEditExcelVisible.EditValue = True
            Me.CheckEditExcelVisible.Enabled = False
            Me.CheckEditExcelVisible.Location = New System.Drawing.Point(1528, 79)
            Me.CheckEditExcelVisible.Name = "CheckEditExcelVisible"
            Me.CheckEditExcelVisible.Properties.Caption = "Excel Visible"
            Me.CheckEditExcelVisible.Size = New System.Drawing.Size(84, 20)
            Me.CheckEditExcelVisible.TabIndex = 67
            '
            'CheckEditDevMode
            '
            Me.CheckEditDevMode.EditValue = True
            Me.CheckEditDevMode.Enabled = False
            Me.CheckEditDevMode.Location = New System.Drawing.Point(1528, 53)
            Me.CheckEditDevMode.Name = "CheckEditDevMode"
            Me.CheckEditDevMode.Properties.Caption = "Dev Mode"
            Me.CheckEditDevMode.Size = New System.Drawing.Size(75, 20)
            Me.CheckEditDevMode.TabIndex = 66
            '
            'btnProcess18
            '
            Me.btnProcess18.Appearance.Options.UseTextOptions = True
            Me.btnProcess18.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess18.Enabled = False
            Me.btnProcess18.Location = New System.Drawing.Point(1418, 91)
            Me.btnProcess18.Name = "btnProcess18"
            Me.btnProcess18.Size = New System.Drawing.Size(91, 23)
            Me.btnProcess18.TabIndex = 65
            Me.btnProcess18.Tag = "STEP9|Save local structure to EDS"
            Me.btnProcess18.Text = "9. Save to EDS"
            Me.btnProcess18.ToolTip = "SAVE TO EDS"
            '
            'chkWorkLocal
            '
            Me.chkWorkLocal.EditValue = True
            Me.chkWorkLocal.Enabled = False
            Me.chkWorkLocal.Location = New System.Drawing.Point(2051, 94)
            Me.chkWorkLocal.Name = "chkWorkLocal"
            Me.chkWorkLocal.Properties.Caption = "CheckEdit1"
            Me.chkWorkLocal.Size = New System.Drawing.Size(75, 20)
            Me.chkWorkLocal.TabIndex = 64
            Me.chkWorkLocal.Visible = False
            '
            'btnProcess17
            '
            Me.btnProcess17.Appearance.Options.UseTextOptions = True
            Me.btnProcess17.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess17.Enabled = False
            Me.btnProcess17.Location = New System.Drawing.Point(1243, 92)
            Me.btnProcess17.Name = "btnProcess17"
            Me.btnProcess17.Size = New System.Drawing.Size(58, 23)
            Me.btnProcess17.TabIndex = 63
            Me.btnProcess17.Tag = "STEP5C|Analyze the ERIs in the 'Manual (SAPI)' folder"
            Me.btnProcess17.Text = "5c. Man"
            Me.btnProcess17.ToolTip = "ANALYZE ERI"
            '
            'btnProcess16
            '
            Me.btnProcess16.Appearance.Options.UseTextOptions = True
            Me.btnProcess16.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess16.Enabled = False
            Me.btnProcess16.Location = New System.Drawing.Point(1179, 92)
            Me.btnProcess16.Name = "btnProcess16"
            Me.btnProcess16.Size = New System.Drawing.Size(58, 23)
            Me.btnProcess16.TabIndex = 62
            Me.btnProcess16.Tag = "STEP5B|Analyze the ERIs in the 'Manual (Current)' folder"
            Me.btnProcess16.Text = "5b. Cur."
            Me.btnProcess16.ToolTip = "ANALYZE ERI"
            '
            'btnProcess15
            '
            Me.btnProcess15.Appearance.Options.UseTextOptions = True
            Me.btnProcess15.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess15.Enabled = False
            Me.btnProcess15.Location = New System.Drawing.Point(1115, 92)
            Me.btnProcess15.Name = "btnProcess15"
            Me.btnProcess15.Size = New System.Drawing.Size(58, 23)
            Me.btnProcess15.TabIndex = 61
            Me.btnProcess15.Tag = "STEP5A|Analyze the ERIs in the 'Manual ERI' Folder"
            Me.btnProcess15.Text = "5a. ERI"
            Me.btnProcess15.ToolTip = "ANALYZE ERI"
            '
            'btnProcess14
            '
            Me.btnProcess14.Appearance.Options.UseTextOptions = True
            Me.btnProcess14.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess14.Enabled = False
            Me.btnProcess14.Location = New System.Drawing.Point(1243, 34)
            Me.btnProcess14.Name = "btnProcess14"
            Me.btnProcess14.Size = New System.Drawing.Size(58, 23)
            Me.btnProcess14.TabIndex = 54
            Me.btnProcess14.Tag = "STEP4C|Import inputs for Manual SAPI versions in the iteration folder"
            Me.btnProcess14.Text = "4c. II Man"
            Me.btnProcess14.ToolTip = "IMPORT INPUTS (SAPI Manual)"
            '
            'btnProcess13
            '
            Me.btnProcess13.Enabled = False
            Me.btnProcess13.Location = New System.Drawing.Point(1038, 92)
            Me.btnProcess13.Name = "btnProcess13"
            Me.btnProcess13.Size = New System.Drawing.Size(55, 23)
            Me.btnProcess13.TabIndex = 53
            Me.btnProcess13.Tag = "STEP3D|Create new Manual SAPI teamplates in the iteration folder."
            Me.btnProcess13.Text = "3d. Man"
            Me.btnProcess13.ToolTip = "CREATE TEMPLATE FILES (SAPI Manual)"
            '
            'btnCheckout
            '
            Me.btnCheckout.ImageOptions.SvgImage = Global.Testing_Winform.My.Resources.Resources.shopping_shoppingcart
            Me.btnCheckout.ImageOptions.SvgImageSize = New System.Drawing.Size(15, 15)
            Me.btnCheckout.Location = New System.Drawing.Point(226, 5)
            Me.btnCheckout.Name = "btnCheckout"
            Me.btnCheckout.Size = New System.Drawing.Size(78, 23)
            Me.btnCheckout.TabIndex = 52
            Me.btnCheckout.Tag = "START|"
            Me.btnCheckout.Text = "Checkout"
            '
            'testPull
            '
            Me.testPull.Enabled = False
            Me.testPull.ImageOptions.SvgImage = Global.Testing_Winform.My.Resources.Resources.movedown
            Me.testPull.ImageOptions.SvgImageSize = New System.Drawing.Size(15, 15)
            Me.testPull.Location = New System.Drawing.Point(556, 87)
            Me.testPull.Name = "testPull"
            Me.testPull.Size = New System.Drawing.Size(60, 23)
            Me.testPull.TabIndex = 51
            Me.testPull.Text = "Pull"
            '
            'testPush
            '
            Me.testPush.Enabled = False
            Me.testPush.ImageOptions.SvgImage = Global.Testing_Winform.My.Resources.Resources.moveup
            Me.testPush.ImageOptions.SvgImageSize = New System.Drawing.Size(15, 15)
            Me.testPush.Location = New System.Drawing.Point(556, 58)
            Me.testPush.Name = "testPush"
            Me.testPush.Size = New System.Drawing.Size(60, 23)
            Me.testPush.TabIndex = 50
            Me.testPush.Text = "Push"
            '
            'btnProcess12
            '
            Me.btnProcess12.Appearance.Options.UseTextOptions = True
            Me.btnProcess12.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess12.Enabled = False
            Me.btnProcess12.Location = New System.Drawing.Point(1115, 5)
            Me.btnProcess12.Name = "btnProcess12"
            Me.btnProcess12.Size = New System.Drawing.Size(186, 23)
            Me.btnProcess12.TabIndex = 48
            Me.btnProcess12.Tag = "STEP4|Import inputs for all files for the current iteration and publsihed version" &
    "s."
            Me.btnProcess12.Text = "4. Import Inputs (II)"
            Me.btnProcess12.ToolTip = "IMPORT INPUTS (BOTH)"
            '
            'btnProcess11
            '
            Me.btnProcess11.Enabled = False
            Me.btnProcess11.Location = New System.Drawing.Point(977, 92)
            Me.btnProcess11.Name = "btnProcess11"
            Me.btnProcess11.Size = New System.Drawing.Size(55, 23)
            Me.btnProcess11.TabIndex = 47
            Me.btnProcess11.Tag = "STEP3C|Create new Maestro SAPI teamplates in the iteration folder."
            Me.btnProcess11.Text = "3c. Mae"
            Me.btnProcess11.ToolTip = "CREATE TEMPLATE FILES (SAPI Maestro)"
            '
            'btnProcess9
            '
            Me.btnProcess9.Enabled = False
            Me.btnProcess9.Location = New System.Drawing.Point(855, 92)
            Me.btnProcess9.Name = "btnProcess9"
            Me.btnProcess9.Size = New System.Drawing.Size(55, 23)
            Me.btnProcess9.TabIndex = 46
            Me.btnProcess9.Tag = "STEP3A|Create new production version templates"
            Me.btnProcess9.Text = "3a. Cur."
            Me.btnProcess9.ToolTip = "CREATE TEMPLATE FILES (PROD)"
            '
            'btnProcess10
            '
            Me.btnProcess10.Enabled = False
            Me.btnProcess10.Location = New System.Drawing.Point(916, 92)
            Me.btnProcess10.Name = "btnProcess10"
            Me.btnProcess10.Size = New System.Drawing.Size(55, 23)
            Me.btnProcess10.TabIndex = 45
            Me.btnProcess10.Tag = "STEP3B|Create new ERI file(s) in the 'Manual ERI' folder."
            Me.btnProcess10.Text = "3b. ERI"
            Me.btnProcess10.ToolTip = "CREATE TEAMPLATE FILES (ERI)"
            '
            'btnClose
            '
            Me.btnClose.Enabled = False
            Me.btnClose.ImageOptions.SvgImage = CType(resources.GetObject("btnClose.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
            Me.btnClose.ImageOptions.SvgImageSize = New System.Drawing.Size(15, 15)
            Me.btnClose.Location = New System.Drawing.Point(317, 5)
            Me.btnClose.Name = "btnClose"
            Me.btnClose.Size = New System.Drawing.Size(75, 23)
            Me.btnClose.TabIndex = 43
            Me.btnClose.Tag = "FINISH|"
            Me.btnClose.Text = "Close"
            '
            'testBugFile
            '
            Me.testBugFile.Enabled = False
            Me.testBugFile.ImageOptions.Image = CType(resources.GetObject("testBugFile.ImageOptions.Image"), System.Drawing.Image)
            Me.testBugFile.ImageOptions.SvgImageSize = New System.Drawing.Size(15, 15)
            Me.testBugFile.Location = New System.Drawing.Point(405, 5)
            Me.testBugFile.Name = "testBugFile"
            Me.testBugFile.Size = New System.Drawing.Size(139, 23)
            Me.testBugFile.TabIndex = 42
            Me.testBugFile.Text = "Add Bug Reference File"
            '
            'btnProcess8
            '
            Me.btnProcess8.Appearance.Options.UseTextOptions = True
            Me.btnProcess8.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess8.Enabled = False
            Me.btnProcess8.Location = New System.Drawing.Point(1322, 34)
            Me.btnProcess8.Name = "btnProcess8"
            Me.btnProcess8.Size = New System.Drawing.Size(186, 23)
            Me.btnProcess8.TabIndex = 39
            Me.btnProcess8.Tag = "STEP7|Generate and compare the results. "
            Me.btnProcess8.Text = "7. Create Results"
            Me.btnProcess8.ToolTip = "CREATE RESULTS"
            '
            'btnProcess7
            '
            Me.btnProcess7.Appearance.Options.UseTextOptions = True
            Me.btnProcess7.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess7.Enabled = False
            Me.btnProcess7.Location = New System.Drawing.Point(1322, 5)
            Me.btnProcess7.Name = "btnProcess7"
            Me.btnProcess7.Size = New System.Drawing.Size(186, 23)
            Me.btnProcess7.TabIndex = 38
            Me.btnProcess7.Tag = "STEP6|Conduct the files in the Maestro folder. "
            Me.btnProcess7.Text = "6. Conduct Maestro Files"
            Me.btnProcess7.ToolTip = "CONDUCT"
            '
            'btnProcess6
            '
            Me.btnProcess6.Appearance.Options.UseTextOptions = True
            Me.btnProcess6.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess6.Enabled = False
            Me.btnProcess6.Location = New System.Drawing.Point(1115, 63)
            Me.btnProcess6.Name = "btnProcess6"
            Me.btnProcess6.Size = New System.Drawing.Size(186, 23)
            Me.btnProcess6.TabIndex = 37
            Me.btnProcess6.Tag = "STEP5|Analyze the all ERIs for the test case other than the Maestro folder."
            Me.btnProcess6.Text = "5. Run All ERI(s)"
            Me.btnProcess6.ToolTip = "ANALYZE ERI"
            '
            'btnProcess5
            '
            Me.btnProcess5.Appearance.Options.UseTextOptions = True
            Me.btnProcess5.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess5.Enabled = False
            Me.btnProcess5.Location = New System.Drawing.Point(1115, 34)
            Me.btnProcess5.Name = "btnProcess5"
            Me.btnProcess5.Size = New System.Drawing.Size(58, 23)
            Me.btnProcess5.TabIndex = 36
            Me.btnProcess5.Tag = "STEP4A|Import inputs for Maestro SAPI versions in the iteration folder"
            Me.btnProcess5.Text = "4a. II Mae"
            Me.btnProcess5.ToolTip = "IMPORT INPUTS (SAPI Maestro)"
            '
            'btnProcess4
            '
            Me.btnProcess4.Appearance.Options.UseTextOptions = True
            Me.btnProcess4.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess4.Enabled = False
            Me.btnProcess4.Location = New System.Drawing.Point(1179, 34)
            Me.btnProcess4.Name = "btnProcess4"
            Me.btnProcess4.Size = New System.Drawing.Size(58, 23)
            Me.btnProcess4.TabIndex = 35
            Me.btnProcess4.Tag = "STEP4B|Import inputs for current production tools)"
            Me.btnProcess4.Text = "4b. II Cur."
            Me.btnProcess4.ToolTip = "IMPORT INPUTS (CURRENT)"
            '
            'btnProcess3
            '
            Me.btnProcess3.Appearance.Options.UseTextOptions = True
            Me.btnProcess3.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess3.Enabled = False
            Me.btnProcess3.Location = New System.Drawing.Point(855, 63)
            Me.btnProcess3.Name = "btnProcess3"
            Me.btnProcess3.Size = New System.Drawing.Size(238, 23)
            Me.btnProcess3.TabIndex = 34
            Me.btnProcess3.Tag = "STEP3|Create new templates (Only SAPI if current files already exist)"
            Me.btnProcess3.Text = "3. Create Template Files"
            Me.btnProcess3.ToolTip = "CREATE TEMPLATE FILES"
            '
            'btnProcess2
            '
            Me.btnProcess2.Appearance.Options.UseTextOptions = True
            Me.btnProcess2.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess2.Enabled = False
            Me.btnProcess2.Location = New System.Drawing.Point(855, 34)
            Me.btnProcess2.Name = "btnProcess2"
            Me.btnProcess2.Size = New System.Drawing.Size(238, 23)
            Me.btnProcess2.TabIndex = 33
            Me.btnProcess2.Tag = "STEP2|Create a new iteration (Maestro & Manual (SAPI)) folders"
            Me.btnProcess2.Text = "2. Create New Iteration"
            Me.btnProcess2.ToolTip = "CREATE ITERATION"
            '
            'btnProcess1
            '
            Me.btnProcess1.Appearance.Options.UseTextOptions = True
            Me.btnProcess1.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.btnProcess1.Enabled = False
            Me.btnProcess1.Location = New System.Drawing.Point(855, 5)
            Me.btnProcess1.Name = "btnProcess1"
            Me.btnProcess1.Size = New System.Drawing.Size(238, 23)
            Me.btnProcess1.TabIndex = 32
            Me.btnProcess1.Tag = "STEP1|Get all files from reference SA file from specified folder"
            Me.btnProcess1.Text = "1. Get Reference SA Files"
            Me.btnProcess1.ToolTip = "GET SA FILES"
            '
            'testLocalWorkarea
            '
            Me.testLocalWorkarea.Location = New System.Drawing.Point(93, 87)
            Me.testLocalWorkarea.Name = "testLocalWorkarea"
            Me.testLocalWorkarea.Size = New System.Drawing.Size(451, 20)
            Me.testLocalWorkarea.TabIndex = 22
            '
            'LabelControl10
            '
            Me.LabelControl10.Location = New System.Drawing.Point(9, 90)
            Me.LabelControl10.Name = "LabelControl10"
            Me.LabelControl10.Size = New System.Drawing.Size(78, 13)
            Me.LabelControl10.TabIndex = 21
            Me.LabelControl10.Text = "Local Work Area"
            '
            'testComb
            '
            Me.testComb.Location = New System.Drawing.Point(650, 35)
            Me.testComb.Name = "testComb"
            Me.testComb.Properties.ReadOnly = True
            Me.testComb.Size = New System.Drawing.Size(180, 20)
            Me.testComb.TabIndex = 20
            '
            'LabelControl9
            '
            Me.LabelControl9.Location = New System.Drawing.Point(562, 38)
            Me.LabelControl9.Name = "LabelControl9"
            Me.LabelControl9.Size = New System.Drawing.Size(83, 13)
            Me.LabelControl9.TabIndex = 19
            Me.LabelControl9.Text = "Test Combination"
            '
            'testNextIteration
            '
            Me.testNextIteration.Location = New System.Drawing.Point(780, 5)
            Me.testNextIteration.Name = "testNextIteration"
            Me.testNextIteration.Properties.ReadOnly = True
            Me.testNextIteration.Size = New System.Drawing.Size(50, 20)
            Me.testNextIteration.TabIndex = 17
            Me.testNextIteration.Visible = False
            '
            'LabelControl8
            '
            Me.LabelControl8.Location = New System.Drawing.Point(706, 8)
            Me.LabelControl8.Name = "LabelControl8"
            Me.LabelControl8.Size = New System.Drawing.Size(68, 13)
            Me.LabelControl8.TabIndex = 16
            Me.LabelControl8.Text = "Next Iteration"
            Me.LabelControl8.Visible = False
            '
            'testID
            '
            Me.testID.Location = New System.Drawing.Point(93, 9)
            Me.testID.Name = "testID"
            Me.testID.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
            Me.testID.Properties.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99", "100"})
            Me.testID.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            Me.testID.Size = New System.Drawing.Size(116, 20)
            Me.testID.TabIndex = 15
            '
            'LabelControl7
            '
            Me.LabelControl7.Location = New System.Drawing.Point(52, 12)
            Me.LabelControl7.Name = "LabelControl7"
            Me.LabelControl7.Size = New System.Drawing.Size(35, 13)
            Me.LabelControl7.TabIndex = 14
            Me.LabelControl7.Text = "Test ID"
            '
            'LabelControl1
            '
            Me.LabelControl1.Location = New System.Drawing.Point(631, 64)
            Me.LabelControl1.Name = "LabelControl1"
            Me.LabelControl1.Size = New System.Drawing.Size(13, 13)
            Me.LabelControl1.TabIndex = 3
            Me.LabelControl1.Text = "BU"
            '
            'testSaFolder
            '
            Me.testSaFolder.EditValue = ""
            Me.testSaFolder.Location = New System.Drawing.Point(93, 35)
            Me.testSaFolder.Name = "testSaFolder"
            Me.testSaFolder.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
            Me.testSaFolder.Size = New System.Drawing.Size(451, 20)
            Me.testSaFolder.TabIndex = 1
            '
            'testFolder
            '
            Me.testFolder.Location = New System.Drawing.Point(93, 61)
            Me.testFolder.Name = "testFolder"
            Me.testFolder.Properties.ReadOnly = True
            Me.testFolder.Size = New System.Drawing.Size(451, 20)
            Me.testFolder.TabIndex = 13
            '
            'testBu
            '
            Me.testBu.Location = New System.Drawing.Point(650, 61)
            Me.testBu.Name = "testBu"
            Me.testBu.Properties.ReadOnly = True
            Me.testBu.Size = New System.Drawing.Size(100, 20)
            Me.testBu.TabIndex = 2
            '
            'testIteration
            '
            Me.testIteration.EditValue = "13"
            Me.testIteration.Location = New System.Drawing.Point(650, 5)
            Me.testIteration.Name = "testIteration"
            Me.testIteration.Properties.Appearance.Options.UseTextOptions = True
            Me.testIteration.Properties.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
            Me.testIteration.Properties.ReadOnly = True
            Me.testIteration.Size = New System.Drawing.Size(50, 20)
            Me.testIteration.TabIndex = 12
            '
            'LabelControl2
            '
            Me.LabelControl2.Location = New System.Drawing.Point(753, 65)
            Me.LabelControl2.Name = "LabelControl2"
            Me.LabelControl2.Size = New System.Drawing.Size(37, 13)
            Me.LabelControl2.TabIndex = 4
            Me.LabelControl2.Text = "Strc. ID"
            '
            'LabelControl6
            '
            Me.LabelControl6.Location = New System.Drawing.Point(562, 8)
            Me.LabelControl6.Name = "LabelControl6"
            Me.LabelControl6.Size = New System.Drawing.Size(82, 13)
            Me.LabelControl6.TabIndex = 11
            Me.LabelControl6.Text = "Current Iteration"
            '
            'testSid
            '
            Me.testSid.Location = New System.Drawing.Point(794, 62)
            Me.testSid.Name = "testSid"
            Me.testSid.Properties.ReadOnly = True
            Me.testSid.Size = New System.Drawing.Size(36, 20)
            Me.testSid.TabIndex = 5
            '
            'LabelControl5
            '
            Me.LabelControl5.Location = New System.Drawing.Point(33, 64)
            Me.LabelControl5.Name = "LabelControl5"
            Me.LabelControl5.Size = New System.Drawing.Size(54, 13)
            Me.LabelControl5.TabIndex = 10
            Me.LabelControl5.Text = "Test Folder"
            '
            'testWo
            '
            Me.testWo.Location = New System.Drawing.Point(650, 87)
            Me.testWo.Name = "testWo"
            Me.testWo.Properties.ReadOnly = True
            Me.testWo.Size = New System.Drawing.Size(100, 20)
            Me.testWo.TabIndex = 6
            '
            'LabelControl4
            '
            Me.LabelControl4.Location = New System.Drawing.Point(41, 38)
            Me.LabelControl4.Name = "LabelControl4"
            Me.LabelControl4.Size = New System.Drawing.Size(46, 13)
            Me.LabelControl4.TabIndex = 8
            Me.LabelControl4.Text = "SA Folder"
            '
            'LabelControl3
            '
            Me.LabelControl3.Location = New System.Drawing.Point(626, 90)
            Me.LabelControl3.Name = "LabelControl3"
            Me.LabelControl3.Size = New System.Drawing.Size(18, 13)
            Me.LabelControl3.TabIndex = 7
            Me.LabelControl3.Text = "WO"
            '
            'frmMain
            '
            Me.Appearance.BackColor = System.Drawing.Color.White
            Me.Appearance.Options.UseBackColor = True
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.AutoScroll = True
            Me.ClientSize = New System.Drawing.Size(1967, 768)
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
            Me.pgUnitTesting.ResumeLayout(False)
            CType(Me.sccMain.Panel1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccMain.Panel1.ResumeLayout(False)
            CType(Me.sccMain.Panel2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccMain.Panel2.ResumeLayout(False)
            CType(Me.sccMain, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccMain.ResumeLayout(False)
            CType(Me.sccTesting.Panel1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccTesting.Panel1.ResumeLayout(False)
            CType(Me.sccTesting.Panel2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccTesting.Panel2.ResumeLayout(False)
            CType(Me.sccTesting, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccTesting.ResumeLayout(False)
            CType(Me.SplitContainerControl1.Panel1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.SplitContainerControl1.Panel1.ResumeLayout(False)
            CType(Me.SplitContainerControl1.Panel2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.SplitContainerControl1.Panel2.ResumeLayout(False)
            CType(Me.SplitContainerControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.SplitContainerControl1.ResumeLayout(False)
            CType(Me.XtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.XtraTabControl1.ResumeLayout(False)
            Me.XtraTabPage1.ResumeLayout(False)
            Me.XtraTabPage2.ResumeLayout(False)
            CType(Me.XtraTabControl2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.XtraTabControl2.ResumeLayout(False)
            Me.XtraTabPage3.ResumeLayout(False)
            CType(Me.SplitContainerControl2.Panel1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.SplitContainerControl2.Panel1.ResumeLayout(False)
            Me.SplitContainerControl2.Panel1.PerformLayout()
            CType(Me.SplitContainerControl2.Panel2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.SplitContainerControl2.Panel2.ResumeLayout(False)
            CType(Me.SplitContainerControl2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.SplitContainerControl2.ResumeLayout(False)
            CType(Me.gcViewer, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.GridView1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.PanelControl1.ResumeLayout(False)
            Me.PanelControl1.PerformLayout()
            CType(Me.CheckEditAutoReport.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.CheckEditExcelVisibleII.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.CheckEditExcelVisible.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.CheckEditDevMode.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.chkWorkLocal.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.testLocalWorkarea.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.testComb.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.testNextIteration.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.testID.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.testSaFolder.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.testFolder.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.testBu.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.testIteration.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.testSid.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.testWo.Properties, System.ComponentModel.ISupportInitialize).EndInit()
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
        Friend WithEvents btnConduct As Button
        Friend WithEvents pgUnitTesting As TabPage
        Friend WithEvents testSaFolder As DevExpress.XtraEditors.ButtonEdit
        Friend WithEvents testIteration As DevExpress.XtraEditors.TextEdit
        Friend WithEvents LabelControl6 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents LabelControl5 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents LabelControl4 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents LabelControl3 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents testWo As DevExpress.XtraEditors.TextEdit
        Friend WithEvents testSid As DevExpress.XtraEditors.TextEdit
        Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents testBu As DevExpress.XtraEditors.TextEdit
        Friend WithEvents testFolder As DevExpress.XtraEditors.TextEdit
        Friend WithEvents SplitContainerControl1 As DevExpress.XtraEditors.SplitContainerControl
        Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
        Friend WithEvents testID As DevExpress.XtraEditors.ComboBoxEdit
        Friend WithEvents LabelControl7 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents testNextIteration As DevExpress.XtraEditors.TextEdit
        Friend WithEvents LabelControl8 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents testComb As DevExpress.XtraEditors.TextEdit
        Friend WithEvents LabelControl9 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents sccTesting As DevExpress.XtraEditors.SplitContainerControl
        Friend WithEvents testLocalWorkarea As DevExpress.XtraEditors.TextEdit
        Friend WithEvents LabelControl10 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents sccMain As DevExpress.XtraEditors.SplitContainerControl
        Friend WithEvents pgcUnitTesting As PropertyGrid
        Friend WithEvents XtraTabControl1 As DevExpress.XtraTab.XtraTabControl
        Friend WithEvents XtraTabPage1 As DevExpress.XtraTab.XtraTabPage
        Friend WithEvents seNetwork As SimpleExplorer
        Friend WithEvents XtraTabPage2 As DevExpress.XtraTab.XtraTabPage
        Friend WithEvents Button1 As Button
        Friend WithEvents btnLoopThroughERI As Button
        Friend WithEvents seSA As SimpleExplorer
        Friend WithEvents XtraTabControl2 As DevExpress.XtraTab.XtraTabControl
        Friend WithEvents XtraTabPage3 As DevExpress.XtraTab.XtraTabPage
        Friend WithEvents seLocal As SimpleExplorer
        Friend WithEvents gcViewer As DevExpress.XtraGrid.GridControl
        Friend WithEvents GridView1 As DevExpress.XtraGrid.Views.Grid.GridView
        Friend WithEvents Panel2 As Panel
        Friend WithEvents Panel1 As Panel
        Friend WithEvents SplitterControl1 As DevExpress.XtraEditors.SplitterControl
        Friend WithEvents btnProcess8 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess7 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess6 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess5 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess4 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess3 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess2 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess1 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents SplitContainerControl2 As DevExpress.XtraEditors.SplitContainerControl
        Friend WithEvents rtbNotes As RichTextBox
        Friend WithEvents LabelControl14 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents rtfactivityLog As RichTextBox
        Friend WithEvents testBugFile As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents mainLogViewer As LogViewer
        Friend WithEvents btnClose As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess12 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess11 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess9 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess10 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents testPull As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents testPush As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnCheckout As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess14 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess13 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess17 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess16 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess15 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents chkWorkLocal As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents btnProcess18 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents LabelConductOptions As DevExpress.XtraEditors.LabelControl
        Friend WithEvents CheckEditExcelVisible As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents CheckEditDevMode As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents CheckEditExcelVisibleII As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents LabelControl11 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents btnProcess20 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess19 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess22 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess21 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnAuto As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnChecking As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents btnProcess23 As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents CheckEditAutoReport As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents LabelControl12 As DevExpress.XtraEditors.LabelControl
#End Region

    End Class
End Namespace