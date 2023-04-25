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
            Me.GridView2 = New DevExpress.XtraGrid.Views.Grid.GridView()
            Me.GridControl1 = New DevExpress.XtraGrid.GridControl()
            Me.GridView1 = New DevExpress.XtraGrid.Views.Grid.GridView()
            Me.GridView3 = New DevExpress.XtraGrid.Views.Grid.GridView()
            Me.pgStructure = New System.Windows.Forms.TabPage()
            Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
            Me.btnImportStrcFiles = New System.Windows.Forms.Button()
            Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
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
            Me.sccTesting = New DevExpress.XtraEditors.SplitContainerControl()
            Me.SplitContainerControl1 = New DevExpress.XtraEditors.SplitContainerControl()
            Me.sccNetwork = New DevExpress.XtraEditors.SplitContainerControl()
            Me.seNetwork = New Testing_Winform.UnitTesting.SimpleExplorer()
            Me.LabelControl11 = New DevExpress.XtraEditors.LabelControl()
            Me.seSA = New Testing_Winform.UnitTesting.SimpleExplorer()
            Me.LabelControl12 = New DevExpress.XtraEditors.LabelControl()
            Me.seLocal = New Testing_Winform.UnitTesting.SimpleExplorer()
            Me.LabelControl13 = New DevExpress.XtraEditors.LabelControl()
            Me.rtbNotes = New System.Windows.Forms.RichTextBox()
            Me.LabelControl14 = New DevExpress.XtraEditors.LabelControl()
            Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
            Me.testPublishedResults = New DevExpress.XtraEditors.SimpleButton()
            Me.testPrevResults = New DevExpress.XtraEditors.SimpleButton()
            Me.testIterationResults = New DevExpress.XtraEditors.SimpleButton()
            Me.testConduct = New DevExpress.XtraEditors.SimpleButton()
            Me.chkWorkLocal = New DevExpress.XtraEditors.CheckEdit()
            Me.testLocalWorkarea = New DevExpress.XtraEditors.TextEdit()
            Me.LabelControl10 = New DevExpress.XtraEditors.LabelControl()
            Me.testComb = New DevExpress.XtraEditors.TextEdit()
            Me.LabelControl9 = New DevExpress.XtraEditors.LabelControl()
            Me.btnNextIteration = New DevExpress.XtraEditors.SimpleButton()
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
            Me.FileSystemWatcher1 = New System.IO.FileSystemWatcher()
            Me.sccReference = New DevExpress.XtraEditors.SplitContainerControl()
            Me.sccMain = New DevExpress.XtraEditors.SplitContainerControl()
            Me.pgcUnitTesting = New System.Windows.Forms.PropertyGrid()
            CType(Me.GridView2, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.GridControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.GridView1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.GridView3, System.ComponentModel.ISupportInitialize).BeginInit()
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
            CType(Me.sccNetwork, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.sccNetwork.Panel1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sccNetwork.Panel1.SuspendLayout()
            CType(Me.sccNetwork.Panel2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sccNetwork.Panel2.SuspendLayout()
            Me.sccNetwork.SuspendLayout()
            CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.PanelControl1.SuspendLayout()
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
            CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.sccReference, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.sccReference.Panel1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sccReference.Panel1.SuspendLayout()
            CType(Me.sccReference.Panel2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sccReference.Panel2.SuspendLayout()
            Me.sccReference.SuspendLayout()
            CType(Me.sccMain, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.sccMain.Panel1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sccMain.Panel1.SuspendLayout()
            CType(Me.sccMain.Panel2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.sccMain.Panel2.SuspendLayout()
            Me.sccMain.SuspendLayout()
            Me.SuspendLayout()
            '
            'GridView2
            '
            Me.GridView2.GridControl = Me.GridControl1
            Me.GridView2.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always
            Me.GridView2.Name = "GridView2"
            Me.GridView2.OptionsFilter.AllowAutoFilterConditionChange = DevExpress.Utils.DefaultBoolean.[False]
            Me.GridView2.OptionsFilter.AllowFilterEditor = False
            Me.GridView2.OptionsFilter.InHeaderSearchMode = DevExpress.XtraGrid.Views.Grid.GridInHeaderSearchMode.Disabled
            Me.GridView2.OptionsView.BestFitMode = DevExpress.XtraGrid.Views.Grid.GridBestFitMode.Full
            Me.GridView2.OptionsView.ColumnAutoWidth = False
            Me.GridView2.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.[True]
            Me.GridView2.OptionsView.ShowGroupPanel = False
            Me.GridView2.OptionsView.ShowIndicator = False
            '
            'GridControl1
            '
            Me.GridControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.GridControl1.Location = New System.Drawing.Point(0, 0)
            Me.GridControl1.MainView = Me.GridView1
            Me.GridControl1.Name = "GridControl1"
            Me.GridControl1.Size = New System.Drawing.Size(88, 206)
            Me.GridControl1.TabIndex = 19
            Me.GridControl1.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView1, Me.GridView3, Me.GridView2})
            '
            'GridView1
            '
            Me.GridView1.GridControl = Me.GridControl1
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
            'GridView3
            '
            Me.GridView3.GridControl = Me.GridControl1
            Me.GridView3.Name = "GridView3"
            '
            'pgStructure
            '
            Me.pgStructure.Controls.Add(Me.SplitContainer3)
            Me.pgStructure.Controls.Add(Me.Panel2)
            Me.pgStructure.Location = New System.Drawing.Point(4, 22)
            Me.pgStructure.Name = "pgStructure"
            Me.pgStructure.Padding = New System.Windows.Forms.Padding(3)
            Me.pgStructure.Size = New System.Drawing.Size(1225, 742)
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
            Me.SplitContainer3.Size = New System.Drawing.Size(1219, 659)
            Me.SplitContainer3.SplitterDistance = 606
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
            Me.SplitContainer4.Panel1.Controls.Add(Me.btnConduct)
            Me.SplitContainer4.Panel1.Controls.Add(Me.btnSaveFndToEDS)
            '
            'SplitContainer4.Panel2
            '
            Me.SplitContainer4.Panel2.Controls.Add(Me.propgridFndXL)
            Me.SplitContainer4.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.SplitContainer4.Size = New System.Drawing.Size(606, 659)
            Me.SplitContainer4.SplitterDistance = 164
            Me.SplitContainer4.TabIndex = 4
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
            Me.propgridFndXL.Size = New System.Drawing.Size(438, 659)
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
            Me.SplitContainer5.Size = New System.Drawing.Size(609, 659)
            Me.SplitContainer5.SplitterDistance = 384
            Me.SplitContainer5.TabIndex = 0
            '
            'propgridFndEDS
            '
            Me.propgridFndEDS.Dock = System.Windows.Forms.DockStyle.Fill
            Me.propgridFndEDS.Location = New System.Drawing.Point(0, 0)
            Me.propgridFndEDS.Name = "propgridFndEDS"
            Me.propgridFndEDS.Size = New System.Drawing.Size(384, 659)
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
            Me.Panel2.Size = New System.Drawing.Size(1219, 77)
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
            Me.pgTNX.Size = New System.Drawing.Size(1225, 742)
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
            Me.SplitContainer1.Size = New System.Drawing.Size(1219, 691)
            Me.SplitContainer1.SplitterDistance = 606
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
            Me.scFromERI.Size = New System.Drawing.Size(606, 691)
            Me.scFromERI.SplitterDistance = 164
            Me.scFromERI.TabIndex = 4
            '
            'propgridTNXERI
            '
            Me.propgridTNXERI.Dock = System.Windows.Forms.DockStyle.Fill
            Me.propgridTNXERI.Location = New System.Drawing.Point(0, 0)
            Me.propgridTNXERI.Name = "propgridTNXERI"
            Me.propgridTNXERI.Size = New System.Drawing.Size(438, 691)
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
            Me.scFromEDS.Size = New System.Drawing.Size(609, 691)
            Me.scFromEDS.SplitterDistance = 384
            Me.scFromEDS.TabIndex = 0
            '
            'propgridTNXEDS
            '
            Me.propgridTNXEDS.Dock = System.Windows.Forms.DockStyle.Fill
            Me.propgridTNXEDS.Location = New System.Drawing.Point(0, 0)
            Me.propgridTNXEDS.Name = "propgridTNXEDS"
            Me.propgridTNXEDS.Size = New System.Drawing.Size(384, 691)
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
            Me.Panel1.Size = New System.Drawing.Size(1219, 45)
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
            Me.pgSQLBackUp.Size = New System.Drawing.Size(1225, 742)
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
            Me.TabControl1.Controls.Add(Me.pgUnitTesting)
            Me.TabControl1.Controls.Add(Me.pgSQLBackUp)
            Me.TabControl1.Controls.Add(Me.pgTNX)
            Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.TabControl1.Location = New System.Drawing.Point(0, 0)
            Me.TabControl1.Name = "TabControl1"
            Me.TabControl1.SelectedIndex = 0
            Me.TabControl1.Size = New System.Drawing.Size(1233, 768)
            Me.TabControl1.TabIndex = 0
            '
            'pgUnitTesting
            '
            Me.pgUnitTesting.Controls.Add(Me.sccMain)
            Me.pgUnitTesting.Controls.Add(Me.PanelControl1)
            Me.pgUnitTesting.Location = New System.Drawing.Point(4, 22)
            Me.pgUnitTesting.Name = "pgUnitTesting"
            Me.pgUnitTesting.Padding = New System.Windows.Forms.Padding(3)
            Me.pgUnitTesting.Size = New System.Drawing.Size(1225, 742)
            Me.pgUnitTesting.TabIndex = 5
            Me.pgUnitTesting.Text = "Testing"
            Me.pgUnitTesting.UseVisualStyleBackColor = True
            '
            'sccTesting
            '
            Me.sccTesting.Dock = System.Windows.Forms.DockStyle.Fill
            Me.sccTesting.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2
            Me.sccTesting.Horizontal = False
            Me.sccTesting.Location = New System.Drawing.Point(0, 0)
            Me.sccTesting.Name = "sccTesting"
            '
            'sccTesting.sccTesting_Panel1
            '
            Me.sccTesting.Panel1.Controls.Add(Me.SplitContainerControl1)
            Me.sccTesting.Panel1.Text = "Panel1"
            '
            'sccTesting.sccTesting_Panel2
            '
            Me.sccTesting.Panel2.Controls.Add(Me.sccReference)
            Me.sccTesting.Panel2.Text = "Panel2"
            Me.sccTesting.Size = New System.Drawing.Size(986, 616)
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
            Me.SplitContainerControl1.Panel1.Controls.Add(Me.sccNetwork)
            Me.SplitContainerControl1.Panel1.Text = "Panel1"
            '
            'SplitContainerControl1.Panel2
            '
            Me.SplitContainerControl1.Panel2.Controls.Add(Me.seLocal)
            Me.SplitContainerControl1.Panel2.Controls.Add(Me.LabelControl13)
            Me.SplitContainerControl1.Panel2.Text = "Panel2"
            Me.SplitContainerControl1.Size = New System.Drawing.Size(986, 400)
            Me.SplitContainerControl1.SplitterPosition = 622
            Me.SplitContainerControl1.TabIndex = 15
            '
            'sccNetwork
            '
            Me.sccNetwork.Dock = System.Windows.Forms.DockStyle.Fill
            Me.sccNetwork.Horizontal = False
            Me.sccNetwork.Location = New System.Drawing.Point(0, 0)
            Me.sccNetwork.Name = "sccNetwork"
            '
            'sccNetwork.sccNetwork_Panel1
            '
            Me.sccNetwork.Panel1.Controls.Add(Me.seNetwork)
            Me.sccNetwork.Panel1.Controls.Add(Me.LabelControl11)
            Me.sccNetwork.Panel1.Text = "Panel1"
            '
            'sccNetwork.sccNetwork_Panel2
            '
            Me.sccNetwork.Panel2.Controls.Add(Me.seSA)
            Me.sccNetwork.Panel2.Controls.Add(Me.LabelControl12)
            Me.sccNetwork.Panel2.Text = "Panel2"
            Me.sccNetwork.Size = New System.Drawing.Size(622, 400)
            Me.sccNetwork.SplitterPosition = 191
            Me.sccNetwork.TabIndex = 15
            '
            'seNetwork
            '
            Me.seNetwork.Dock = System.Windows.Forms.DockStyle.Fill
            Me.seNetwork.Location = New System.Drawing.Point(0, 16)
            Me.seNetwork.Margin = New System.Windows.Forms.Padding(0)
            Me.seNetwork.Name = "seNetwork"
            Me.seNetwork.SelectedFile = Nothing
            Me.seNetwork.Size = New System.Drawing.Size(622, 175)
            Me.seNetwork.TabIndex = 14
            '
            'LabelControl11
            '
            Me.LabelControl11.Appearance.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
            Me.LabelControl11.Appearance.Options.UseFont = True
            Me.LabelControl11.Dock = System.Windows.Forms.DockStyle.Top
            Me.LabelControl11.Location = New System.Drawing.Point(0, 0)
            Me.LabelControl11.Name = "LabelControl11"
            Me.LabelControl11.Size = New System.Drawing.Size(176, 16)
            Me.LabelControl11.TabIndex = 0
            Me.LabelControl11.Text = "Testing Directory: Network"
            '
            'seSA
            '
            Me.seSA.Dock = System.Windows.Forms.DockStyle.Fill
            Me.seSA.Location = New System.Drawing.Point(0, 16)
            Me.seSA.Margin = New System.Windows.Forms.Padding(0)
            Me.seSA.Name = "seSA"
            Me.seSA.SelectedFile = Nothing
            Me.seSA.Size = New System.Drawing.Size(622, 183)
            Me.seSA.TabIndex = 15
            '
            'LabelControl12
            '
            Me.LabelControl12.Appearance.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
            Me.LabelControl12.Appearance.Options.UseFont = True
            Me.LabelControl12.Dock = System.Windows.Forms.DockStyle.Top
            Me.LabelControl12.Location = New System.Drawing.Point(0, 0)
            Me.LabelControl12.Name = "LabelControl12"
            Me.LabelControl12.Size = New System.Drawing.Size(154, 16)
            Me.LabelControl12.TabIndex = 1
            Me.LabelControl12.Text = "SA Reference Directory"
            '
            'seLocal
            '
            Me.seLocal.Dock = System.Windows.Forms.DockStyle.Fill
            Me.seLocal.Location = New System.Drawing.Point(0, 16)
            Me.seLocal.Margin = New System.Windows.Forms.Padding(0)
            Me.seLocal.Name = "seLocal"
            Me.seLocal.SelectedFile = Nothing
            Me.seLocal.Size = New System.Drawing.Size(354, 384)
            Me.seLocal.TabIndex = 15
            '
            'LabelControl13
            '
            Me.LabelControl13.Appearance.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
            Me.LabelControl13.Appearance.Options.UseFont = True
            Me.LabelControl13.Dock = System.Windows.Forms.DockStyle.Top
            Me.LabelControl13.Location = New System.Drawing.Point(0, 0)
            Me.LabelControl13.Name = "LabelControl13"
            Me.LabelControl13.Size = New System.Drawing.Size(154, 16)
            Me.LabelControl13.TabIndex = 16
            Me.LabelControl13.Text = "Testing Directory: Local"
            '
            'rtbNotes
            '
            Me.rtbNotes.Dock = System.Windows.Forms.DockStyle.Fill
            Me.rtbNotes.Location = New System.Drawing.Point(0, 16)
            Me.rtbNotes.Name = "rtbNotes"
            Me.rtbNotes.Size = New System.Drawing.Size(888, 190)
            Me.rtbNotes.TabIndex = 17
            Me.rtbNotes.Text = ""
            '
            'LabelControl14
            '
            Me.LabelControl14.Appearance.Font = New System.Drawing.Font("Tahoma", 10.0!, System.Drawing.FontStyle.Bold)
            Me.LabelControl14.Appearance.Options.UseFont = True
            Me.LabelControl14.Dock = System.Windows.Forms.DockStyle.Top
            Me.LabelControl14.Location = New System.Drawing.Point(0, 0)
            Me.LabelControl14.Name = "LabelControl14"
            Me.LabelControl14.Size = New System.Drawing.Size(69, 16)
            Me.LabelControl14.TabIndex = 18
            Me.LabelControl14.Text = "Test Notes"
            '
            'PanelControl1
            '
            Me.PanelControl1.Controls.Add(Me.testPublishedResults)
            Me.PanelControl1.Controls.Add(Me.testPrevResults)
            Me.PanelControl1.Controls.Add(Me.testIterationResults)
            Me.PanelControl1.Controls.Add(Me.testConduct)
            Me.PanelControl1.Controls.Add(Me.chkWorkLocal)
            Me.PanelControl1.Controls.Add(Me.testLocalWorkarea)
            Me.PanelControl1.Controls.Add(Me.LabelControl10)
            Me.PanelControl1.Controls.Add(Me.testComb)
            Me.PanelControl1.Controls.Add(Me.LabelControl9)
            Me.PanelControl1.Controls.Add(Me.btnNextIteration)
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
            Me.PanelControl1.Size = New System.Drawing.Size(1219, 120)
            Me.PanelControl1.TabIndex = 16
            '
            'testPublishedResults
            '
            Me.testPublishedResults.Enabled = False
            Me.testPublishedResults.Location = New System.Drawing.Point(806, 35)
            Me.testPublishedResults.Name = "testPublishedResults"
            Me.testPublishedResults.Size = New System.Drawing.Size(125, 23)
            Me.testPublishedResults.TabIndex = 27
            Me.testPublishedResults.Text = "Published Tools Results"
            '
            'testPrevResults
            '
            Me.testPrevResults.Enabled = False
            Me.testPrevResults.Location = New System.Drawing.Point(806, 9)
            Me.testPrevResults.Name = "testPrevResults"
            Me.testPrevResults.Size = New System.Drawing.Size(125, 23)
            Me.testPrevResults.TabIndex = 26
            Me.testPrevResults.Text = "Previous SA Results"
            '
            'testIterationResults
            '
            Me.testIterationResults.Enabled = False
            Me.testIterationResults.Location = New System.Drawing.Point(700, 9)
            Me.testIterationResults.Name = "testIterationResults"
            Me.testIterationResults.Size = New System.Drawing.Size(75, 23)
            Me.testIterationResults.TabIndex = 25
            Me.testIterationResults.Text = "Save Results"
            '
            'testConduct
            '
            Me.testConduct.BackgroundImage = Global.Testing_Winform.My.Resources.Resources.download
            Me.testConduct.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
            Me.testConduct.Enabled = False
            Me.testConduct.ImageOptions.Image = Global.Testing_Winform.My.Resources.Resources.download
            Me.testConduct.ImageOptions.SvgImageColorizationMode = DevExpress.Utils.SvgImageColorizationMode.Full
            Me.testConduct.Location = New System.Drawing.Point(957, 9)
            Me.testConduct.Name = "testConduct"
            Me.testConduct.PaintStyle = DevExpress.XtraEditors.Controls.PaintStyles.Light
            Me.testConduct.Size = New System.Drawing.Size(77, 49)
            Me.testConduct.TabIndex = 24
            '
            'chkWorkLocal
            '
            Me.chkWorkLocal.Enabled = False
            Me.chkWorkLocal.Location = New System.Drawing.Point(856, 63)
            Me.chkWorkLocal.Name = "chkWorkLocal"
            Me.chkWorkLocal.Properties.Caption = "Work Local"
            Me.chkWorkLocal.Size = New System.Drawing.Size(75, 20)
            Me.chkWorkLocal.TabIndex = 23
            '
            'testLocalWorkarea
            '
            Me.testLocalWorkarea.Location = New System.Drawing.Point(93, 93)
            Me.testLocalWorkarea.Name = "testLocalWorkarea"
            Me.testLocalWorkarea.Size = New System.Drawing.Size(451, 20)
            Me.testLocalWorkarea.TabIndex = 22
            '
            'LabelControl10
            '
            Me.LabelControl10.Location = New System.Drawing.Point(9, 96)
            Me.LabelControl10.Name = "LabelControl10"
            Me.LabelControl10.Size = New System.Drawing.Size(78, 13)
            Me.LabelControl10.TabIndex = 21
            Me.LabelControl10.Text = "Local Work Area"
            '
            'testComb
            '
            Me.testComb.Location = New System.Drawing.Point(645, 63)
            Me.testComb.Name = "testComb"
            Me.testComb.Properties.ReadOnly = True
            Me.testComb.Size = New System.Drawing.Size(194, 20)
            Me.testComb.TabIndex = 20
            '
            'LabelControl9
            '
            Me.LabelControl9.Location = New System.Drawing.Point(556, 66)
            Me.LabelControl9.Name = "LabelControl9"
            Me.LabelControl9.Size = New System.Drawing.Size(83, 13)
            Me.LabelControl9.TabIndex = 19
            Me.LabelControl9.Text = "Test Combination"
            '
            'btnNextIteration
            '
            Me.btnNextIteration.Enabled = False
            Me.btnNextIteration.Location = New System.Drawing.Point(700, 35)
            Me.btnNextIteration.Name = "btnNextIteration"
            Me.btnNextIteration.Size = New System.Drawing.Size(75, 23)
            Me.btnNextIteration.TabIndex = 18
            Me.btnNextIteration.Text = "Create"
            '
            'testNextIteration
            '
            Me.testNextIteration.Location = New System.Drawing.Point(644, 37)
            Me.testNextIteration.Name = "testNextIteration"
            Me.testNextIteration.Properties.ReadOnly = True
            Me.testNextIteration.Size = New System.Drawing.Size(50, 20)
            Me.testNextIteration.TabIndex = 17
            '
            'LabelControl8
            '
            Me.LabelControl8.Location = New System.Drawing.Point(570, 40)
            Me.LabelControl8.Name = "LabelControl8"
            Me.LabelControl8.Size = New System.Drawing.Size(68, 13)
            Me.LabelControl8.TabIndex = 16
            Me.LabelControl8.Text = "Next Iteration"
            '
            'testID
            '
            Me.testID.Location = New System.Drawing.Point(93, 15)
            Me.testID.Name = "testID"
            Me.testID.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
            Me.testID.Properties.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99", "100"})
            Me.testID.Properties.TextEditStyle = DevExpress.XtraEditors.Controls.TextEditStyles.DisableTextEditor
            Me.testID.Size = New System.Drawing.Size(89, 20)
            Me.testID.TabIndex = 15
            '
            'LabelControl7
            '
            Me.LabelControl7.Location = New System.Drawing.Point(52, 18)
            Me.LabelControl7.Name = "LabelControl7"
            Me.LabelControl7.Size = New System.Drawing.Size(35, 13)
            Me.LabelControl7.TabIndex = 14
            Me.LabelControl7.Text = "Test ID"
            '
            'LabelControl1
            '
            Me.LabelControl1.Location = New System.Drawing.Point(188, 18)
            Me.LabelControl1.Name = "LabelControl1"
            Me.LabelControl1.Size = New System.Drawing.Size(13, 13)
            Me.LabelControl1.TabIndex = 3
            Me.LabelControl1.Text = "BU"
            '
            'testSaFolder
            '
            Me.testSaFolder.EditValue = ""
            Me.testSaFolder.Location = New System.Drawing.Point(93, 41)
            Me.testSaFolder.Name = "testSaFolder"
            Me.testSaFolder.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton()})
            Me.testSaFolder.Size = New System.Drawing.Size(451, 20)
            Me.testSaFolder.TabIndex = 1
            '
            'testFolder
            '
            Me.testFolder.Location = New System.Drawing.Point(93, 67)
            Me.testFolder.Name = "testFolder"
            Me.testFolder.Properties.ReadOnly = True
            Me.testFolder.Size = New System.Drawing.Size(451, 20)
            Me.testFolder.TabIndex = 13
            '
            'testBu
            '
            Me.testBu.Location = New System.Drawing.Point(207, 15)
            Me.testBu.Name = "testBu"
            Me.testBu.Properties.ReadOnly = True
            Me.testBu.Size = New System.Drawing.Size(100, 20)
            Me.testBu.TabIndex = 2
            '
            'testIteration
            '
            Me.testIteration.Location = New System.Drawing.Point(644, 11)
            Me.testIteration.Name = "testIteration"
            Me.testIteration.Properties.ReadOnly = True
            Me.testIteration.Size = New System.Drawing.Size(50, 20)
            Me.testIteration.TabIndex = 12
            '
            'LabelControl2
            '
            Me.LabelControl2.Location = New System.Drawing.Point(313, 18)
            Me.LabelControl2.Name = "LabelControl2"
            Me.LabelControl2.Size = New System.Drawing.Size(59, 13)
            Me.LabelControl2.TabIndex = 4
            Me.LabelControl2.Text = "Structure ID"
            '
            'LabelControl6
            '
            Me.LabelControl6.Location = New System.Drawing.Point(556, 14)
            Me.LabelControl6.Name = "LabelControl6"
            Me.LabelControl6.Size = New System.Drawing.Size(82, 13)
            Me.LabelControl6.TabIndex = 11
            Me.LabelControl6.Text = "Current Iteration"
            '
            'testSid
            '
            Me.testSid.Location = New System.Drawing.Point(378, 15)
            Me.testSid.Name = "testSid"
            Me.testSid.Properties.ReadOnly = True
            Me.testSid.Size = New System.Drawing.Size(36, 20)
            Me.testSid.TabIndex = 5
            '
            'LabelControl5
            '
            Me.LabelControl5.Location = New System.Drawing.Point(33, 70)
            Me.LabelControl5.Name = "LabelControl5"
            Me.LabelControl5.Size = New System.Drawing.Size(54, 13)
            Me.LabelControl5.TabIndex = 10
            Me.LabelControl5.Text = "Test Folder"
            '
            'testWo
            '
            Me.testWo.Location = New System.Drawing.Point(420, 11)
            Me.testWo.Name = "testWo"
            Me.testWo.Properties.ReadOnly = True
            Me.testWo.Size = New System.Drawing.Size(100, 20)
            Me.testWo.TabIndex = 6
            '
            'LabelControl4
            '
            Me.LabelControl4.Location = New System.Drawing.Point(41, 44)
            Me.LabelControl4.Name = "LabelControl4"
            Me.LabelControl4.Size = New System.Drawing.Size(46, 13)
            Me.LabelControl4.TabIndex = 8
            Me.LabelControl4.Text = "SA Folder"
            '
            'LabelControl3
            '
            Me.LabelControl3.Location = New System.Drawing.Point(396, 14)
            Me.LabelControl3.Name = "LabelControl3"
            Me.LabelControl3.Size = New System.Drawing.Size(18, 13)
            Me.LabelControl3.TabIndex = 7
            Me.LabelControl3.Text = "WO"
            '
            'FileSystemWatcher1
            '
            Me.FileSystemWatcher1.EnableRaisingEvents = True
            Me.FileSystemWatcher1.SynchronizingObject = Me
            '
            'sccReference
            '
            Me.sccReference.Dock = System.Windows.Forms.DockStyle.Fill
            Me.sccReference.Location = New System.Drawing.Point(0, 0)
            Me.sccReference.Name = "sccReference"
            '
            'sccReference.sccReference_Panel1
            '
            Me.sccReference.Panel1.Controls.Add(Me.rtbNotes)
            Me.sccReference.Panel1.Controls.Add(Me.LabelControl14)
            Me.sccReference.Panel1.Text = "Panel1"
            '
            'sccReference.sccReference_Panel2
            '
            Me.sccReference.Panel2.Controls.Add(Me.GridControl1)
            Me.sccReference.Panel2.Text = "Panel2"
            Me.sccReference.Size = New System.Drawing.Size(986, 206)
            Me.sccReference.SplitterPosition = 888
            Me.sccReference.TabIndex = 20
            '
            'sccMain
            '
            Me.sccMain.Dock = System.Windows.Forms.DockStyle.Fill
            Me.sccMain.FixedPanel = DevExpress.XtraEditors.SplitFixedPanel.Panel2
            Me.sccMain.Location = New System.Drawing.Point(3, 123)
            Me.sccMain.Name = "sccMain"
            '
            'sccMain.sccMain_Panel1
            '
            Me.sccMain.Panel1.Controls.Add(Me.sccTesting)
            Me.sccMain.Panel1.Text = "Panel1"
            '
            'sccMain.sccMain_Panel2
            '
            Me.sccMain.Panel2.Controls.Add(Me.pgcUnitTesting)
            Me.sccMain.Panel2.Text = "Panel2"
            Me.sccMain.Size = New System.Drawing.Size(1219, 616)
            Me.sccMain.SplitterPosition = 223
            Me.sccMain.TabIndex = 21
            '
            'pgcUnitTesting
            '
            Me.pgcUnitTesting.Dock = System.Windows.Forms.DockStyle.Fill
            Me.pgcUnitTesting.Location = New System.Drawing.Point(0, 0)
            Me.pgcUnitTesting.Name = "pgcUnitTesting"
            Me.pgcUnitTesting.Size = New System.Drawing.Size(223, 616)
            Me.pgcUnitTesting.TabIndex = 0
            '
            'frmMain
            '
            Me.Appearance.BackColor = System.Drawing.Color.White
            Me.Appearance.Options.UseBackColor = True
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.AutoScroll = True
            Me.ClientSize = New System.Drawing.Size(1233, 768)
            Me.Controls.Add(Me.TabControl1)
            Me.IconOptions.Image = CType(resources.GetObject("frmMain.IconOptions.Image"), System.Drawing.Image)
            Me.Name = "frmMain"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.Text = "EDS & Excel Testing"
            Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
            CType(Me.GridView2, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.GridControl1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.GridView1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.GridView3, System.ComponentModel.ISupportInitialize).EndInit()
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
            Me.SplitContainerControl1.Panel2.PerformLayout()
            CType(Me.SplitContainerControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.SplitContainerControl1.ResumeLayout(False)
            CType(Me.sccNetwork.Panel1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccNetwork.Panel1.ResumeLayout(False)
            Me.sccNetwork.Panel1.PerformLayout()
            CType(Me.sccNetwork.Panel2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccNetwork.Panel2.ResumeLayout(False)
            Me.sccNetwork.Panel2.PerformLayout()
            CType(Me.sccNetwork, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccNetwork.ResumeLayout(False)
            CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.PanelControl1.ResumeLayout(False)
            Me.PanelControl1.PerformLayout()
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
            CType(Me.FileSystemWatcher1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.sccReference.Panel1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccReference.Panel1.ResumeLayout(False)
            Me.sccReference.Panel1.PerformLayout()
            CType(Me.sccReference.Panel2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccReference.Panel2.ResumeLayout(False)
            CType(Me.sccReference, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccReference.ResumeLayout(False)
            CType(Me.sccMain.Panel1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccMain.Panel1.ResumeLayout(False)
            CType(Me.sccMain.Panel2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccMain.Panel2.ResumeLayout(False)
            CType(Me.sccMain, System.ComponentModel.ISupportInitialize).EndInit()
            Me.sccMain.ResumeLayout(False)
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
        Friend WithEvents seNetwork As SimpleExplorer
        Friend WithEvents seLocal As SimpleExplorer
        Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
        Friend WithEvents testID As DevExpress.XtraEditors.ComboBoxEdit
        Friend WithEvents LabelControl7 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents btnNextIteration As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents testNextIteration As DevExpress.XtraEditors.TextEdit
        Friend WithEvents LabelControl8 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents testComb As DevExpress.XtraEditors.TextEdit
        Friend WithEvents LabelControl9 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents rtbNotes As RichTextBox
        Friend WithEvents sccTesting As DevExpress.XtraEditors.SplitContainerControl
        Friend WithEvents testLocalWorkarea As DevExpress.XtraEditors.TextEdit
        Friend WithEvents LabelControl10 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents sccNetwork As DevExpress.XtraEditors.SplitContainerControl
        Friend WithEvents LabelControl11 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents seSA As SimpleExplorer
        Friend WithEvents LabelControl12 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents LabelControl13 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents LabelControl14 As DevExpress.XtraEditors.LabelControl
        Friend WithEvents chkWorkLocal As DevExpress.XtraEditors.CheckEdit
        Friend WithEvents FileSystemWatcher1 As IO.FileSystemWatcher
        Friend WithEvents testConduct As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents testPrevResults As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents testIterationResults As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents GridControl1 As DevExpress.XtraGrid.GridControl
        Friend WithEvents GridView1 As DevExpress.XtraGrid.Views.Grid.GridView
        Friend WithEvents GridView2 As DevExpress.XtraGrid.Views.Grid.GridView
        Friend WithEvents GridView3 As DevExpress.XtraGrid.Views.Grid.GridView
        Friend WithEvents Panel2 As Panel
        Friend WithEvents Panel1 As Panel
        Friend WithEvents testPublishedResults As DevExpress.XtraEditors.SimpleButton
        Friend WithEvents sccReference As DevExpress.XtraEditors.SplitContainerControl
        Friend WithEvents sccMain As DevExpress.XtraEditors.SplitContainerControl
        Friend WithEvents pgcUnitTesting As PropertyGrid

#End Region

    End Class
End Namespace