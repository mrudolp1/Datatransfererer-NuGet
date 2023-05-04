'Namespace DevExpress.XtraDialogs.Demos
Namespace UnitTesting
    Partial Class SimpleExplorer

        ''' <summary> 
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary> 
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (Me.components IsNot Nothing) Then
                Me.components.Dispose()
            End If

            MyBase.Dispose(disposing)
        End Sub

#Region "Component Designer generated code"
        ''' <summary> 
        ''' Required method for Designer support - do not modify 
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SimpleExplorer))
            Me.fileExplorerAssistant = New DevExpress.XtraDialogs.FileExplorerAssistant(Me.components)
            Me.leftPanel = New DevExpress.XtraEditors.SidePanel()
            Me.treeList = New DevExpress.XtraTreeList.TreeList()
            Me.gridControl = New DevExpress.XtraGrid.GridControl()
            Me.gridView1 = New DevExpress.XtraGrid.Views.Grid.GridView()
            Me.toolbarPanel = New DevExpress.Utils.Layout.TablePanel()
            Me.btnBack = New DevExpress.XtraEditors.SimpleButton()
            Me.currentPathEdit = New DevExpress.XtraEditors.BreadCrumbEdit()
            Me.btnForward = New DevExpress.XtraEditors.SimpleButton()
            Me.btnUp = New DevExpress.XtraEditors.SimpleButton()
            Me.searchBox = New DevExpress.XtraEditors.SearchControl()
            Me.topPanel = New DevExpress.XtraEditors.SidePanel()
            CType(Me.fileExplorerAssistant, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.leftPanel.SuspendLayout()
            CType(Me.treeList, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.gridControl, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.gridView1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.toolbarPanel, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.toolbarPanel.SuspendLayout()
            CType(Me.currentPathEdit.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.searchBox.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.topPanel.SuspendLayout()
            Me.SuspendLayout()
            '
            'leftPanel
            '
            Me.leftPanel.Controls.Add(Me.treeList)
            Me.leftPanel.Dock = System.Windows.Forms.DockStyle.Left
            Me.leftPanel.Location = New System.Drawing.Point(0, 48)
            Me.leftPanel.Margin = New System.Windows.Forms.Padding(0)
            Me.leftPanel.Name = "leftPanel"
            Me.leftPanel.Size = New System.Drawing.Size(17, 349)
            Me.leftPanel.TabIndex = 0
            Me.leftPanel.Visible = False
            '
            'treeList
            '
            Me.treeList.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
            Me.treeList.Dock = System.Windows.Forms.DockStyle.Fill
            Me.treeList.Location = New System.Drawing.Point(0, 0)
            Me.treeList.Margin = New System.Windows.Forms.Padding(0)
            Me.treeList.Name = "treeList"
            Me.treeList.Size = New System.Drawing.Size(16, 349)
            Me.treeList.TabIndex = 0
            '
            'gridControl
            '
            Me.gridControl.Dock = System.Windows.Forms.DockStyle.Fill
            Me.gridControl.Location = New System.Drawing.Point(17, 48)
            Me.gridControl.MainView = Me.gridView1
            Me.gridControl.Margin = New System.Windows.Forms.Padding(0)
            Me.gridControl.Name = "gridControl"
            Me.gridControl.Size = New System.Drawing.Size(609, 349)
            Me.gridControl.TabIndex = 1
            Me.gridControl.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.gridView1})
            '
            'gridView1
            '
            Me.gridView1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
            Me.gridView1.GridControl = Me.gridControl
            Me.gridView1.Name = "gridView1"
            Me.gridView1.OptionsCustomization.AllowGroup = False
            Me.gridView1.OptionsMenu.EnableColumnMenu = False
            Me.gridView1.OptionsMenu.EnableGroupPanelMenu = False
            Me.gridView1.OptionsSelection.MultiSelect = True
            Me.gridView1.OptionsView.ShowGroupPanel = False
            '
            'toolbarPanel
            '
            Me.toolbarPanel.AutoSize = True
            Me.toolbarPanel.Columns.AddRange(New DevExpress.Utils.Layout.TablePanelColumn() {New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 1.0!), New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 1.0!), New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 1.0!), New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Relative, 1.0!), New DevExpress.Utils.Layout.TablePanelColumn(DevExpress.Utils.Layout.TablePanelEntityStyle.Absolute, 0!)})
            Me.toolbarPanel.Controls.Add(Me.btnBack)
            Me.toolbarPanel.Controls.Add(Me.currentPathEdit)
            Me.toolbarPanel.Controls.Add(Me.btnForward)
            Me.toolbarPanel.Controls.Add(Me.btnUp)
            Me.toolbarPanel.Controls.Add(Me.searchBox)
            Me.toolbarPanel.Dock = System.Windows.Forms.DockStyle.Fill
            Me.toolbarPanel.Location = New System.Drawing.Point(12, 8)
            Me.toolbarPanel.Margin = New System.Windows.Forms.Padding(0)
            Me.toolbarPanel.Name = "toolbarPanel"
            Me.toolbarPanel.Rows.AddRange(New DevExpress.Utils.Layout.TablePanelRow() {New DevExpress.Utils.Layout.TablePanelRow(DevExpress.Utils.Layout.TablePanelEntityStyle.AutoSize, 1.0!)})
            Me.toolbarPanel.Size = New System.Drawing.Size(602, 31)
            Me.toolbarPanel.TabIndex = 2
            '
            'btnBack
            '
            Me.btnBack.AutoSize = True
            Me.toolbarPanel.SetColumn(Me.btnBack, 0)
            Me.btnBack.ImageOptions.SvgImage = CType(resources.GetObject("btnBack.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
            Me.btnBack.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
            Me.btnBack.Location = New System.Drawing.Point(0, 1)
            Me.btnBack.Margin = New System.Windows.Forms.Padding(0)
            Me.btnBack.Name = "btnBack"
            Me.btnBack.PaintStyle = DevExpress.XtraEditors.Controls.PaintStyles.Light
            Me.toolbarPanel.SetRow(Me.btnBack, 0)
            Me.btnBack.ShowFocusRectangle = DevExpress.Utils.DefaultBoolean.[False]
            Me.btnBack.Size = New System.Drawing.Size(30, 28)
            Me.btnBack.TabIndex = 0
            '
            'currentPathEdit
            '
            Me.toolbarPanel.SetColumn(Me.currentPathEdit, 3)
            Me.currentPathEdit.Location = New System.Drawing.Point(92, 5)
            Me.currentPathEdit.Margin = New System.Windows.Forms.Padding(2, 0, 0, 0)
            Me.currentPathEdit.Name = "currentPathEdit"
            Me.currentPathEdit.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
            Me.toolbarPanel.SetRow(Me.currentPathEdit, 0)
            Me.currentPathEdit.Size = New System.Drawing.Size(510, 20)
            Me.currentPathEdit.TabIndex = 4
            '
            'btnForward
            '
            Me.btnForward.AutoSize = True
            Me.toolbarPanel.SetColumn(Me.btnForward, 1)
            Me.btnForward.ImageOptions.SvgImage = CType(resources.GetObject("btnForward.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
            Me.btnForward.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
            Me.btnForward.Location = New System.Drawing.Point(30, 1)
            Me.btnForward.Margin = New System.Windows.Forms.Padding(0)
            Me.btnForward.Name = "btnForward"
            Me.btnForward.PaintStyle = DevExpress.XtraEditors.Controls.PaintStyles.Light
            Me.toolbarPanel.SetRow(Me.btnForward, 0)
            Me.btnForward.ShowFocusRectangle = DevExpress.Utils.DefaultBoolean.[False]
            Me.btnForward.Size = New System.Drawing.Size(30, 28)
            Me.btnForward.TabIndex = 2
            '
            'btnUp
            '
            Me.btnUp.AutoSize = True
            Me.toolbarPanel.SetColumn(Me.btnUp, 2)
            Me.btnUp.ImageOptions.SvgImage = CType(resources.GetObject("btnUp.ImageOptions.SvgImage"), DevExpress.Utils.Svg.SvgImage)
            Me.btnUp.ImageOptions.SvgImageSize = New System.Drawing.Size(24, 24)
            Me.btnUp.Location = New System.Drawing.Point(60, 1)
            Me.btnUp.Margin = New System.Windows.Forms.Padding(0)
            Me.btnUp.Name = "btnUp"
            Me.btnUp.PaintStyle = DevExpress.XtraEditors.Controls.PaintStyles.Light
            Me.toolbarPanel.SetRow(Me.btnUp, 0)
            Me.btnUp.ShowFocusRectangle = DevExpress.Utils.DefaultBoolean.[False]
            Me.btnUp.Size = New System.Drawing.Size(30, 28)
            Me.btnUp.TabIndex = 2
            '
            'searchBox
            '
            Me.searchBox.Client = Me.gridControl
            Me.toolbarPanel.SetColumn(Me.searchBox, 4)
            Me.searchBox.Location = New System.Drawing.Point(0, 0)
            Me.searchBox.Margin = New System.Windows.Forms.Padding(8, 0, 0, 0)
            Me.searchBox.Name = "searchBox"
            Me.searchBox.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Repository.ClearButton(), New DevExpress.XtraEditors.Repository.SearchButton()})
            Me.searchBox.Properties.Client = Me.gridControl
            Me.toolbarPanel.SetRow(Me.searchBox, 0)
            Me.searchBox.Size = New System.Drawing.Size(0, 20)
            Me.searchBox.TabIndex = 2
            Me.searchBox.Visible = False
            '
            'topPanel
            '
            Me.topPanel.AllowResize = False
            Me.topPanel.Controls.Add(Me.toolbarPanel)
            Me.topPanel.Dock = System.Windows.Forms.DockStyle.Top
            Me.topPanel.Location = New System.Drawing.Point(0, 0)
            Me.topPanel.Margin = New System.Windows.Forms.Padding(0)
            Me.topPanel.Name = "topPanel"
            Me.topPanel.Padding = New System.Windows.Forms.Padding(12, 8, 12, 8)
            Me.topPanel.Size = New System.Drawing.Size(626, 48)
            Me.topPanel.TabIndex = 5
            '
            'SimpleExplorer
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            Me.Controls.Add(Me.gridControl)
            Me.Controls.Add(Me.leftPanel)
            Me.Controls.Add(Me.topPanel)
            Me.Margin = New System.Windows.Forms.Padding(0)
            Me.Name = "SimpleExplorer"
            Me.Size = New System.Drawing.Size(626, 397)
            CType(Me.fileExplorerAssistant, System.ComponentModel.ISupportInitialize).EndInit()
            Me.leftPanel.ResumeLayout(False)
            CType(Me.treeList, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.gridControl, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.gridView1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.toolbarPanel, System.ComponentModel.ISupportInitialize).EndInit()
            Me.toolbarPanel.ResumeLayout(False)
            Me.toolbarPanel.PerformLayout()
            CType(Me.currentPathEdit.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.searchBox.Properties, System.ComponentModel.ISupportInitialize).EndInit()
            Me.topPanel.ResumeLayout(False)
            Me.topPanel.PerformLayout()
            Me.ResumeLayout(False)

        End Sub

#End Region
        Private fileExplorerAssistant As DevExpress.XtraDialogs.FileExplorerAssistant

        Private leftPanel As DevExpress.XtraEditors.SidePanel

        Private treeList As DevExpress.XtraTreeList.TreeList

        Private gridControl As DevExpress.XtraGrid.GridControl

        Private gridView1 As DevExpress.XtraGrid.Views.Grid.GridView

        Private toolbarPanel As DevExpress.Utils.Layout.TablePanel

        Private WithEvents btnBack As DevExpress.XtraEditors.SimpleButton

        Private currentPathEdit As DevExpress.XtraEditors.BreadCrumbEdit

        Private WithEvents btnForward As DevExpress.XtraEditors.SimpleButton

        Private WithEvents btnUp As DevExpress.XtraEditors.SimpleButton

        Private searchBox As DevExpress.XtraEditors.SearchControl

        Private topPanel As DevExpress.XtraEditors.SidePanel
    End Class
End Namespace

