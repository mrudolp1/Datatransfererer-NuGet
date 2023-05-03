'Namespace DevExpress.XtraDialogs.Demos
Namespace UnitTesting
    Partial Class LogViewer

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
            Me.fileExplorerAssistant = New DevExpress.XtraDialogs.FileExplorerAssistant(Me.components)
            Me.gcTestLog = New DevExpress.XtraGrid.GridControl()
            Me.GridView2 = New DevExpress.XtraGrid.Views.Grid.GridView()
            Me.pnCheckButtons = New DevExpress.XtraEditors.PanelControl()
            Me.booEvent = New DevExpress.XtraEditors.CheckButton()
            Me.booDebug = New DevExpress.XtraEditors.CheckButton()
            Me.booError = New DevExpress.XtraEditors.CheckButton()
            Me.booWarning = New DevExpress.XtraEditors.CheckButton()
            Me.booInfo = New DevExpress.XtraEditors.CheckButton()
            CType(Me.fileExplorerAssistant, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.gcTestLog, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.GridView2, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.pnCheckButtons, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.pnCheckButtons.SuspendLayout()
            Me.SuspendLayout()
            '
            'gcTestLog
            '
            Me.gcTestLog.Dock = System.Windows.Forms.DockStyle.Fill
            Me.gcTestLog.Location = New System.Drawing.Point(0, 26)
            Me.gcTestLog.MainView = Me.GridView2
            Me.gcTestLog.Name = "gcTestLog"
            Me.gcTestLog.Size = New System.Drawing.Size(626, 371)
            Me.gcTestLog.TabIndex = 25
            Me.gcTestLog.ViewCollection.AddRange(New DevExpress.XtraGrid.Views.Base.BaseView() {Me.GridView2})
            '
            'GridView2
            '
            Me.GridView2.GridControl = Me.gcTestLog
            Me.GridView2.HorzScrollVisibility = DevExpress.XtraGrid.Views.Base.ScrollVisibility.Always
            Me.GridView2.Name = "GridView2"
            Me.GridView2.OptionsBehavior.AllowAddRows = DevExpress.Utils.DefaultBoolean.[False]
            Me.GridView2.OptionsBehavior.AllowDeleteRows = DevExpress.Utils.DefaultBoolean.[False]
            Me.GridView2.OptionsBehavior.ReadOnly = True
            Me.GridView2.OptionsCustomization.AllowColumnMoving = False
            Me.GridView2.OptionsFilter.AllowAutoFilterConditionChange = DevExpress.Utils.DefaultBoolean.[False]
            Me.GridView2.OptionsFilter.AllowFilterEditor = False
            Me.GridView2.OptionsFilter.InHeaderSearchMode = DevExpress.XtraGrid.Views.Grid.GridInHeaderSearchMode.Disabled
            Me.GridView2.OptionsLayout.Columns.AddNewColumns = False
            Me.GridView2.OptionsLayout.Columns.RemoveOldColumns = False
            Me.GridView2.OptionsView.BestFitMode = DevExpress.XtraGrid.Views.Grid.GridBestFitMode.Full
            Me.GridView2.OptionsView.ColumnAutoWidth = False
            Me.GridView2.OptionsView.ColumnHeaderAutoHeight = DevExpress.Utils.DefaultBoolean.[True]
            Me.GridView2.OptionsView.ShowGroupPanel = False
            '
            'pnCheckButtons
            '
            Me.pnCheckButtons.Controls.Add(Me.booEvent)
            Me.pnCheckButtons.Controls.Add(Me.booDebug)
            Me.pnCheckButtons.Controls.Add(Me.booError)
            Me.pnCheckButtons.Controls.Add(Me.booWarning)
            Me.pnCheckButtons.Controls.Add(Me.booInfo)
            Me.pnCheckButtons.Dock = System.Windows.Forms.DockStyle.Top
            Me.pnCheckButtons.Location = New System.Drawing.Point(0, 0)
            Me.pnCheckButtons.Name = "pnCheckButtons"
            Me.pnCheckButtons.Size = New System.Drawing.Size(626, 26)
            Me.pnCheckButtons.TabIndex = 24
            '
            'booEvent
            '
            Me.booEvent.Appearance.Options.UseTextOptions = True
            Me.booEvent.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.booEvent.Checked = True
            Me.booEvent.Dock = System.Windows.Forms.DockStyle.Left
            Me.booEvent.ImageOptions.SvgImage = Global.Testing_Winform.My.Resources.Resources.rightangleconnector
            Me.booEvent.ImageOptions.SvgImageSize = New System.Drawing.Size(15, 15)
            Me.booEvent.Location = New System.Drawing.Point(482, 2)
            Me.booEvent.Name = "booEvent"
            Me.booEvent.Size = New System.Drawing.Size(120, 22)
            Me.booEvent.TabIndex = 1
            Me.booEvent.Text = "EVENT"
            '
            'booDebug
            '
            Me.booDebug.Appearance.Options.UseTextOptions = True
            Me.booDebug.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.booDebug.Checked = True
            Me.booDebug.Dock = System.Windows.Forms.DockStyle.Left
            Me.booDebug.ImageOptions.SvgImage = Global.Testing_Winform.My.Resources.Resources.charttype_line
            Me.booDebug.ImageOptions.SvgImageSize = New System.Drawing.Size(15, 15)
            Me.booDebug.Location = New System.Drawing.Point(362, 2)
            Me.booDebug.Name = "booDebug"
            Me.booDebug.Size = New System.Drawing.Size(120, 22)
            Me.booDebug.TabIndex = 2
            Me.booDebug.Text = "DEBUG"
            '
            'booError
            '
            Me.booError.Appearance.Options.UseTextOptions = True
            Me.booError.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.booError.Checked = True
            Me.booError.Dock = System.Windows.Forms.DockStyle.Left
            Me.booError.ImageOptions.SvgImage = Global.Testing_Winform.My.Resources.Resources.highimportance
            Me.booError.ImageOptions.SvgImageSize = New System.Drawing.Size(15, 15)
            Me.booError.Location = New System.Drawing.Point(242, 2)
            Me.booError.Name = "booError"
            Me.booError.Size = New System.Drawing.Size(120, 22)
            Me.booError.TabIndex = 4
            Me.booError.Text = "ERROR"
            '
            'booWarning
            '
            Me.booWarning.Appearance.Options.UseTextOptions = True
            Me.booWarning.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.booWarning.Checked = True
            Me.booWarning.Dock = System.Windows.Forms.DockStyle.Left
            Me.booWarning.ImageOptions.SvgImage = Global.Testing_Winform.My.Resources.Resources.warning
            Me.booWarning.ImageOptions.SvgImageSize = New System.Drawing.Size(15, 15)
            Me.booWarning.Location = New System.Drawing.Point(122, 2)
            Me.booWarning.Name = "booWarning"
            Me.booWarning.Size = New System.Drawing.Size(120, 22)
            Me.booWarning.TabIndex = 3
            Me.booWarning.Text = "WARNING"
            '
            'booInfo
            '
            Me.booInfo.Appearance.Options.UseTextOptions = True
            Me.booInfo.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Near
            Me.booInfo.Checked = True
            Me.booInfo.Dock = System.Windows.Forms.DockStyle.Left
            Me.booInfo.ImageOptions.SvgImage = Global.Testing_Winform.My.Resources.Resources.about
            Me.booInfo.ImageOptions.SvgImageSize = New System.Drawing.Size(15, 15)
            Me.booInfo.Location = New System.Drawing.Point(2, 2)
            Me.booInfo.Name = "booInfo"
            Me.booInfo.Size = New System.Drawing.Size(120, 22)
            Me.booInfo.TabIndex = 0
            Me.booInfo.Text = "INFO"
            '
            'LogViewer
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            Me.Controls.Add(Me.gcTestLog)
            Me.Controls.Add(Me.pnCheckButtons)
            Me.Margin = New System.Windows.Forms.Padding(0)
            Me.Name = "LogViewer"
            Me.Size = New System.Drawing.Size(626, 397)
            CType(Me.fileExplorerAssistant, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.gcTestLog, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.GridView2, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.pnCheckButtons, System.ComponentModel.ISupportInitialize).EndInit()
            Me.pnCheckButtons.ResumeLayout(False)
            Me.ResumeLayout(False)

        End Sub

#End Region
        Private fileExplorerAssistant As DevExpress.XtraDialogs.FileExplorerAssistant
        Friend WithEvents gcTestLog As DevExpress.XtraGrid.GridControl
        Friend WithEvents GridView2 As DevExpress.XtraGrid.Views.Grid.GridView
        Friend WithEvents pnCheckButtons As DevExpress.XtraEditors.PanelControl
        Friend WithEvents booEvent As DevExpress.XtraEditors.CheckButton
        Friend WithEvents booDebug As DevExpress.XtraEditors.CheckButton
        Friend WithEvents booError As DevExpress.XtraEditors.CheckButton
        Friend WithEvents booWarning As DevExpress.XtraEditors.CheckButton
        Friend WithEvents booInfo As DevExpress.XtraEditors.CheckButton
    End Class
End Namespace

