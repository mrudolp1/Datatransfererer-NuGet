Imports System
Imports DevExpress.XtraEditors
Imports DevExpress.XtraDialogs.FileExplorerExtensions
Imports DevExpress.Data
Imports CCI_Engineering_Templates

Namespace UnitTesting
    Partial Public Class SimpleExplorer
        Inherits XtraUserControl

        Private ReadOnly listView As GridControlExtension

        Private ReadOnly breadCrumb As BreadCrumbExtension

        Private ReadOnly folderTree As TreeListExtension

        Public Property SelectedFile As IO.FileInfo = Nothing
        Public Property ResultsDT As New DataTable

        Public Sub New()
            InitializeComponent()
            If Not fileExplorerAssistant.IsDesignMode Then
                '<gridControl>
                listView = fileExplorerAssistant.Attach(gridControl, Sub(x)
                                                                         'x.FilterString = "Text files (*.txt)|*.txt")
                                                                     End Sub)
                AddHandler listView.FocusedLinkChanged, AddressOf OnListViewFocusedLinkChanged
                AddHandler listView.CurrentItemChanged, AddressOf OnListViewCurrentItemChanged
                AddHandler gridView1.RowClick, AddressOf gridView1_RowClick
                AddHandler gridView1.DoubleClick, AddressOf gridView1_DoubleClick
                '</gridControl>
                AddHandler listView.AllowGoBackChanged, AddressOf OnUpdateNavigationButtons
                AddHandler listView.AllowGoForwardChanged, AddressOf OnUpdateNavigationButtons
                AddHandler listView.AllowGoUpChanged, AddressOf OnUpdateNavigationButtons
                '<currentPathEdit>
                breadCrumb = fileExplorerAssistant.Attach(currentPathEdit, Sub(x) AddHandler x.CurrentItemChanged, AddressOf OnCurrentPathEditCurrentItemChanged)
                '</currentPathEdit>
                '<treeList>
                folderTree = fileExplorerAssistant.Attach(TreeList, Sub(x)
                                                                        x.RootNodes.Add(New EnvironmentSpecialFolderNode(Environment.SpecialFolder.Desktop))
                                                                    End Sub)
                AddHandler folderTree.CurrentItemChanged, AddressOf OnTreeCurrentItemChanged
                folderTree.AutoExpandToCurrent = True
                '</treeList>

                'listView.SetCurrentPath("C:\SAPI Work Area")
                'folderTree.SetCurrentPath("C:\SAPI Work Area")

                Me.listView.MultiSelect = True
            End If
        End Sub

        Public Sub SetCurrentDirectory(Path As String)
            folderTree.RootNodes.Clear()
            folderTree.RootNodes.Add(New PathNode(Path))
            folderTree.SetCurrentPath(Path)

            Try
                Me.gridView1.Columns(2).Visible = False
                Me.gridView1.Columns(4).Visible = False
            Catch
            End Try
        End Sub

        Public Sub SetRootandCurrentPath()
            'folderTree.RootNodes.Clear()
            'folderTree.RootNodes.Add(New PathNode(path))
        End Sub

        Private Sub OnUpdateNavigationButtons(ByVal sender As Object, ByVal e As EventArgs)
            btnUp.Enabled = listView.CanGoUp
            btnBack.Enabled = listView.CanGoBack
            btnForward.Enabled = listView.CanGoForward
        End Sub

        '<gridControl>
        Private Sub OnListViewFocusedLinkChanged(ByVal sender As Object, ByVal e As FocusedLinkChangedEventArgs)

        End Sub

        Private Sub OnListViewCurrentItemChanged(ByVal sender As Object, ByVal e As CurrentItemChangedEventArgs)
            breadCrumb.SetCurrentItem(e.CurrentItem)
            Try
                Me.gridView1.Columns(2).Visible = False
                Me.gridView1.Columns(4).Visible = False
            Catch
            End Try
        End Sub

        '</gridControl>
        '<treeList>
        Private Sub OnTreeCurrentItemChanged(ByVal sender As Object, ByVal e As CurrentItemChangedEventArgs)
            listView.SetCurrentItem(e.CurrentItem)
            Try
                Me.gridView1.Columns(2).Visible = False
                Me.gridView1.Columns(4).Visible = False
            Catch
            End Try
        End Sub

        '</treeList>
        '<currentPathEdit>
        Private Sub OnCurrentPathEditCurrentItemChanged(ByVal sender As Object, ByVal e As CurrentItemChangedEventArgs)
            listView.SetCurrentItem(e.CurrentItem)
            Try
                Me.gridView1.Columns(2).Visible = False
                Me.gridView1.Columns(4).Visible = False
            Catch
            End Try
        End Sub

        '</currentPathEdit>
        Private Sub btnBack_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnBack.Click
            listView.GoBack()
            Try
                Me.gridView1.Columns(2).Visible = False
                Me.gridView1.Columns(4).Visible = False
            Catch
            End Try
        End Sub

        Private Sub btnForward_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnForward.Click
            listView.GoForward()
            Try
                Me.gridView1.Columns(2).Visible = False
                Me.gridView1.Columns(4).Visible = False
            Catch
            End Try
        End Sub

        Private Sub btnUp_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnUp.Click
            listView.GoUp()
            Try
                Me.gridView1.Columns(2).Visible = False
                Me.gridView1.Columns(4).Visible = False
            Catch
            End Try
        End Sub

        Private Sub gridView1_RowClick(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs)
            If Me.Name.ToLower = "selocal" Then
                Dim info As IO.FileInfo
                Dim fName As String
                Dim path As String
                Dim finalDT As DataTable
                fName = gridView1.GetRowCellValue(gridView1.FocusedRowHandle, "Name")
                path = breadCrumb.CurrentPath
                info = New System.IO.FileInfo(path & "\" & fName)
                If info IsNot Nothing Then
                    If fName.Contains(".") And fName.Contains("xlsm") Then

                        SelectedFile = info
                        finalDT = SummarizedResults(info)

                    ElseIf fName.Contains(".csv") Then
                        finalDT = CSVtoDatatable(info)
                    End If

                    frmMain.GridView1.Columns.Clear()
                    frmMain.GridControl1.DataSource = Nothing
                    frmMain.GridControl1.DataSource = finalDT
                    frmMain.GridControl1.RefreshDataSource()
                    frmMain.GridView1.BestFitColumns(True)
                End If
            Else
            End If

            Try
                Me.gridView1.Columns(2).Visible = False
                Me.gridView1.Columns(4).Visible = False
            Catch
            End Try
        End Sub

        Private Sub gridView1_DoubleClick(sender As Object, e As EventArgs)
            Try
                Me.gridView1.Columns(2).Visible = False
                Me.gridView1.Columns(4).Visible = False
            Catch
            End Try
        End Sub


    End Class
End Namespace
