Imports System
Imports DevExpress.XtraEditors
Imports DevExpress.XtraDialogs.FileExplorerExtensions
Imports DevExpress.Data

'Namespace DevExpress.XtraDialogs.Demos

Partial Public Class SimpleExplorer
    Inherits XtraUserControl

    Private ReadOnly listView As GridControlExtension

    Private ReadOnly breadCrumb As BreadCrumbExtension

    Private ReadOnly folderTree As TreeListExtension



    Public Sub New()
        InitializeComponent()
        If Not fileExplorerAssistant.IsDesignMode Then
            '<gridControl>
            listView = fileExplorerAssistant.Attach(gridControl, Sub(x)
                                                                     'x.FilterString = "Text files (*.txt)|*.txt")
                                                                 End Sub)
            AddHandler listView.FocusedLinkChanged, AddressOf OnListViewFocusedLinkChanged
            AddHandler listView.CurrentItemChanged, AddressOf OnListViewCurrentItemChanged
            '</gridControl>
            AddHandler listView.AllowGoBackChanged, AddressOf OnUpdateNavigationButtons
            AddHandler listView.AllowGoForwardChanged, AddressOf OnUpdateNavigationButtons
            AddHandler listView.AllowGoUpChanged, AddressOf OnUpdateNavigationButtons
            '<currentPathEdit>
            breadCrumb = fileExplorerAssistant.Attach(currentPathEdit, Sub(x) AddHandler x.CurrentItemChanged, AddressOf OnCurrentPathEditCurrentItemChanged)
            '</currentPathEdit>
            '<treeList>
            folderTree = fileExplorerAssistant.Attach(treeList, Sub(x)
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
    End Sub

    '</gridControl>
    '<treeList>
    Private Sub OnTreeCurrentItemChanged(ByVal sender As Object, ByVal e As CurrentItemChangedEventArgs)
        listView.SetCurrentItem(e.CurrentItem)
    End Sub

    '</treeList>
    '<currentPathEdit>
    Private Sub OnCurrentPathEditCurrentItemChanged(ByVal sender As Object, ByVal e As CurrentItemChangedEventArgs)
        listView.SetCurrentItem(e.CurrentItem)
    End Sub

    '</currentPathEdit>
    Private Sub btnBack_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnBack.Click
        listView.GoBack()
    End Sub

    Private Sub btnForward_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnForward.Click
        listView.GoForward()
    End Sub

    Private Sub btnUp_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnUp.Click
        listView.GoUp()
    End Sub

End Class
'End Namespace
