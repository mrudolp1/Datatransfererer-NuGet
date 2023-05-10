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
                '</gridControl>

                '<gridview>
                AddHandler gridView1.RowClick, AddressOf gridView1_RowClick
                AddHandler gridView1.DoubleClick, AddressOf gridView1_DoubleClick
                '</gridview>

                '<listview>
                AddHandler listView.AllowGoBackChanged, AddressOf OnUpdateNavigationButtons
                AddHandler listView.AllowGoForwardChanged, AddressOf OnUpdateNavigationButtons
                AddHandler listView.AllowGoUpChanged, AddressOf OnUpdateNavigationButtons
                '</listview

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

                'Allow multiselect to permit users to select multiple folders or files at the same time. 
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
            HideDateAndSize()
        End Sub

        '</gridControl>
        '<treeList>
        Private Sub OnTreeCurrentItemChanged(ByVal sender As Object, ByVal e As CurrentItemChangedEventArgs)
            listView.SetCurrentItem(e.CurrentItem)
            HideDateAndSize()
        End Sub

        '</treeList>
        '<currentPathEdit>
        Private Sub OnCurrentPathEditCurrentItemChanged(ByVal sender As Object, ByVal e As CurrentItemChangedEventArgs)
            listView.SetCurrentItem(e.CurrentItem)
            HideDateAndSize()
        End Sub

        '</currentPathEdit>
        Private Sub btnBack_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnBack.Click
            listView.GoBack()
            HideDateAndSize()
        End Sub

        Private Sub btnForward_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnForward.Click
            listView.GoForward()
            HideDateAndSize()
        End Sub

        Private Sub btnUp_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnUp.Click
            listView.GoUp()
            HideDateAndSize()
        End Sub

        Private Sub gridView1_RowClick(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowClickEventArgs)
            'Only perform this action for specific files
            'It was taking too long to load tools and csvs from network drives while working at home
            If Me.Name.ToLower = "selocal" Then
                ButtonclickToggle(Me.Cursor, Cursors.WaitCursor)
                Dim info As IO.FileInfo
                Dim fName As String
                Dim path As String
                Dim loadDt As DataTable

                'Get the filename from the sepcified row click
                fName = gridView1.GetRowCellValue(gridView1.FocusedRowHandle, "Name")
                'Get the path from the bread crumb path (could change as the user navigates)
                path = breadCrumb.CurrentPath
                'Create a new fileinfo based on the path and file name
                info = New System.IO.FileInfo(path & "\" & fName)

                'If the info is something (i.e. not a file folder) then attempt to load results or csv data
                If info IsNot Nothing Then
                    If fName.Contains(".") And fName.Contains("xlsm") Then

                        SelectedFile = info
                        loadDt = frmMain.SummarizedResults(info)
                    ElseIf fName.Contains(".csv") Then
                        loadDt = CSVtoDatatable(info)
                    ElseIf fName.Contains(".txt") Or info.Extension.ToLower = ".eri" Then
                        Using sr As New IO.StreamReader(info.FullName)
                            Dim tempDt As New DataTable
                            tempDt.Columns.Add("Text")
                            Dim newRow As String()

                            While Not sr.EndOfStream
                                'newRow = sr.ReadLine.Split("|")
                                'If newRow.Count > 0 Then
                                '    If tempDt.Columns.Count = 0 Then
                                '        tempDt.Columns.Add("Time", GetType(System.String))
                                '        tempDt.Columns.Add("Message", GetType(System.String))
                                '    End If
                                '    Dim combined As String

                                '    Try
                                '        combined = newRow(1) & "|" & newRow(2)
                                '    Catch ex As Exception
                                '        Try
                                '            combined = newRow(1)
                                '        Catch ex1 As Exception
                                '            combined = ""
                                '        End Try
                                '    End Try

                                '    tempDt.Rows.Add(newRow(0), combined)
                                'Else
                                '    If tempDt.Columns.Count = 0 Then
                                '        tempDt.Columns.Add("Text", GetType(System.String))
                                '    End If
                                '    tempDt.Rows.Add(sr.ReadLine)
                                'End If
                                tempDt.Rows.Add(sr.ReadLine)
                            End While

                            loadDt = tempDt
                            sr.Close()
                        End Using
                    End If

                    'Set the reference grid on the main form to the returned datatable
                    frmMain.GridView1.Columns.Clear()
                    frmMain.gcViewer.DataSource = Nothing
                    frmMain.gcViewer.DataSource = loadDt
                    frmMain.gcViewer.RefreshDataSource()
                    frmMain.GridView1.BestFitColumns(True)

                    Try
                        If loadDt IsNot Nothing Then frmMain.LogActivity("DEBUG | Loaded file for viewing: " & info.FullName, True)
                    Catch ex As Exception
                    End Try
                End If
                ButtonclickToggle(Me.Cursor, Cursors.Default)
            Else
            End If

            HideDateAndSize()

        End Sub

        'When a row is double clicked it was displaying all columns again
        Private Sub gridView1_DoubleClick(sender As Object, e As EventArgs)
            HideDateAndSize()
        End Sub

        'Since Date and file size don't update consistently, hide these columns from the view to alleviate confusion
        'This is basically just constantly done to force these columns to b ehidden. 
        Private Sub HideDateAndSize()
            Try
                Me.gridView1.Columns(2).Visible = False
                Me.gridView1.Columns(4).Visible = False
            Catch
            End Try
        End Sub


    End Class
End Namespace