Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions

Public Module UnitTestingExtensions
    <Extension>
    Public Function ResultsToDataTable(ByVal myTnx As tnxModel) As DataTable
        Dim tnxDt As New DataTable
        tnxDt.Columns.Add("Type", GetType(System.String))
        tnxDt.Columns.Add("Rating", GetType(System.String))
        tnxDt.Columns.Add("Tool", GetType(System.String))
        tnxDt.TableName = New FileInfo(myTnx.filePath).Name

        For Each up In myTnx.geometry.upperStructure
            For Each res In up.Results
                If res.rating > 0 Then tnxDt.Rows.Add(res.result_lkup, res.rating, "TNX Upper Section " & up.Rec)
            Next
        Next

        For Each down In myTnx.geometry.baseStructure
            For Each res In down.Results
                If res.rating > 0 Then tnxDt.Rows.Add(res.result_lkup, res.rating, "TNX Base Section " & down.Rec)
            Next
        Next

        Return tnxDt
    End Function

    <Extension()>
    Public Sub AddERIs(ByVal myList As List(Of String), ByVal myDir As String, Optional ByVal purgeERI As Boolean = True)
        For Each info As FileInfo In New DirectoryInfo(myDir).GetFiles
            If info.Extension.ToLower() = ".eri" Then
                'All eris permitted
                myList.Add(info.FullName)
                'LogActivity("DEBUG | ERI: " & info.Name & " found")
            ElseIf info.Name.ToLower.Contains(".eri.") Or info.Extension.ToLower = ".tfnx" Then
                If purgeERI Then
                    info.Delete()
                    'LogActivity("DEBUG | File Deleted: " & info.Name & "")
                End If
            End If
        Next
    End Sub

    'Custom extension to sort the results datatables by check/failure mode and tool name
    '''Extension specific to datatables
    '''Adds a reference column for results comparison
    '''Names the table based on the optional parameter provided
    '''Sorts the datatable by Type and Tool
    <Extension()>
    Public Sub ResultsSorting(ByRef dt As DataTable, Optional ByVal addColumn As String = Nothing)
        If addColumn IsNot Nothing Then
            Dim newcolumn As New Data.DataColumn("Summary Type", GetType(System.String))
            newcolumn.DefaultValue = addColumn
            dt.Columns.Add(newcolumn)
        End If

        dt.Columns(1).ColumnName = "Rating Old"

        Dim newRatingColumn As New Data.DataColumn("Rating", GetType(System.String))
        dt.Columns.Add(newRatingColumn)

        For Each dr As DataRow In dt.Rows
            dr.Item("Rating") = dr.Item("Rating Old").ToString
        Next

        dt.Columns.Remove("Rating Old")
        If addColumn IsNot Nothing Then dt.TableName = addColumn
        dt.AsDataView.Sort = "Type ASC, Tool ASC"
    End Sub

    'Determine if 2 datatables have the same exact values 
    '''Returns a boolean determining if they are the sam
    '''Returns a databale of the compared values
    <Extension()>
    Public Function IsMatching(ByRef dt As DataTable, ByVal comparer As DataTable) As Tuple(Of Boolean, DataTable)
        Dim dtVal As Double = Double.NaN
        Dim comparerVal As Double = Double.NaN
        Dim delta As Double = Double.NaN
        Dim perDelta As Double = Double.NaN

        Dim matching As Boolean = True 'Item1
        Dim diffDt As New DataTable 'Item2
        diffDt.TableName = dt.TableName & " v. " & comparer.TableName
        diffDt.Columns.Add(dt.TableName & " File", GetType(System.String))
        diffDt.Columns.Add("Check/Failure Mode", GetType(System.String))
        diffDt.Columns.Add(dt.TableName & " Val", GetType(System.String))
        diffDt.Columns.Add(comparer.TableName & " Val", GetType(System.String))
        diffDt.Columns.Add("Delta", GetType(System.String))
        diffDt.Columns.Add("% Difference", GetType(System.String))
        diffDt.Columns.Add("Status", GetType(System.String))
        For i As Integer = 0 To Math.Max(dt.Rows.Count, comparer.Rows.Count) - 1
            Dim dtRow As DataRow
            Dim comparerRow As DataRow

            Try
                dtRow = dt.Rows(i)
            Catch ex As Exception
                dtRow = Nothing
            End Try

            If dtRow IsNot Nothing Then
                For Each dr As DataRow In comparer.Rows
                    If dr.Item("Type").ToString = dtRow.Item("Type").ToString Then
                        If dr.Item("Tool").ToString.Contains(dtRow.Item("Tool").ToString) Or dtRow.Item("Tool").ToString.Contains(dr.Item("Tool").ToString) Then
                            comparerRow = dr
                            Exit For
                        Else
                            comparerRow = Nothing
                        End If
                    Else
                        comparerRow = Nothing
                    End If
                Next

                If IsNumeric(dtRow.Item("Rating")) Then
                    dtVal = CType(dtRow.Item("Rating"), Double)
                End If
            Else
                comparerRow = comparer.Rows(i)
            End If

            If comparerRow IsNot Nothing Then
                If IsNumeric(comparerRow.Item("Rating")) Then
                    comparerVal = CType(comparerRow.Item("Rating"), Double)
                End If
            End If

            If dtVal <> Double.NaN And comparerVal <> Double.NaN Then
                delta = Math.Round(dtVal - comparerVal, 3)
                perDelta = Math.Round((comparerVal - dtVal) / (dtVal) * 100, 2)
            End If

            Try
                diffDt.Rows.Add(
                        dtRow.Item("Tool").ToString,
                        dtRow.Item("Type").ToString,
                        IIf(Double.IsNaN(dtVal), "N/A", dtVal),
                        IIf(Double.IsNaN(comparerVal), "N/A", comparerVal),
                        IIf(Double.IsNaN(delta), "N/A", delta),
                        IIf(Double.IsNaN(perDelta), "N/A", perDelta),
                        IIf(Double.IsNaN(delta) Or Math.Abs(delta) > 0.1, "Fail", "Pass")
                       )
            Catch ex As Exception

            End Try

            If delta = Double.NaN Or delta <> 0 Then
                matching = False
            End If

            comparerVal = Double.NaN
            dtVal = Double.NaN
            delta = Double.NaN
            perDelta = Double.NaN
        Next

        Return New Tuple(Of Boolean, DataTable)(matching, diffDt)
    End Function

    <Extension()>
    Public Function TemplateVersion(ByVal file As FileInfo) As String
        Dim ver As String = Nothing
        Dim name As String = file.Name.Replace(file.Extension, "")
        Dim pattern As New Regex("\d+(\.\d+)+")
        Dim sMatch As Match = pattern.Match(name)

        If sMatch.Success Then
            ver = sMatch.Value
        Else
            ver = "-"
        End If

        Return ver
    End Function
End Module