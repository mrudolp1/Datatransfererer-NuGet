
Imports System.Data.SqlClient
Imports System.Security.Principal

Module DoDaSQL
    Private SQLAdapter As SqlDataAdapter
    Private sqlCon As New SqlConnection

#Region "Main SQL Functions"
    <DebuggerStepThrough()>
    Public Function sqlLoader(ByVal SQLCommand As String, ByVal SaveToTableName As String, ByVal SaveToDataSet As DataSet, ByVal ActiveDatabase As String, ByVal Impersonator As WindowsIdentity, ByRef erNo As String) As Boolean
        ClearDataTable(SaveToTableName, SaveToDataSet)
        Using impersonatedUser As WindowsImpersonationContext = Impersonator.Impersonate()
            sqlCon = New SqlConnection(ActiveDatabase)
            sqlCon.Open()

            Try
                SQLAdapter = New SqlDataAdapter(SQLCommand, sqlCon)
                SQLAdapter.Fill(SaveToDataSet, SaveToTableName)
            Catch ex As Exception
                sqlCon.Close()
                Console.WriteLine("Error: " & erNo & vbNewLine & ex.Message, "Error: " & erNo)
                erNo = "Error: " & erNo & vbNewLine & ex.Message
                Return False
            End Try

            sqlCon.Close()
        End Using

        Return True
    End Function

    Public Function sqlLoader(ByVal SQLCommand As String, ByRef SaveToDataSet As DataSet, ByVal ActiveDatabase As String, ByVal Impersonator As WindowsIdentity, Optional ByVal erNo As Integer = 0, Optional ByVal ClearSaveToDataSet As Boolean = True) As Boolean

        Dim errors As Boolean = False

        Using impersonatedUser As WindowsImpersonationContext = Impersonator.Impersonate()
            Using sqlCon As New SqlConnection(ActiveDatabase)
                sqlCon.Open()

                Try
                    If ClearSaveToDataSet Then SaveToDataSet.Reset()
                    SQLAdapter = New SqlDataAdapter(SQLCommand, sqlCon)
                    SQLAdapter.Fill(SaveToDataSet)
                Catch ex As Exception
                    Console.WriteLine("Error: " & erNo.ToString & vbNewLine & ex.Message, "Error: " & erNo.ToString)
                    errors = True
                End Try
            End Using
        End Using

        Return True
    End Function

    Public Function sqlLoader(ByVal SQLCommands As List(Of String), ByVal SaveToDataTables As List(Of String), ByRef SaveToDataSet As DataSet, ByVal ActiveDatabase As String, ByVal Impersonator As WindowsIdentity, Optional ByVal erNo As Integer = 0, Optional ByVal ClearSaveToDataTable As Boolean = True) As Boolean

        'This overload accepts a list of SQL commands and table names and only opens the SQL connection one time to execute all commands. - DHS

        If SQLCommands.Count <> SaveToDataTables.Count Then Return False

        Dim errors As Boolean = False

        Using impersonatedUser As WindowsImpersonationContext = Impersonator.Impersonate()
            Using sqlCon As New SqlConnection(ActiveDatabase)
                sqlCon.Open()

                For i = 0 To SQLCommands.Count - 1
                    Try
                        If ClearSaveToDataTable Then ClearDataTable(SaveToDataTables(i), SaveToDataSet)
                        Using SQLAdapter As New SqlDataAdapter(SQLCommands(i), sqlCon)
                            SQLAdapter.Fill(SaveToDataSet, SaveToDataTables(i))
                        End Using
                    Catch ex As Exception
                        erNo = erNo + i
                        Console.WriteLine("Error: " & erNo.ToString & vbNewLine & ex.Message, "Error: " & erNo.ToString)
                        errors = True
                    End Try
                Next
            End Using
        End Using

        Return True
    End Function

    Public Function sqlLoader(ByVal SQLCommand As List(Of String), ByVal tableName As String, ByVal SaveToDataSet As DataSet, ByVal ActiveDatabase As String, ByVal Impersonator As WindowsIdentity, ByVal erNo As Integer) As Boolean

        'This overload accepts a list of SQL commands and adds all the output to one datatable. - DHS

        Dim errors As Boolean = False

        Using impersonatedUser As WindowsImpersonationContext = Impersonator.Impersonate()
            Using sqlCon As New SqlConnection(ActiveDatabase)
                sqlCon.Open()

                For i = 0 To SQLCommand.Count - 1
                    Try
                        ClearDataTable(tableName, SaveToDataSet)
                        Using SQLAdapter As New SqlDataAdapter(SQLCommand(i), sqlCon)
                            SQLAdapter.Fill(SaveToDataSet, tableName)
                        End Using
                    Catch ex As Exception
                        erNo = erNo + i
                        Console.WriteLine("Error: " & erNo & vbNewLine & ex.Message, "Error: " & erNo.ToString)
                        errors = True
                    End Try
                Next
            End Using
        End Using

        Return True
    End Function

    <DebuggerStepThrough()>
    Public Function sqlSender(ByVal SQLCommand As String, ByVal ActiveDatabase As String, ByVal Impersonator As WindowsIdentity, ByRef erNo As String) As Boolean
        Using impersonatedUser As WindowsImpersonationContext = Impersonator.Impersonate()
            sqlCon = New SqlConnection(ActiveDatabase)
            Dim sqlCmd = New SqlCommand(SQLCommand, sqlCon)
            sqlCon.Open()

            Try
                sqlCmd.ExecuteNonQuery()
            Catch ex As Exception
                sqlCon.Close()
                Console.WriteLine("Error: " & erNo & vbNewLine & ex.Message, "Error: " & erNo)
                erNo = "Error: " & erNo & vbNewLine & ex.Message
                Return False
            End Try

            sqlCon.Close()
        End Using

        Return True
    End Function
#End Region

    <DebuggerStepThrough()>
    Public Function token(s As String) As String
        Dim m As String = ""
        For x As Integer = 0 To 1000
            Try
                m = m & Chr(s.Split(":")(x) / Chr(51).ToString)
            Catch
                Exit For
            End Try
        Next
        Return m
    End Function

    <DebuggerStepThrough()>
    Private Sub ClearDataTable(ByVal SQLSource As String, ByVal SaveToDataSet As DataSet)
        Try
            If SaveToDataSet.Tables.Contains(SQLSource) Then
                SaveToDataSet.Tables(SQLSource).Clear()
                SaveToDataSet.Tables.Remove(SQLSource)
            End If
        Catch
        End Try
    End Sub
End Module