
Imports System.Data.SqlClient
Imports System.Security.Principal
Imports Oracle.ManagedDataAccess.Client

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

        Try
            Using impersonatedUser As WindowsImpersonationContext = Impersonator.Impersonate(),
                sqlCon As New SqlConnection(ActiveDatabase),
                SQLAdapter = New SqlDataAdapter(SQLCommand, sqlCon)

                sqlCon.Open()
                If ClearSaveToDataSet Then SaveToDataSet.Reset()
                SQLAdapter.Fill(SaveToDataSet)
                sqlCon.Close()
            End Using
        Catch ex As Exception
            'Console.WriteLine("Error: " & erNo.ToString & vbNewLine & ex.Message, "Error: " & erNo.ToString)
            errors = True
        End Try

        Return Not errors
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
        Dim errors As Boolean = False

        Try
            Using impersonatedUser As WindowsImpersonationContext = Impersonator.Impersonate(),
                    sqlCon = New SqlConnection(ActiveDatabase),
                    sqlCmd = New SqlCommand(SQLCommand, sqlCon)
                sqlCon.Open()
                sqlCmd.ExecuteNonQuery()
                sqlCon.Close()
            End Using
        Catch ex As Exception
            errors = True
            Throw
        End Try

        Return Not errors
    End Function

    Public Function safeSqlSender(ByVal SQLCommand As SQLCommand, ByVal ActiveDatabase As String, ByVal Impersonator As WindowsIdentity, ByRef erNo As String) As Boolean
        Using impersonatedUser As WindowsImpersonationContext = Impersonator.Impersonate()
            sqlCon = New SqlConnection(ActiveDatabase)
            Dim sqlCmd = SQLCommand
            sqlCmd.Connection = sqlCon
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
    Public Function safeSqlTransactionSender(ByVal SQLCommands As List(of SQLCommand), ByVal ActiveDatabase As String, ByVal Impersonator As WindowsIdentity, ByRef erNo As String) As Boolean
        Using impersonatedUser As WindowsImpersonationContext = Impersonator.Impersonate()
            sqlCon = New SqlConnection(ActiveDatabase)
            sqlCon.Open()
            Dim transaction = sqlCon.BeginTransaction()

            For Each cmd In SQLCommands
                Try
                    cmd.Connection = sqlCon
                    cmd.Transaction = transaction
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    Console.WriteLine("Error: " & erNo & vbNewLine & ex.Message, "Error: " & erNo)
                    Try
                        transaction.Rollback()
                    Catch ex2 As Exception
                        'TODO set error num
                        Return False
                    Finally
                        sqlCon.Close()
                    End Try
                
                    Return False
                End Try
            Next
            transaction.commit()
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


Module DoDaORACLE
    Private Const ordsDataSource = "(DESCRIPTION =    (ADDRESS = (PROTOCOL = TCP)(HOST = prd-scan)(PORT = 1521))    (CONNECT_DATA =      (SERVER = DEDICATED)      (SERVICE_NAME = ordsprd_batch.crowncastle.com)    )  )"
    Private Const isitDataSource = "(DESCRIPTION =    (ADDRESS = (PROTOCOL = TCP)(HOST = prd-scan)(PORT = 1521))    (CONNECT_DATA =      (SERVICE_NAME = isitprd_utl.crowncastle.com)      (SERVER = DEDICATED)    )  )"
    Private Const odsDataSource = "(DESCRIPTION =    (ADDRESS = (PROTOCOL = TCP)(HOST = prd-scan)(PORT = 1521))    (CONNECT_DATA =      (SERVICE_NAME = odsprd_app.crowncastle.com)      (SERVER = DEDICATED)    )  )"
    'Private Const isitDataSource = "(DESCRIPTION =    (ADDRESS = (PROTOCOL = TCP)(HOST = uat-scan)(PORT = 1521))    (CONNECT_DATA =      (SERVICE_NAME = isituat_batch.crowncastle.com)      (SERVER = DEDICATED)    )  )"
    Private Const ntoken = "270:207:234:213:204:207:258"
    Private Const wtoken = "366:264:339:216:357:159:192:297:171:216"
    Public Function OracleLoader(ByVal SQLCommand As String, ByVal SaveToTableName As String, ByRef SaveToDataSet As DataSet, ByVal erNo As Integer, ByVal db As String) As Boolean
        Dim oraDatasource As String
        Dim dt As New DataTable
        Select Case db
            Case "isit"
                oraDatasource = isitDataSource
            Case "ods"
                oraDatasource = odsDataSource
            Case Else
                'ORDS is the catch-all because it has links to the other DBs
                oraDatasource = ordsDataSource
        End Select
        dtClearer(SaveToTableName)
        Dim sb As OracleConnectionStringBuilder = New OracleConnectionStringBuilder()
        sb.DataSource = oraDatasource
        sb.UserID = token(ntoken)
        sb.Password = token(wtoken)
        'By default pooling = true which means that oracle moves connections to an inactive pool when they are closed by the program.
        'This makes reconnecting faster but was causing issues with the connection idle_time being exceeded.
        sb.Pooling = False
        Dim bOraSuccess As Boolean = True
        Using oraCon As New OracleConnection(sb.ToString())
            Try
                Dim oDa = New OracleDataAdapter(SQLCommand, oraCon)
                oDa.Fill(SaveToDataSet, SaveToTableName)

            Catch ex As Exception
                bOraSuccess = False
                Console.WriteLine(SQLCommand)
                'sendToast("Failure loading data:" & vbCrLf & ex.Message, "Error " & erNo)
            End Try
            oraCon.Close()
        End Using
        Return bOraSuccess
    End Function
End Module