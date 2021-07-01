
Imports System.Data.SqlClient
Imports System.Security.Principal

Module DoDaSQL
    Private SQLAdapter As SqlDataAdapter
    Private sqlCon As New SqlConnection

#Region "Main SQL Functions"
    <DebuggerStepThrough()>
    Function sqlLoader(ByVal SQLCommand As String, ByVal SQLSource As String, ByVal SaveToDataSet As DataSet, ByVal ActiveDatabase As String, ByVal Impersonator As WindowsIdentity, ByRef erNo As String) As Boolean
        ClearDataTable(SQLSource, SaveToDataSet)
        Using impersonatedUser As WindowsImpersonationContext = Impersonator.Impersonate()
            sqlCon = New SqlConnection(ActiveDatabase)
            sqlCon.Open()

            Try
                SQLAdapter = New SqlDataAdapter(SQLCommand, sqlCon)
                SQLAdapter.Fill(SaveToDataSet, SQLSource)
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

    <DebuggerStepThrough()>
    Function sqlSender(ByVal SQLCommand As String, ByVal ActiveDatabase As String, ByVal Impersonator As WindowsIdentity, ByRef erNo As String) As Boolean
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
    Private Function token(s As String) As String
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