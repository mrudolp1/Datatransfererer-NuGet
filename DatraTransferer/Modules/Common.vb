Imports System.IO

Module IDoDeclare
    Public ds As New DataSet
    Public queryPath As String = System.Windows.Forms.Application.StartupPath & "\Queries\"
    Public BUNumber As String = "812123"
    Public STR_ID As String = "A"
End Module

Public Module Common

    Function GetExistingModelQuery() As String
        Return QueryBuilderFromFile(queryPath & "Existing Model (SELECT).sql").Replace("[BU NUMBER]", BUNumber).Replace("[STRUCTURE_ID]", STR_ID)
    End Function

    Function SaveExistingModelQuery() As String
        Return QueryBuilderFromFile(queryPath & "Existing Model (IN_UP).sql").Replace("[BU NUMBER]", BUNumber).Replace("[STRUCTURE_ID]", STR_ID)
    End Function

    Function SaveFoundationQuery(ByVal fndID As String, ByVal fndType As String) As String
        Return QueryBuilderFromFile(queryPath & "Foundations (IN_UP).sql").Replace("'[Foundation ID]'", IIf(fndID = "", "NULL", "'" & fndID & "'")).Replace("[FOUNDATION TYPE]", fndType)
    End Function

    Function QueryBuilderFromFile(ByVal filetoread As String) As String
        Dim sqlReader As New StreamReader(filetoread)
        Dim temp As String = ""

        Do While sqlReader.Peek <> -1
            temp += sqlReader.ReadLine.Trim + vbNewLine
        Loop

        sqlReader.Dispose()
        sqlReader.Close()

        Return temp
    End Function
End Module

Public Class SQLParameter
    Public Property sqlDatatable As String
    Public Property sqlQuery As String

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal DataTableName As String, ByVal QueryFileName As String)
        sqlDatatable = DataTableName
        sqlQuery = QueryFileName
    End Sub
End Class