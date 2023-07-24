
Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions

Public Module UnitTestingExtensions

    <Extension()>
    Public Function WriteAllToFile(ByVal writer As String, ByVal filepath As String, Optional ByVal append As Boolean = False) As Boolean
        Try
            ' Use a StreamWriter to write to the file
            ' The 'True' argument appends to the file if it already exists
            Using sw As New StreamWriter(filepath, append)
                ' Write the log message to the file
                sw.Write(writer)
            End Using

            Return True
        Catch ex As Exception
            ' Handle the exception
            Console.WriteLine("Error writing to log file: " & ex.Message)
            Return False
        End Try
    End Function

    <Extension()>
    Public Function WriteLineToFile(ByVal writer As String, ByVal filepath As String, Optional ByVal append As Boolean = False) As Boolean
        Try
            ' Use a StreamWriter to write to the file
            ' The 'True' argument appends to the file if it already exists
            Using sw As New StreamWriter(filepath, append)
                ' Write the log message to the file
                sw.WriteLine(writer)
            End Using

            Return True
        Catch ex As Exception
            ' Handle the exception
            Console.WriteLine("Error writing to log file: " & ex.Message)
            Return False
        End Try
    End Function

    'Archive files in a directory to another directory.
    <Extension()>
    Public Function ArchiveFiles(ByVal dirTo As DirectoryInfo, ByVal dirFrom As String) As String
        Dim archiveLog As String = ""

        For Each fold As DirectoryInfo In New DirectoryInfo(dirFrom).GetDirectories
            If Not fold.Name.ToLower.Contains("archive") Then
                Try
                    fold.MoveTo(dirTo.FullName & "\" & fold.Name)
                Catch
                End Try
                archiveLog.NewLine("DEBUG | " & fold.Name & " moved to " & dirTo.FullName.Replace(dirFrom, ""))
            End If
        Next

        For Each file As FileInfo In New DirectoryInfo(dirFrom).GetFiles
            file.MoveTo(dirTo.FullName & "\" & file.Name)
            archiveLog.NewLine("DEBUG | " & file.Name & " moved to " & dirTo.FullName.Replace(dirFrom, ""))
        Next

        Return archiveLog
    End Function



    'Extension for datatables to export to CSV using the datatabletocsv method
    '''Requires a filepath for where to save the csv
    <Extension()>
    Public Sub ToCSV(ByVal dt As DataTable, ByVal FilePath As String)
        DatatableToCSV(dt, FilePath)
    End Sub

    'Replaces all special characters in a string that aren't allowed in file folder or file names
    <Extension()>
    Public Function ToDirectoryString(ByVal str As String) As String
        str = str.Replace("#", "")
        str = str.Replace("%", "")
        str = str.Replace("&", "")
        str = str.Replace("{", "")
        str = str.Replace("}", "")
        str = str.Replace("/", "")
        str = str.Replace("\", "")
        str = str.Replace("<", "")
        str = str.Replace(">", "")
        str = str.Replace("*", "")
        str = str.Replace("?", "")
        str = str.Replace("$", "")
        str = str.Replace("!", "")
        str = str.Replace("'", "")
        str = str.Replace("""", "")
        str = str.Replace(":", "")
        str = str.Replace("@", "")
        str = str.Replace("+", "")
        str = str.Replace("`", "")
        str = str.Replace("|", "")
        str = str.Replace("=", "")

        Return str
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

    <Extension()>
    Public Sub NewLine(ByRef input As String, nextLine As String)
        If input = "" Then
            input = nextLine
        Else
            input = input & vbCrLf & nextLine
        End If
    End Sub

End Module