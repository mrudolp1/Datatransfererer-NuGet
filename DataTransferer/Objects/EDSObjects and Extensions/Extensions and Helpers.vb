Imports System.ComponentModel
Imports System.Security.Principal
Imports DevExpress.Spreadsheet
Imports System.IO
Imports System.Reflection
Imports DevExpress.DataAccess.Excel
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Module Extensions

    <Extension()>
    Public Function NullableToString(input As Object) As String
        'Extensions of objects are handled differently than all other extensions. This extension will be applicable to all type except object.
        'Reference: https://devblogs.microsoft.com/vbteam/extension-methods-and-late-binding-extension-methods-part-4/
        Dim inputString As String = ""
        If input IsNot Nothing Then
            Try
                inputString = input.ToString
            Catch ex As Exception
                Debug.Print("Failed to convert object to string.")
            End Try
        End If

        Return inputString
    End Function

    <Extension()>
    Public Sub NewLine(ByRef input As String, nextLine As String)
        input = input & vbCrLf & nextLine
    End Sub

    <Extension()>
    Public Function FormatDBValue(input As String) As String
        'Handles nullable values and quoatations needed for DB values

        'Single quotes are not permissible in fields in SQL tables
        'this additional piece will replace single quotes with 2 quotes to allow the fields to be formatted the same when loading back from EDS>
        If input IsNot Nothing Then
            If input.ToString.Contains("'") Then
                input = input.Replace("'", "''")
            End If
        End If

        If String.IsNullOrEmpty(input) Then
            FormatDBValue = "NULL"
        Else
            FormatDBValue = "'" & input & "'"
        End If

        Return FormatDBValue
    End Function

    <Extension()>
    Public Function AddtoDBString(startingString As String, newString As String, Optional isDBValue As Boolean = False) As String
        'isValue should be false if you're creating a string of column names. They should not be in single quotes like the values.
        If isDBValue Then newString = newString.FormatDBValue

        If String.IsNullOrEmpty(startingString) Then
            startingString = newString
        Else
            startingString += ", " & newString
        End If

        Return startingString
    End Function

    <Extension()>
    Public Function SelectDBString(startingString As String, newString As String, Optional isDBValue As Boolean = False) As String
        'isValue should be false if you're creating a string of column names. They should not be in single quotes like the values.
        If isDBValue Then newString = newString.FormatDBValue

        If String.IsNullOrEmpty(startingString) Then
            startingString = newString
        Else
            startingString += "AND " & newString
        End If

        Return startingString
    End Function

    '<Extension()>
    'Public Function AddtoDBString(startingString As String, newObject As Object, Optional isDBValue As Boolean = False) As String
    '    'isValue should be false if you're creating a string of column names. They should not be in single quotes like the values.
    '    Dim newString As String = ""

    '    If isDBValue Then newString = newString.FormatDBValue

    '    If String.IsNullOrEmpty(startingString) Then
    '        startingString = newString
    '    Else
    '        startingString += ", " & newString
    '    End If

    '    Return startingString
    'End Function

    <Extension()>
    Public Function AddtoDBStringOG(startingString As String, newString As String, Optional isDBValue As Boolean = False) As String
        'isValue should be false if you're creating a string of column names. They should not be in single quotes like the values.
        'If String.IsNullOrEmpty(newString) Then Return startingString

        If isDBValue Then newString = newString.FormatDBValue

        If String.IsNullOrEmpty(startingString) Then
            startingString = newString
        Else
            startingString += ", " & newString
        End If

        Return startingString
    End Function

    '<Extension()>
    'Public Function ToDBString(aString As String, Optional isValue As Boolean = True) As String
    '    If aString = String.Empty Or aString Is Nothing Then
    '        Return "NULL"
    '    Else
    '        If isValue Then aString = "'" & aString & "'"
    '        Return aString
    '    End If
    'End Function

    '<Extension()>
    'Public Function AddtoDBString(astring As String, ByRef newString As String, Optional isValue As Boolean = True) As String
    '    'isValue should be false if you're creating a string of column names. They should not be in single quotes like the values.
    '    If astring = String.Empty Or astring Is Nothing Then
    '        astring = newString.ToDBString(isValue)
    '    Else
    '        astring += ", " & newString.ToDBString(isValue)
    '    End If
    '    Return astring
    'End Function

    '<Extension()>
    'Public Function GetDistinct(Of T As EDSObject)(alist As List(Of T)) As List(Of T)
    '    'Notes: Removes duplicates from list of tnxDatabaseEntry by using their CompareMe function
    '    'Making this generic (Of T As tnxDatabaseEntry) allows it to work for all subclasses of tnxDatabaseEntry

    '    Dim distinctList As New List(Of T)

    '    For Each item In alist
    '        Dim addToList As Boolean = True
    '        For Each distinctItem In distinctList
    '            If item.CompareMe(distinctItem) Then
    '                'Not distinct
    '                addToList = False
    '                Exit For
    '            End If
    '        Next
    '        If addToList Then distinctList.Add(item)
    '    Next

    '    Return distinctList
    'End Function

    <Extension()>
    Public Function EDSDELETEListQueryBuilder(Of T As EDSObjectWithQueries)(alist As List(Of T)) As String
        EDSDELETEListQueryBuilder = ""
        Dim currentSortedList As List(Of T) = alist.ToList

        For i As Integer = 0 To currentSortedList.Count - 1
            EDSDELETEListQueryBuilder += currentSortedList(i).SQLDelete
        Next

        Return EDSDELETEListQueryBuilder
    End Function

    <Extension()>
    Public Function EDSListQueryBuilder(Of T As EDSObjectWithQueries)(alist As List(Of T), Optional prevList As List(Of T) = Nothing, Optional ByVal AllowUpdate As Boolean = True) As String

        EDSListQueryBuilder = ""

        'Create a shallow copy of the lists for sorting and comparison
        'Sort lists by ID descending with Null IDs at the bottom
        Dim currentSortedList As List(Of T) = alist.ToList
        currentSortedList.Sort()
        currentSortedList.Reverse()

        Dim prevSortedList As List(Of T)
        If prevList Is Nothing Then
            prevSortedList = New List(Of T)
        Else
            prevSortedList = prevList.ToList
            prevSortedList.Sort()
            prevSortedList.Reverse()
        End If

        Dim i As Integer = 0
        Do While i <= Math.Max(currentSortedList.Count, prevSortedList.Count) - 1

            If i > currentSortedList.Count - 1 Then
                'Delete items in previous list if there is nothing left in current list
                EDSListQueryBuilder += prevSortedList(i).SQLDelete
            ElseIf i > prevSortedList.Count - 1 Then
                'Insert items in current list if there is nothing left in previous list
                EDSListQueryBuilder += currentSortedList(i).SQLInsert
            Else
                'Compare IDs
                If currentSortedList(i).ID = prevSortedList(i).ID And AllowUpdate Then
                    EDSListQueryBuilder += prevSortedList(i).SQLSetID
                    If Not currentSortedList(i).Equals(prevSortedList(i)) Then
                        'Update existing
                        EDSListQueryBuilder += currentSortedList(i).SQLUpdate
                    Else
                        'Save Results Only
                        EDSListQueryBuilder += currentSortedList(i).Results.EDSResultQuery
                    End If
                ElseIf currentSortedList(i).ID < prevSortedList(i).ID Then
                    EDSListQueryBuilder += prevSortedList(i).SQLDelete
                    currentSortedList.Insert(i, Nothing)
                Else
                    'currentSortedList(i).ID > prevSortedList(i).ID
                    EDSListQueryBuilder += currentSortedList(i).SQLInsert
                    prevSortedList.Insert(i, Nothing)
                End If
            End If

            i += 1

        Loop

        Return EDSListQueryBuilder

    End Function

    <Extension()>
    Public Function EDSResultQuery(Of T As EDSResult)(alist As List(Of T), Optional ByVal ResultsParentID As Integer? = Nothing) As String

        EDSResultQuery = ""

        ''DHS 6/29/2023
        ''Prevent Saving Guyed Tower Results because they currently break the query.
        ''Needs to be removed after Peter updates XMLOutput
        'If alist.FirstOrDefault?.ParentStructure?.tnx?.geometry Is Nothing OrElse
        '    alist.FirstOrDefault.ParentStructure.tnx.geometry.TowerType.Contains("Guyed") Then
        '    Return EDSResultQuery
        'End If

        For Each result In alist
            'debugging top guy pull-off NaN values 7/6/23 - DHS
            'If result.result_lkup.Contains("top guy pull-off") Then
            '    Debug.WriteLine("top guy pull-off")
            'End If

            If result.rating Is Nothing Then Continue For

            If Not result.foreign_key.HasValue Then result.foreign_key = ResultsParentID

            EDSResultQuery += result.Insert(result.foreign_key) & vbCrLf

        Next

        Return EDSResultQuery

    End Function

    <Extension()>
    Public Function CheckChange(Of T)(value1 As T, value2 As T, ByRef changes As List(Of AnalysisChange), Optional categoryName As String = Nothing, Optional fieldName As String = Nothing) As Boolean

        If value1 Is Nothing And value2 Is Nothing Then Return True

        'Check if this is an EDSObject
        Dim EDSValue1 As EDSObject = TryCast(value1, EDSObject)
        Dim EDSValue2 As EDSObject = TryCast(value2, EDSObject)
        If EDSValue1 IsNot Nothing AndAlso EDSValue2 IsNot Nothing Then
            Return EDSValue1.Equals(EDSValue2, changes)
        End If

        'Check if this is a collection (list) of any object, iterate through if needed
        Dim CollectionValue1 As IEnumerable(Of Object) = TryCast(value1, IEnumerable(Of Object))
        Dim CollectionValue2 As IEnumerable(Of Object) = TryCast(value2, IEnumerable(Of Object))
        If CollectionValue1 IsNot Nothing AndAlso CollectionValue2 IsNot Nothing Then
            If CollectionValue1.Count <> CollectionValue2.Count Then
                changes.Add(New AnalysisChange(categoryName, fieldName & "Quantity", CollectionValue1.Count.ToString, CollectionValue2.Count.ToString))
                Return False
            Else
                'Check if this is a collection (list) of EDSObjects, iterate through if needed
                Dim EDSCollectionValue1 As IEnumerable(Of EDSObject) = TryCast(value1, IEnumerable(Of EDSObject))
                Dim EDSCollectionValue2 As IEnumerable(Of EDSObject) = TryCast(value2, IEnumerable(Of EDSObject))
                Dim ListChanges As Boolean = True
                If EDSCollectionValue1 IsNot Nothing AndAlso EDSCollectionValue2 IsNot Nothing Then
                    For i As Integer = 0 To EDSCollectionValue1.Count - 1
                        'If fieldName IsNot Nothing Then fieldName = fieldName & " (" & i & ")"
                        ListChanges = If(EDSCollectionValue1(i).Equals(EDSCollectionValue2(i), changes), ListChanges, False)
                    Next
                Else
                    For i As Integer = 0 To CollectionValue1.Count - 1
                        If fieldName IsNot Nothing Then fieldName = fieldName & " (" & i & ")"
                        If Not CollectionValue1(i).Equals(CollectionValue2(i)) Then
                            changes.Add(New AnalysisChange(categoryName, fieldName, CollectionValue1(i).ToString, CollectionValue2(i).ToString))
                            ListChanges = False
                        End If
                    Next
                End If
                Return ListChanges
            End If
        End If

        'Try to compare values directly
        Try
            If value1 Is Nothing Xor value2 Is Nothing Then
                changes.Add(New AnalysisChange(categoryName, fieldName, value1.ToString, value2.ToString))
                Return False
            ElseIf Not value1.Equals(value2) Then
                changes.Add(New AnalysisChange(categoryName, fieldName, value1.ToString, value2.ToString))
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            changes.Add(New AnalysisChange(categoryName, fieldName, "Comparison Failed", ""))
            Return False
        End Try

    End Function

    <Extension()>
    Public Function CheckChange(CollectionValue1 As IEnumerable(Of Object), CollectionValue2 As IEnumerable(Of Object), ByRef changes As List(Of AnalysisChange), Optional categoryName As String = Nothing, Optional fieldName As String = Nothing) As Boolean

        If CollectionValue1 IsNot Nothing AndAlso CollectionValue2 IsNot Nothing Then
            If CollectionValue1.Count <> CollectionValue2.Count Then
                changes.Add(New AnalysisChange(categoryName, fieldName & "Quantity", CollectionValue1.Count.ToString, CollectionValue2.Count.ToString))
                Return False
            Else
                For i As Integer = 0 To CollectionValue1.Count - 1
                    If fieldName IsNot Nothing Then fieldName = fieldName & " (" & i & ")"
                    CollectionValue1(i).CheckChange(CollectionValue2(i), changes, categoryName, fieldName)
                Next
            End If
        End If

    End Function

    <Extension()>
    Public Function GetExtension(docFormat As DocumentFormat) As String

        Select Case docFormat
            Case DocumentFormat.Xlsx, DocumentFormat.OpenXml
                Return ".xlsx"
            Case DocumentFormat.Xls
                Return ".xls"
            Case DocumentFormat.Xlsm
                Return ".xlsm"
            Case DocumentFormat.Xlsb
                Return ".xlsb"
            Case DocumentFormat.Xlt
                Return ".xlt"
            Case DocumentFormat.Xltm
                Return ".xltm"
            Case DocumentFormat.Xltx
                Return ".xltx"
            Case DocumentFormat.XmlSpreadsheet2003
                Return ".xml"
            Case DocumentFormat.Text
                Return ".txt"
            Case DocumentFormat.Csv
                Return ".csv"
            Case Else
                Return ""
        End Select

    End Function

End Module

Public Module myLittleHelpers

    ''' <summary>
    ''' Increments the name following the same pattern that file explorer uses (- Copy, - Copy (2), - Copy (3)...)
    ''' </summary>
    ''' <paramname="destFileName"></param>
    ''' <returns></returns>
    Public Function IncrementFileName(ByVal destFileName As String) As String
        If File.Exists(destFileName) Then
            Dim fileName = Path.GetFileNameWithoutExtension(destFileName)
            Dim ext = Path.GetExtension(destFileName)
            Dim fileNameEndPosition = fileName.IndexOf(" - Copy")
            Dim newFileName = fileName & " - Copy"
            If fileNameEndPosition > 0 Then
                Dim baseFileName = fileName.Substring(0, fileNameEndPosition)
                Dim fileNameEnd = fileName.Substring(fileNameEndPosition)
                Dim copyNum = 0
                'Split the string on parenthesis, get the first value in between (odd index), and try to convert it to integer
                Integer.TryParse(fileNameEnd.Split("("c, ")"c).Where(Function(item, index) index Mod 2 <> 0).FirstOrDefault(), copyNum)
                copyNum += 1

                newFileName = baseFileName & " - Copy (" & copyNum.ToString() & ")"
            End If
            Return IncrementFileName(Path.Combine(Path.GetDirectoryName(destFileName), newFileName & ext))
        Else
            Return destFileName
        End If
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="resourceName"></param>
    ''' <returns></returns>
    Public Function LoadResourceFileStream(resourceName As String) As String
        Dim assembly As Assembly = Assembly.GetExecutingAssembly()
        Dim output As String = ""

        Using stream As Stream = assembly.GetManifestResourceStream(resourceName)
            If Not stream Is Nothing Then
                Using reader As New StreamReader(stream)
                    output = reader.ReadToEnd()
                End Using
            End If
        End Using

        Return output
    End Function

    Public Function FindResourcebyFileName(fileName As String) As String
        Return Assembly.GetExecutingAssembly().GetManifestResourceNames().Single(Function(str) str.EndsWith(fileName))
    End Function


    Public Function trueFalseYesNo(input As String) As Boolean?
        If input.ToLower = "yes" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function trueFalseYesNo(input As Boolean?) As String
        If input Then
            Return "Yes"
        Else
            Return "No"
        End If
    End Function
    Public Function BooltoBitString(input As Boolean?) As String
        If input Then
            Return "1"
        Else
            Return "0"
        End If
    End Function

    Public Function DBtoNullableInt(ByVal item As Object) As Integer?
        If IsDBNull(item) Then
            Return Nothing
        Else
            Try
                Return CInt(item)
            Catch ex As Exception
                Return Nothing
            End Try
        End If
    End Function

    Public Function DBtoNullableDbl(ByVal item As Object, Optional ByVal precision As Integer = 4) As Double?
        If IsDBNull(item) Then
            Return Nothing
        Else
            Try
                Return Math.Round(CDbl(item), precision)
            Catch ex As Exception
                Return Nothing
            End Try
        End If
    End Function

    Public Function DBtoNullableDec(ByVal item As Object, Optional ByVal precision As Integer = 4) As Decimal?
        If IsDBNull(item) Then
            Return Nothing
        Else
            Try
                Return Math.Round(CDbl(item), precision)
            Catch ex As Exception
                Return Nothing
            End Try
        End If
    End Function

    Public Function DBtoNullableBool(ByVal item As Object) As Boolean?
        If IsDBNull(item) Then
            Return Nothing
        Else
            Try
                Return CBool(item)
            Catch ex As Exception
                Return Nothing
            End Try
        End If
    End Function

    Public Function DBtoStr(ByVal item As Object) As String
        'Strings are nullable, but the default value is "" so that's what should be used
        'CStr(Nothing) = "" which is the default value of a string, so that's what we should use with DBNull too. This works better for comparing EDS to ERI opbjects
        If IsDBNull(item) Then
            Return ""
        Else
            Try
                Return CStr(item)
            Catch ex As Exception
                Return ""
            End Try
        End If
    End Function

    Public Function GetTypeNullable(Of T)(ByVal obj As T) As Type
        'This should return the type of a Nullable even if it's null
        If Nullable.GetUnderlyingType(GetType(T)) Is Nothing Then
            Return GetType(T)
        Else
            Return Nullable.GetUnderlyingType(GetType(T))
        End If
    End Function

#Region "Maestro Helpers"
    'pass in Work Order [string]
    'Public Function MessageLog(ByVal wo As String, ByVal dbError As String, ByVal logPath As String, Optional ByVal successLog As Boolean = False) As Boolean
    '    'our log format will be: WO, Success, Error
    '    Dim dt As DateTime = DateTime.Now


    '    Dim catString As String

    '    catString = wo & ", " & successLog & "," & dbError & "," & dt

    '    If Not File.Exists(logPath) Then
    '        File.Create(logPath).Close()

    '        FileOpen(1, logPath, OpenMode.Append)
    '        PrintLine(1, "WO, Success, Error, Time")
    '        PrintLine(1, catString)

    '        'PrintLine(1, "***************")
    '        FileClose(1)
    '    Else
    '        FileOpen(1, logPath, OpenMode.Append)
    '        PrintLine(1, catString)

    '        'PrintLine(1, "***************")
    '        FileClose(1)
    '    End If
    '    Return True
    'End Function

    Public Function MessageLog(ByVal wo As String, ByVal dbError As String, ByVal logPath As String, Optional ByVal successLog As Boolean = False) As Boolean
        Dim dt As DateTime = DateTime.Now
        Dim catString As String = wo & ", " & successLog & "," & dbError & "," & dt.ToString("yyyy-MM-dd HH:mm:ss")
        Try
            Using sw As New StreamWriter(logPath, True)
                If Not File.Exists(logPath) Then
                    sw.WriteLine("WO, Success, Error, Time")
                End If
                sw.WriteLine(catString)
            End Using
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function


    Public Function CheckIfNull(ByRef input As Object, Optional ByRef intBool As Boolean = False) As Object
        ' Check if the input is an integer, double, decimal or if intBool is set to True
        If TypeOf input Is Integer Or TypeOf input Is Double Or TypeOf input Is Decimal Or intBool Then
            ' If the input is DBNull, return 0
            If IsDBNull(input) Then
                Return 0
            End If
            ' Otherwise, return the input
            Return input
        Else
            ' If the input is DBNull, return an empty string
            If IsDBNull(input) Then
                Return ""
            End If
            ' Otherwise, return the input with any double single quotes replaced with single quotes
            Try
                Return input.Replace("''", "'")
            Catch ex As Exception
                ' If there is an exception, return the input without modification
                Return input
            End Try
        End If
    End Function


#End Region
End Module






