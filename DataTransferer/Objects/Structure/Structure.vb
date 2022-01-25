Imports System.ComponentModel
Imports System.Security.Principal
Imports DevExpress.Spreadsheet
Imports System.IO
Imports DevExpress.DataAccess.Excel
Imports System.Runtime.CompilerServices

Public Module Extensions

    <Extension()>
    Public Function ToDBString(aString As String, Optional isValue As Boolean = True) As String
        If aString = String.Empty Or aString Is Nothing Then
            Return "NULL"
        Else
            If isValue Then aString = "'" & aString & "'"
            Return aString
        End If
    End Function

    <Extension()>
    Public Function AddtoDBString(astring As String, ByRef newString As String, Optional isValue As Boolean = True) As String
        'isValue should be false if you're creating a string of column names. They should not be in single quotes like the values.
        If astring = String.Empty Or astring Is Nothing Then
            astring = newString.ToDBString(isValue)
        Else
            astring += ", " & newString.ToDBString(isValue)
        End If
        Return astring
    End Function

    <Extension()>
    Public Function GetDistinct(Of T As EDSObject)(alist As List(Of T)) As List(Of T)
        'Notes: Removes duplicates from list of tnxDatabaseEntry by using their CompareMe function
        'Making this generic (Of T As tnxDatabaseEntry) allows it to work for all subclasses of tnxDatabaseEntry

        Dim distinctList As New List(Of T)

        For Each item In alist
            Dim addToList As Boolean = True
            For Each distinctItem In distinctList
                If item.CompareMe(distinctItem) Then
                    'Not distinct
                    addToList = False
                    Exit For
                End If
            Next
            If addToList Then distinctList.Add(item)
        Next

        Return distinctList
    End Function

End Module

Public Module myLittleHelpers
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

    Public Function DBtoNullableDbl(ByVal item As Object) As Double?
        If IsDBNull(item) Then
            Return Nothing
        Else
            Try
                Return Math.Round(CDbl(item), 6)
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
End Module

Partial Public Class EDSObject

    Public Property ID As Integer?
    Public Property activeDatabase As String
    Public Property databaseIdentity As String


    Public Function CompareMe(Of T As EDSObject)(toCompare As T, Optional SetID As Boolean = False, Optional ByRef strDiff As String = Nothing) As Boolean
        'Compare another tnxDatabaseEntry object to itself using the objects comparer.
        'Making this generic (Of T As tnxDatabaseEntry) allows it to work for all subclasses of tnxDatabaseEntry

        If toCompare Is Nothing Or Me.GetType() IsNot toCompare.GetType() Then Return False

        Dim comparer As New ObjectsComparer.Comparer(Of T)()

        If strDiff Is Nothing Then
            CompareMe = comparer.Compare(CType(Me, T), toCompare)
        Else
            Dim Differences As IEnumerable(Of ObjectsComparer.Difference)
            CompareMe = comparer.Compare(CType(Me, T), toCompare, Differences)
            strDiff = String.Join(vbCrLf, Differences)
        End If

        If CompareMe And SetID Then Me.ID = toCompare.ID

        Return CompareMe

    End Function

End Class

Partial Public Class structure_model
    Inherits EDSObject

    Public Property BU As String
    Public Property strID As String
    Public Property order As String
    Public Property tnx As tnxModel
    Public Property foundations As EDSFoundationGroup
    Public Property connections As DataTransfererCCIplate
    Public Property pole As DataTransfererCCIpole

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal BU As String, ByVal Strucutre_ID As String, ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        'Create from database using BU and Structure ID


    End Sub

    Public Sub New(ByVal WO As String, ByVal MyDataSet As DataSet, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        'Create from database using WO


    End Sub

#End Region

#Region "Save to EDS"
    Public Sub saveToEDS()

    End Sub

#End Region
End Class

Partial Public Class EDSFoundationGroup
    Inherits EDSObject

    Public Property foundations As New List(Of EDSDataTransferer)

End Class


Partial Public Class EDSDataTransferer

    Private _newWorkBook As New Workbook

    Public Property workBookPath As String

End Class




