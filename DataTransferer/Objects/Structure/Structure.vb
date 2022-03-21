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

    <Extension()>
    Public Function CompareEDSLists(Of T As EDSObject)(List1 As List(Of T), List2 As List(Of T), Optional SetID As Boolean = False) As Boolean
        'Compare lists of EDSObjects that are not in the same order.
        'If SetID then items in list1 will be updated with the IDs of matching items in list2.
        'Typical use case list1 would be excel tools from the current analysis, list2 would be the tools on EDS.
        CompareEDSLists = True

        If List1.Count <> List2.Count Then
            CompareEDSLists = False
            'If setting IDs from list2, we need to keep going even tho we no the overall lists are different
            If Not SetID Then Return CompareEDSLists
        End If

        For Each item In List1
            Dim CompareItem As Boolean
            For Each item2 In List2
                CompareItem = item.CompareMe(item2, SetID)
                If CompareItem Then Exit For
            Next
            If Not CompareItem Then
                CompareEDSLists = False
                If Not SetID Then Exit For
            End If
        Next

        Return CompareEDSLists
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

    Public Function GetTypeNullable(Of T)(ByVal obj As T) As Type
        'This should return the type of a Nullable even if it's null
        If Nullable.GetUnderlyingType(GetType(T)) Is Nothing Then
            Return GetType(T)
        Else
            Return Nullable.GetUnderlyingType(GetType(T))
        End If
    End Function
End Module

'Partial Public Class EDSQueries
'    'WIP - This depends on changes we might make to the query method - DHS 2/1/22
'    Public Property BaseFolder As String = System.Windows.Forms.Application.StartupPath & "\Data Transferer Queries\"

'    Public Property ClassFolder As String
'    <Description("Upload base on active model for BU and Structure ID"), DisplayName("Upload Query")>
'    Public Property QueryUp As String
'    <Description("Sub query for uploading from a higher level"), DisplayName("Upload Sub Query")>
'    Public Property QueryUpSub As String
'    <Description("Download base on WO"), DisplayName("Download Query for Specific WO")>
'    Public Property QueryDownWO As String
'    <Description("Download base on active model for BU and Structure ID"), DisplayName("Download Query for Active Model")>
'    Public Property QueryDownActive As String

'    Public Sub New()

'    End Sub

'    Public Sub New(ByVal FolderPathRelativetoQueriesFolder As String, ByVal QueryUp As String, ByVal QueryUpSub As String, ByVal QueryDownWO As String, ByVal QueryDownActive As String)
'        Me.ClassFolder = Me.BaseFolder & FolderPathRelativetoQueriesFolder
'        Me.QueryUp = Me.BaseFolder & FolderPathRelativetoQueriesFolder & QueryUp
'        Me.QueryUpSub = Me.BaseFolder & FolderPathRelativetoQueriesFolder & QueryUpSub
'        Me.QueryDownWO = Me.BaseFolder & FolderPathRelativetoQueriesFolder & QueryDownWO
'        Me.QueryDownActive = Me.BaseFolder & FolderPathRelativetoQueriesFolder & QueryDownActive
'    End Sub
'End Class

Partial Public Class EDSObject

    Public Property ID As Integer?
    Public Property BU As String
    Public Property strID As String
    Public Property WO As String
    'Public Property EDSQueries As New EDSQueries
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

    Public Property tnx As tnxModel
    Public Property foundations As EDSFoundationGroup
    Public Property connections As DataTransfererCCIplate
    Public Property pole As DataTransfererCCIpole

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub
#End Region
End Class






