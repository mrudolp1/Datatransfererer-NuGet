Imports System.ComponentModel
Imports System.Security.Principal
Imports DevExpress.Spreadsheet
Imports System.IO
Imports DevExpress.DataAccess.Excel
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient

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
    Public Function EDSListQuery(Of T As EDSObjectWithQueries)(alist As List(Of T), prevList As List(Of T)) As String

        EDSListQuery = ""

        'Copy previous list into a new list which you can delete from as needed
        Dim prevDeleteableList As New List(Of T)
        For Each prevItem In prevList
            prevDeleteableList.Add(prevItem)
        Next

        For Each item In alist
            Dim IDMatch As Boolean = False
            For Each prevItem In prevDeleteableList
                If prevItem.ID = item.ID Then
                    IDMatch = True
                    If Not item.CompareMe(prevItem) Then
                        EDSListQuery += item.Update
                    End If
                    prevDeleteableList.Remove(prevItem)
                    Exit For
                End If
            Next
            If Not IDMatch Then
                'Need to add inserted items to comparison list.
                EDSListQuery += item.Insert
            End If
        Next

        For Each prevItem In prevDeleteableList
            'Adds items from the previous item list to the current item list with the dbstaus set to delete
            'I don't love this approach but I don't want to manage another list for deletions
            EDSListQuery += prevItem.Delete
        Next

        Return EDSListQuery

    End Function

    <Extension()>
    Public Iterator Function Add(Of T)(ByVal e As IEnumerable(Of T), ByVal value As T) As IEnumerable(Of T)
        'Allow you to add to an IEnumerable like it is a list.
        'Useful for working with the ObjectComparer class which stores the differences as IEnumerable(of Difference)
        'Refernce: https://stackoverflow.com/a/1210311
        For Each cur In e
            Yield cur
        Next

        Yield value
    End Function

    '<Extension()>
    'Public Function CompareEDSLists(Of T As EDSObject)(List1 As List(Of T), List2 As List(Of T), Optional SetID As Boolean = False) As Boolean
    '    'Compare lists of EDSObjects that are not in the same order.
    '    'If SetID then items in list1 will be updated with the IDs of matching items in list2.
    '    'Typical use case list1 would be excel tools from the current analysis, list2 would be the tools on EDS.
    '    CompareEDSLists = True

    '    If List1.Count <> List2.Count Then
    '        CompareEDSLists = False
    '        'If setting IDs from list2, we need to keep going even tho we no the overall lists are different
    '        If Not SetID Then Return CompareEDSLists
    '    End If

    '    For Each item In List1
    '        Dim CompareItem As Boolean
    '        For Each item2 In List2
    '            CompareItem = item.CompareMe(item2, SetID)
    '            If CompareItem Then Exit For
    '        Next
    '        If Not CompareItem Then
    '            CompareEDSLists = False
    '            If Not SetID Then Exit For
    '        End If
    '    Next

    '    Return CompareEDSLists
    'End Function

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

Partial Public MustInherit Class EDSObject

    Public Property ID As Integer?
    Public Property bus_unit As String
    Public Property structure_id As String
    'Public Property work_order_seq_num As String
    Public Property activeDatabase As String
    Public Property databaseIdentity As WindowsIdentity
    Public Property differences As List(Of ObjectsComparer.Difference)

    Public Overridable Function CreateChangeSummary() As String
        Dim summary As String = ""

        For Each chng As AnalysisChanges In changeList
            summary += chng.CategoryName & " " & chng.FieldName & " = " & chng.NewValue & " | Previously: " & chng.PreviousValue & vbNewLine
        Next

        Return summary

    End Function

    Public Overridable Function CompareMe(Of T As EDSObject)(toCompare As T) As Boolean
        'Compare another EDSObject object to itself using the objects comparer.
        'Making this generic (Of T As EDSObject) allows it to work for all subclasses of EDSObject

        If toCompare Is Nothing Or Me.GetType() IsNot toCompare.GetType() Then Return False

        Dim comparer As New ObjectsComparer.Comparer(Of T)()

        Dim differences As IEnumerable(Of ObjectsComparer.Difference) = Nothing

        CompareMe = comparer.Compare(CType(Me, T), toCompare, differences)

        Me.differences = differences.ToList

        Return CompareMe

    End Function

    Public Overridable Sub Absorb(ByRef Parent As EDSObject)
        Me.bus_unit = Parent.bus_unit
        Me.structure_id = Parent.structure_id
        'Me.work_order_seq_num = Parent.work_order_seq_num
        Me.activeDatabase = Parent.activeDatabase
        Me.databaseIdentity = Parent.databaseIdentity
    End Sub

End Class

Partial Public MustInherit Class EDSObjectWithQueries
    Inherits EDSObject

    Public MustOverride ReadOnly Property Insert() As String
    Public MustOverride ReadOnly Property Update() As String
    Public MustOverride ReadOnly Property Delete() As String

    'Public MustOverride Function SQLInsertUpdateDelete() As String

    Public MustOverride Function SQLInsertValues() As String

    Public MustOverride Function SQLInsertFields() As String

    Public MustOverride Function SQLUpdate() As String



End Class

Partial Public MustInherit Class EDSExcelObject
    'This should be inherited by the main tool class. Subclasses such as soil layers can probably inherit the EDSObjectWithQueries
    Inherits EDSObjectWithQueries

    Public Property workBookPath As String
    Public MustOverride ReadOnly Property templatePath As String
    Public Property fileType As DocumentFormat = DocumentFormat.Xlsm
    Public MustOverride ReadOnly Property EDSTableName As String
    Public MustOverride ReadOnly Property excelDTParams As List(Of EXCELDTParameter)

#Region "Save to Excel"
    Public MustOverride Sub workBookFiller(ByRef wb As Workbook)

    Public Sub SavetoExcel()
        Dim wb As New Workbook

        If workBookPath = "" Then
            Debug.Print("No workbook path specified.")
            Exit Sub
        End If

        Try
            wb.LoadDocument(templatePath, fileType)
            wb.BeginUpdate()

            'Put the jelly in the donut
            workBookFiller(wb)

            wb.Calculate()
            wb.EndUpdate()
            wb.SaveDocument(workBookPath, fileType)
        Catch ex As Exception
            Debug.Print("Error Saving Workbook: " & ex.Message)
        End Try

    End Sub
#End Region

End Class

Partial Public MustInherit Class EDSFoundation
    Inherits EDSExcelObject

    Public MustOverride ReadOnly Property foundationType As String

End Class

Partial Public Class EDSStructure
    Inherits EDSObject

    Public Property tnx As tnxModel
    'Public Property foundations As EDSFoundationGroup
    Public Property PierandPads As New List(Of PierAndPad)
    Public Property Piles As New List(Of Pile)
    Public Property UnitBases As New List(Of SST_Unit_Base)
    Public Property DrilledPiers As New List(Of DrilledPier)
    Public Property GuyAnchorBlocks As New List(Of GuyedAnchorBlock)
    Public Property connections As DataTransfererCCIplate
    Public Property pole As DataTransfererCCIpole

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal BU As String, ByVal structureID As String, filePaths As String())
        Me.bus_unit = BU
        Me.structure_id = structureID
        'Uncomment your foundation type for testing when it's ready. 
        LoadFromFiles(filePaths)
    End Sub
    Public Sub New(ByVal BU As String, ByVal structureID As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Me.bus_unit = BU
        Me.structure_id = structureID
        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase

        LoadFromEDS(BU, structureID, LogOnUser, ActiveDatabase)
    End Sub
#End Region

#Region "EDS"
    Public Sub LoadFromEDS(ByVal BU As String, ByVal structureID As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

        Dim query As String = QueryBuilderFromFile(queryPath & "Structure\Structure (SELECT).sql").Replace("[BU]", BU.ToDBString).Replace("[STRID]", structureID.ToDBString)
        Dim tableNames() As String = {"TNX", "Base Structure", "Upper Structure", "Guys", "Members", "Materials", "Pier and Pad", "Unit Base", "Pile", "Drilled Pier", "Anchor Block", "Soil Profiles", "Soil Layers", "Connections", "Pole"}


        Using strDS As New DataSet

            sqlLoader(query, strDS, ActiveDatabase, LogOnUser, 500)

            'name tables from tableNames list
            For i = 0 To strDS.Tables.Count - 1
                strDS.Tables(i).TableName = tableNames(i)
            Next

            'Load TNX Model
            Me.tnx = New tnxModel(strDS, Me)

            'Pier and Pad
            For Each dr As DataRow In strDS.Tables("Pier and Pad").Rows
                Me.PierandPads.Add(New PierAndPad(dr, Me))
            Next

            'For additional tools we'll need to update the constructor to use a datarow and pass through the dataset byref for sub tables (i.e. soil profiles)
            'That constructor will grab datarows from the sub data tables based on the foreign key in datarow
            'For Each dr As DataRow In strDS.Tables("Drilled Pier").Rows
            '    Me.DrilledPiers.Add(New DrilledPier(dr, strDS))
            'Next

        End Using

    End Sub


    Public Sub SavetoEDS(ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase

        Dim existingStructure As New EDSStructure(Me.bus_unit, Me.structure_id, Me.databaseIdentity, Me.activeDatabase)

        Dim structureQuery As String = ""
        'structureQuery += Me.tnx.EDSQuery(existingStructure.tnx)
        structureQuery += Me.PierandPads.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.UnitBases.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.Piles.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.DrilledPiers.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.GuyAnchorBlocks.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.connections.EDSQuery(existingStructure.PierandPads)
        'structureQuery += Me.pole.EDSQuery(existingStructure.PierandPads)

        MessageBox.Show(structureQuery)

        sqlSender(structureQuery, ActiveDatabase, LogOnUser, 0.ToString)

    End Sub
#End Region

#Region "Excel"
    Public Sub LoadFromFiles(filePaths As String())

        For Each item As String In filePaths
            If item.EndsWith(".eri") Then
                Me.tnx = New tnxModel(item)
            ElseIf item.Contains("Pier and Pad Foundation") Then
                Me.PierandPads.Add(New PierAndPad(item, Me))
            ElseIf item.Contains("Pile Foundation") Then
                'Me.Piles.Add(New Pile(item))
            ElseIf item.Contains("SST Unit Base Foundation") Then
                'Me.UnitBases.Add(New UnitBase(item))
            ElseIf item.Contains("Drilled Pier Foundation") Then
                'Me.DrilledPiers.Add(New DrilledPier(item))
            ElseIf item.Contains("Guyed Anchor Block Foundation") Then
                'Me.GuyAnchorBlocks.Add(New GuyedAnchorBlock(item))
            End If
        Next
    End Sub

    Public Sub SaveToolstoExcel(folderPath As String)
        'Uncomment your foundation type for testing when it's ready.
        Dim i As Integer
        Dim fileNum As String = ""

        Me.tnx.GenerateERI(Path.Combine(folderPath, Me.bus_unit, ".eri"))

        For i = 0 To Me.PierandPads.Count
            'I think we need a better way to get filename and maintain meaningful file names after they've gone through the database.
            'This works for now, just basing the name off the template name.
            fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
            PierandPads(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(PierandPads(i).templatePath) & fileNum)
            PierandPads(i).SavetoExcel()
        Next
        'For i = 0 To Me.Piles.Count
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    Piles(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(Piles(i).templatePath) & fileNum)
        '    Piles(i).SavetoExcel()
        'Next
        'For i = 0 To Me.UnitBases.Count
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    UnitBases(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(UnitBases(i).templatePath) & fileNum)
        '    UnitBases(i).SavetoExcel()
        'Next
        'For i = 0 To Me.DrilledPiers.Count
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    DrilledPiers(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(DrilledPiers(i).templatePath) & fileNum)
        '    DrilledPiers(i).SavetoExcel()
        'Next
        'For i = 0 To Me.GuyAnchorBlocks.Count
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    GuyAnchorBlocks(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(GuyAnchorBlocks(i).templatePath) & fileNum)
        '    GuyAnchorBlocks(i).SavetoExcel()
        'Next
    End Sub
#End Region

#Region "Check Changes"

#End Region
End Class






