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

    Public Sub SetDBStatus(Of T As EDSObjectWithQueries)(alist As List(Of T), prevList As List(Of T))

        'Copy previous list into a new list which you can delete from as needed
        Dim prevDeleteableList As New List(Of T)
        For Each prevItem In prevList
            prevDeleteableList.Add(prevItem)
        Next

        For Each item In alist
            If item.ID Is Nothing Then
                item.dbStatus = dbStatuses.Insert
                Continue For
            End If
            Dim IDMatch As Boolean = False
            For Each prevItem In prevDeleteableList
                If prevItem.ID = item.ID Then
                    IDMatch = True
                    If item.CompareMe(prevItem) Then
                        item.dbStatus = dbStatuses.NoChange
                    Else
                        item.dbStatus = dbStatuses.Update
                    End If
                    prevDeleteableList.Remove(prevItem)
                    Exit For
                End If
            Next
            If Not IDMatch Then item.dbStatus = dbStatuses.Insert
        Next

        For Each prevItem In prevDeleteableList
            'Adds items from the previous item list to the current item list with the dbstaus set to delete
            'I don't love this approach but I don't want to manage another list for deletions
            prevItem.dbStatus = dbStatuses.Delete
            alist.Add(prevItem)
        Next

    End Sub

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

Public Enum dbStatuses
    NoChange
    Insert
    Update
    Delete
End Enum

Partial Public MustInherit Class EDSObject

    Public Property ID As Integer?
    Public Property bus_unit As String
    Public Property structure_id As String
    Public Property work_order_seq_num As String
    Public Property activeDatabase As String
    Public Property databaseIdentity As String
    Public Property dbComparison As Comparison

    'Public MustOverride Function CompareMe() As Boolean


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

Partial Public MustInherit Class EDSObjectWithQueries
    Inherits EDSObject

    Public Property dbStatus As dbStatuses
    Public MustOverride ReadOnly Property Insert() As String
    Public MustOverride ReadOnly Property Update() As String
    Public MustOverride ReadOnly Property Delete() As String

    Public MustOverride Function SQLInsertUpdateDelete() As String

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

        LoadFromEDS(BU, structureID, LogOnUser, ActiveDatabase)
    End Sub
#End Region

#Region "EDS"
    Public Sub LoadFromEDS(ByVal BU As String, ByVal structureID As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)


        Dim query As String = QueryBuilderFromFile(queryPath & "Structure\Structure (SELECT).sql").Replace("[BU]", BU.ToDBString).Replace("[STRID]", strID.ToDBString)
        Dim tableNames() As String = {"TNX", "Base Structure", "Upper Structure", "Guys", "Members", "Materials", "Pier and Pad", "Unit Base", "Pile", "Drilled Pier", "Anchor Block", "Soil Profiles", "Soil Layers", "Connections", "Pole"}


        Using strDS As New DataSet

            sqlLoader(query, strDS, ActiveDatabase, LogOnUser, 500)

            'name tables from tableNames list
            For i = 0 To strDS.Tables.Count - 1
                strDS.Tables(i).TableName = tableNames(i)
            Next

            'Load TNX Model
            Me.tnx = New tnxModel(strDS)

            'Pier and Pad
            For Each dr As DataRow In strDS.Tables("Pier and Pad").Rows
                Me.PierandPads.Add(New PierAndPad(dr))
            Next

            'For additional tools we'll need to update the constructor to use a datarow and pass through the dataset byref for sub tables (i.e. soil profiles)
            'That constructor will grab datarows from the sub data tables based on the foreign key in datarow
            'For Each dr As DataRow In strDS.Tables("Drilled Pier").Rows
            '    Me.DrilledPiers.Add(New DrilledPier(dr, strDS))
            'Next

        End Using

    End Sub


    Public Sub SavetoEDS(ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

        If Me.ID IsNot Nothing Then Exit Sub 'If the ID has been specified, it's because it already exists and does not need to be updated

        'Dim fndGrpUpQuery As String = QueryBuilderFromFile(queryPath & "Foundation Group\Foundation Group (IN_UP) with Model.sql")

        'fndGrpUpQuery = fndGrpUpQuery.Replace("[BU NUMBER]", Me.BU.ToDBString)
        'fndGrpUpQuery = fndGrpUpQuery.Replace("[STRUCTURE ID]", Me.strID.ToDBString)
        'fndGrpUpQuery = fndGrpUpQuery.Replace("[FOUNDATION GROUP ID]", Me.ID.ToString.ToDBString)

        Dim structureQuery As String = ""
        'Dim fndSubQuery As String = QueryBuilderFromFile(queryPath & "Pier and Pad\Pier and Pad (INSERT).sql")
        For i = 0 To PierandPads.Count - 1
            structureQuery += PierandPads(i).InsertUpdateDelete(i)
        Next
        'fndSubQuery = QueryBuilderFromFile(queryPath & "Pile\Pile (IN_UP SUB).sql")
        'For i = 0 To Piles.Count - 1
        '    fndSubQueryList += Piles(i).GenerateGroupSubQuery(i)
        'Next
        'fndSubQuery = QueryBuilderFromFile(queryPath & "Unit Base\Unit Base (IN_UP SUB).sql")
        'For i = 0 To UnitBases.Count - 1
        '    fndSubQueryList += UnitBases(i).GenerateGroupSubQuery(i)
        'Next
        'fndSubQuery = QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Pier (IN_UP SUB).sql")
        'For i = 0 To DrilledPiers.Count - 1
        '    fndSubQueryList += DrilledPiers(i).GenerateGroupSubQuery(i)
        'Next
        'fndSubQuery = QueryBuilderFromFile(queryPath & "Guy Anchor Block\Guy Anchor Block (IN_UP SUB).sql")
        'For i = 0 To GuyAnchorBlocks.Count - 1
        '    fndSubQueryList += GuyAnchorBlocks(i).GenerateGroupSubQuery(i)
        'Next

        'fndGrpUpQuery = fndGrpUpQuery.Replace("[FOUNDATIONS]", fndSubQueryList)

        Using strcDS As New DataSet
            'Create the foundation group in EDS, return table of all foundation details added to group
            sqlLoader(structureQuery, "foundations", strcDS, ActiveDatabase, LogOnUser, 0.ToString)

            'If strcDS.Tables.Contains("foundations") Then
            '    'Assign new group ID
            '    Me.ID = DBtoNullableInt(strcDS.Tables("foundations").Rows(0).Item("FndGrpID"))
            '    'Go through table of foundation details and assign foundation IDs to foundations that were newly added to EDS
            '    For i = 0 To strcDS.Tables("foundations").Rows.Count - 1
            '        Dim fndIndex As Integer = CInt(strcDS.Tables("foundations").Rows(0).Item("FndIndex"))
            '        Select Case strcDS.Tables("foundations").Rows(i).Item("FndType").ToString
            '            Case "Pier and Pad"
            '                PierandPads(fndIndex).ID = DBtoNullableInt(strcDS.Tables("foundations").Rows(i).Item("fndid"))
            '                'Case "Pile"
            '                '    Piles(fndIndex).ID = DBtoNullableInt(fndDS.Tables("foundations").Rows(i).Item("fndid"))
            '                'Case "Unit Base"
            '                '    UnitBases(fndIndex).ID = DBtoNullableInt(fndDS.Tables("foundations").Rows(i).Item("fndid"))
            '                'Case "Drilled Pier"
            '                '    DrilledPiers(fndIndex).ID = DBtoNullableInt(fndDS.Tables("foundations").Rows(i).Item("fndid"))
            '                'Case "Guy Anchor Block"
            '                '    GuyAnchorBlocks(fndIndex).ID = DBtoNullableInt(fndDS.Tables("foundations").Rows(i).Item("fndid"))
            '        End Select
            '    Next
            'End If
        End Using

    End Sub
#End Region

#Region "Excel"
    Public Sub LoadFromFiles(filePaths As String())
        For Each item As String In filePaths
            If item.Contains("Pier and Pad Foundation") Then
                Me.PierandPads.Add(New PierAndPad(item, Me.bus_unit, Me.structure_id))
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

        Me.tnx.GenerateERI(Path.Combine(folderPath, Me.bus_unit, ".eri") 

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
    Public Function Compare(ByVal previous As EDSStructure) As Boolean


        ' Me.tnx.
        'Dim prevPPdelete As List(Of PierAndPad)
        'For Each prev In previous.PierandPads
        '    prevPPdelete.Add(prev)
        'Next
        'For Each pp In PierandPads
        '    If pp.ID Is Nothing Then
        '        pp.dbStatus = dbStatuses.Insert
        '        Continue For
        '    End If
        '    Dim IDMatch As Boolean = False
        '    For Each prevPP In previous.PierandPads
        '        If prevPP.ID = pp.ID Then
        '            IDMatch = True
        '            pp.CompareMe(prevPP)
        '            Exit For
        '        End If
        '    Next
        '    If Not IDMatch Then pp.dbStatus = dbStatuses.Insert
        'Next

        Me.PierandPads.SetDbStatus(previous.PierandPads)




    End Function
#End Region
End Class






