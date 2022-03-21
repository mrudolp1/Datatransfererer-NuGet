Option Strict On

Imports System.ComponentModel
Imports System.Security.Principal
Imports DevExpress.Spreadsheet
Imports System.IO
Imports DevExpress.DataAccess.Excel
Imports System.Runtime.CompilerServices

Partial Public MustInherit Class EDSFoundation
    'Should align with an EDS (general) Foundation Detail and should be inherited by the class that represent the specific foundation details
    Inherits EDSObject

    Public Property workBookPath As String
    Public MustOverride ReadOnly Property templatePath As String
    Public Property fileType As DocumentFormat = DocumentFormat.Xlsm
    Public MustOverride ReadOnly Property foundationType As String
    Public MustOverride ReadOnly Property EDSTableName As String
    Public Property fndDetailID As Integer?
    Public Property guyGroupID As Integer?
    Public MustOverride ReadOnly Property excelDTParams As List(Of EXCELDTParameter)

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    'Public Sub New(ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String, ByVal BU As String, ByVal Strucutre_ID As String)
    '    Dim foundationData As New DataSet

    'End Sub
#End Region

#Region "Save to EDS"
    Public MustOverride Function PopulateSubQuery(blankQuery As String, index As Integer) As String

    Public MustOverride Function GenerateSQLValues() As String

#End Region
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

Partial Public Class EDSFoundationGroup
    'This represents all the foundations associated with a structure.
    'This class will be responsible for uploading and downloading all the foundations from EDS.
    Inherits EDSObject

    Public Property PierandPads As New List(Of PierAndPad)
    Public Property Piles As New List(Of Pile)
    Public Property UnitBases As New List(Of SST_Unit_Base)
    Public Property DrilledPiers As New List(Of DrilledPier)
    Public Property GuyAnchorBlocks As New List(Of GuyedAnchorBlock)

    Public Function CountAll() As Integer
        Return PierandPads.Count + Piles.Count + UnitBases.Count + DrilledPiers.Count + GuyAnchorBlocks.Count
    End Function


#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal BU As String, ByVal strID As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Using fndGrpDS As New DataSet
            sqlLoader(QueryBuilderFromFile(queryPath & "Foundation Group\Foundation Group (SELECT ID Active).sql").Replace("[BU]", BU).Replace("[STRC ID]", strID.ToDBString), "fndGrp", fndGrpDS, ActiveDatabase, LogOnUser, "0")
            If fndGrpDS.Tables.Contains("fndGrp") Then
                LoadAllFoundationsFromEDS(DBtoStr(fndGrpDS.Tables("fndGrp").Rows(0).Item("foundation_group_id")), LogOnUser, ActiveDatabase)
            End If
        End Using
    End Sub

    Public Sub New(ByVal fndGrpID As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

        LoadAllFoundationsFromEDS(fndGrpID, LogOnUser, ActiveDatabase)

    End Sub

    Public Sub New(ByVal BU As String, ByVal strID As String, filePaths As String())
        Me.BU = BU
        Me.strID = strID
        'Uncomment your foundation type for testing when it's ready. 
        For Each item As String In filePaths
            If item.Contains("Pier and Pad Foundation") Then
                Me.PierandPads.Add(New PierAndPad(item))
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
#End Region

    Public Sub LoadAllFoundationsFromEDS(fndGrpID As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

        Dim tableNames As New List(Of String) From {"Pier and Pad", "Unit Base", "Pile"} 'Add "Drilled Pier", "Guy Anchor"
        Dim queries As New List(Of String)

        queries.Add(QueryBuilderFromFile(queryPath & "Pier and Pad\Pier and Pad (SELECT Details Grp).sql").Replace("[FNDGRPID]", fndGrpID))
        'queries.Add(QueryBuilderFromFile(queryPath & "Unit Base\Unit Base (SELECT Details Grp).sql").Replace("[FNDGRPID]", fndGrpID))
        'queries.Add(QueryBuilderFromFile(queryPath & "Pile\Pile (SELECT Details Grp).sql").Replace("[FNDGRPID]", fndGrpID))
        'queries.Add(QueryBuilderFromFile(queryPath & "Drilled Pier\Drilled Pier (SELECT Details Grp).sql").Replace("[FNDGRPID]", fndGrpID))
        'queries.Add(QueryBuilderFromFile(queryPath & "Guy Anchor Block\Guy Anchor Block (SELECT Details Grp).sql").Replace("[FNDGRPID]", fndGrpID))

        Using fndDS As New DataSet

            sqlLoader(queries, tableNames, fndDS, ActiveDatabase, LogOnUser, 500)

            If fndDS.Tables.Contains("Pier and Pad") Then
                For Each dr As DataRow In fndDS.Tables("Pier and Pad").Rows
                    Me.PierandPads.Add(New PierAndPad(dr))
                Next
            End If

            'If fndDS.Tables.Contains("Unit Base") Then
            '    For Each fnd As DataRow In fndDS.Tables("Unit Base").Rows
            '        Me.foundations.Add(New UnitBase(fnd))
            '    Next
            'End If

            'If fndDS.Tables.Contains("Pile") Then
            '    For Each fnd As DataRow In fndDS.Tables("Pile").Rows
            '        Me.foundations.Add(New Pile(fnd))
            '    Next
            'End If
        End Using

    End Sub



#Region "Save Data"
    Public Sub SaveAllFoundationstoExcel(folderPath As String)
        'Uncomment your foundation type for testing when it's ready.
        Dim i As Integer
        Dim fileNum As String = ""
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

    Public Sub SaveAllFoundationsEDS(ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

        If Me.ID IsNot Nothing Then Exit Sub 'If the ID has been specified, it's because it already exists and does not need to be updated

        Dim fndGrpUpQuery As String = QueryBuilderFromFile(queryPath & "Foundation Group\Foundation Group (IN_UP) with Model.sql")

        fndGrpUpQuery = fndGrpUpQuery.Replace("[BU NUMBER]", Me.BU.ToDBString)
        fndGrpUpQuery = fndGrpUpQuery.Replace("[STRUCTURE ID]", Me.strID.ToDBString)
        fndGrpUpQuery = fndGrpUpQuery.Replace("[FOUNDATION GROUP ID]", Me.ID.ToString.ToDBString)

        Dim fndSubQueryList As String = ""
        Dim fndSubQuery As String = QueryBuilderFromFile(queryPath & "Pier and Pad\Pier and Pad (IN_UP SUB).sql")
        For i = 0 To PierandPads.Count - 1
            fndSubQueryList += PierandPads(i).PopulateSubQuery(fndSubQuery, i)
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

        fndGrpUpQuery = fndGrpUpQuery.Replace("[FOUNDATIONS]", fndSubQueryList)

        Using fndDS As New DataSet
            'Create the foundation group in EDS, return table of all foundation details added to group
            sqlLoader(fndGrpUpQuery, "foundations", fndDS, ActiveDatabase, LogOnUser, 0.ToString)

            If fndDS.Tables.Contains("foundations") Then
                'Assign new group ID
                Me.ID = DBtoNullableInt(fndDS.Tables("foundations").Rows(0).Item("FndGrpID"))
                'Go through table of foundation details and assign foundation IDs to foundations that were newly added to EDS
                For i = 0 To fndDS.Tables("foundations").Rows.Count - 1
                    Dim fndIndex As Integer = CInt(fndDS.Tables("foundations").Rows(0).Item("FndIndex"))
                    Select Case fndDS.Tables("foundations").Rows(i).Item("FndType").ToString
                        Case "Pier and Pad"
                            PierandPads(fndIndex).ID = DBtoNullableInt(fndDS.Tables("foundations").Rows(i).Item("fndid"))
                            'Case "Pile"
                            '    Piles(fndIndex).ID = DBtoNullableInt(fndDS.Tables("foundations").Rows(i).Item("fndid"))
                            'Case "Unit Base"
                            '    UnitBases(fndIndex).ID = DBtoNullableInt(fndDS.Tables("foundations").Rows(i).Item("fndid"))
                            'Case "Drilled Pier"
                            '    DrilledPiers(fndIndex).ID = DBtoNullableInt(fndDS.Tables("foundations").Rows(i).Item("fndid"))
                            'Case "Guy Anchor Block"
                            '    GuyAnchorBlocks(fndIndex).ID = DBtoNullableInt(fndDS.Tables("foundations").Rows(i).Item("fndid"))
                    End Select
                Next
            End If
        End Using

    End Sub
#End Region

#Region "Compare"
    Public Overloads Function CompareMe(ByRef currentFndGrp As EDSFoundationGroup, Optional SetID As Boolean = False) As Boolean
        'Compare this object to another instance, return true if they're the same
        'If SetID is set to true, IDs will be copied from the comparrison object to this object for both the foundaiton group and all foundations

        CompareMe = PierandPads.CompareEDSLists(currentFndGrp.PierandPads, SetID) 'And
        '            Piles.CompareEDSLists(currentFndGrp.Piles, SetID) And
        '            UnitBases.CompareEDSLists(currentFndGrp.UnitBases, SetID) And
        '            DrilledPiers.CompareEDSLists(currentFndGrp.DrilledPiers, SetID) And
        '            GuyAnchorBlocks.CompareEDSLists(currentFndGrp.GuyAnchorBlocks, SetID)

        If CompareMe And SetID Then
            Me.ID = currentFndGrp.ID
        End If

        Return CompareMe
    End Function

#End Region
End Class

