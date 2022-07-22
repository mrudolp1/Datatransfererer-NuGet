﻿Imports System.ComponentModel
Imports System.Security.Principal
Imports DevExpress.Spreadsheet
Imports System.IO
Imports DevExpress.DataAccess.Excel
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient

Partial Public Class EDSStructure
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String = "Structure Model"

    Public Property tnx As tnxModel
    Public Property connections As DataTransfererCCIplate
    Public Property pole As DataTransfererCCIpole
    Public Property structureCodeCriteria As SiteCodeCriteria
    Public Property PierandPads As New List(Of PierAndPad)
    Public Property Piles As New List(Of Pile)
    Public Property UnitBases As New List(Of UnitBase)
    'Public Property UnitBases As New List(Of SST_Unit_Base) 'Challs version - DNU
    Public Property DrilledPiers As New List(Of DrilledPier)
    Public Property GuyAnchorBlocks As New List(Of GuyedAnchorBlock)

    'The structure class should return itself if the parent is requested
    Private _ParentStructure As EDSStructure
    Public Overrides Property ParentStructure As EDSStructure
        Get
            Return Me
        End Get
        Set(value As EDSStructure)
            _ParentStructure = value
        End Set
    End Property

    Private Shared _SQLQueryVariables() As String = New String() {"@TopLevel", "@SubLevel1", "@SubLevel2", "@SubLevel3", "@SubLevel4"}

    Public Shared Function SQLQueryTableVar(depth As Integer) As String
        depth = Math.Max(0, depth)
        depth = Math.Min(_SQLQueryVariables.Length - 1, depth)
        Return _SQLQueryVariables(depth)
    End Function

    Public Shared Function SQLQueryIDVar(depth As Integer) As String
        Return SQLQueryTableVar(depth) & "ID"
    End Function

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

    Public Overrides Function ToString() As String
        Return Me.bus_unit & " - " & Me.structure_id
    End Function
#End Region

#Region "EDS"
    Public Sub LoadFromEDS(ByVal BU As String, ByVal structureID As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

        Dim query As String = QueryBuilderFromFile(queryPath & "Structure\Structure (SELECT).sql").Replace("[BU]", BU.FormatDBValue()).Replace("[STRID]", structureID.FormatDBValue())
        Dim tableNames() As String = {"TNX", "Base Structure", "Upper Structure", "Guys", "Members", "Materials", "Pier and Pad", "Unit Base", "Pile", "Drilled Pier", "Anchor Block", "Soil Profiles", "Soil Layers", "Connections", "Pole", "Site Code Criteria"}

        Using strDS As New DataSet

            sqlLoader(query, strDS, ActiveDatabase, LogOnUser, 500)

            'name tables from tableNames list
            For i = 0 To strDS.Tables.Count - 1
                strDS.Tables(i).TableName = tableNames(i)
            Next

            'If no site code criteria exists, fetch data from ORACLE to use for the first analysis. 
            'Still need to find all Topo inputs
            'Just set other parameters as default values 
            'If Not strDS.Tables("Site Code Criteria").Rows.Count > 0 Then
            '    OracleLoader("
            '        SELECT
            '                str.bus_unit
            '                ,str.structure_id
            '                ,tr.standard_code tia_current
            '                ,tr.bldg_code ibc_current
            '                ,str.ground_elev elev_agl
            '                ,str.hgt_no_appurt
            '                ,str.crest_height
            '                ,str.distance_from_crest
            '                ,sit.site_name
            '                ,'False' rev_h_section_15_5
            '                ,0 tower_point_elev
            '                --,pi.eng_app_id
            '                --,pi.crrnt_rvsn_num
            '            FROM
            '                isit_aim.structure                      str
            '                ,isit_aim.site                          sit
            '                ,rpt_appl.eng_tower_rating_vw           tr
            '                --,isit_aim.work_orders                 wo
            '                --,isit_isite.project_info              pi
            '            WHERE
            '                --wo.work_order_seqnum = 'XXXXXXX'
            '                str.bus_unit = '" & bus_unit & "' --Comment out when switching to WO
            '                AND str.structure_id = '" & structure_id & "' --Comment out when switching to WO
            '                AND str.bus_unit = sit.bus_unit
            '                AND str.bus_unit = tr.bus_unit
            '                --AND wo.bus_unit = str.bus_unit
            '                --AND wo.structure_id = str.structure_id
            '                --AND pi.eng_app_id = wo.eng_app_id(+)

            '        ", "Site Code Criteria", strDS, 3000, "ords")
            'End If
            'Me.structureCodeCriteria = New SiteCodeCriteria(strDS.Tables("Site Code Criteria").Rows(0))

            'Load TNX Model
            If strDS.Tables("TNX").Rows.Count > 0 Then
                Me.tnx = New tnxModel(strDS, Me)
            End If

            'Pier and Pad
            For Each dr As DataRow In strDS.Tables("Pier and Pad").Rows
                Me.PierandPads.Add(New PierAndPad(dr, Me))
            Next

            'Unit Base (CHall - DNU)
            'For Each dr As DataRow In strDS.Tables("Unit Base").Rows
            '    Me.UnitBases.Add(New SST_Unit_Base(dr, Me))
            'Next

            'Unit Base
            For Each dr As DataRow In strDS.Tables("Unit Base").Rows
                Me.UnitBases.Add(New UnitBase(dr, Me))
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
        For Each level In _SQLQueryVariables
            structureQuery += "DECLARE " & level & " TABLE(ID INT)" & vbCrLf
            structureQuery += "DECLARE " & level & "ID INT" & vbCrLf
        Next
        'Use the declared variables in the sub queries to pass along IDs that are needed as foreign keys.
        structureQuery += "BEGIN TRANSACTION" & vbCrLf

        structureQuery += Me.tnx?.EDSQueryBuilder(existingStructure.tnx)
        structureQuery += Me.PierandPads.EDSListQueryBuilder(existingStructure.PierandPads)
        structureQuery += Me.UnitBases.EDSListQueryBuilder(existingStructure.UnitBases)
        'structureQuery += Me.Piles.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.DrilledPiers.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.GuyAnchorBlocks.EDSListQuery(existingStructure.PierandPads)
        'structureQuery += Me.connections.EDSQuery(existingStructure.PierandPads)
        'structureQuery += Me.pole.EDSQuery(existingStructure.PierandPads)

        structureQuery += "COMMIT"

        'MessageBox.Show(structureQuery)

        sqlSender(structureQuery, ActiveDatabase, LogOnUser, 0.ToString)

    End Sub
#End Region

#Region "Files"
    Public Sub LoadFromFiles(filePaths As String())

        For Each item As String In filePaths
            If item.EndsWith(".eri") Then
                Me.tnx = New tnxModel(item, Me)
            ElseIf item.Contains("Pier and Pad Foundation") Then
                Me.PierandPads.Add(New PierAndPad(item, Me))
            ElseIf item.Contains("Pile Foundation") Then
                'Me.Piles.Add(New Pile(item))
            ElseIf item.Contains("SST Unit Base Foundation") Then
                'Me.UnitBases.Add(New SST_Unit_Base(item, Me)) 'Chall version - DNU
                Me.UnitBases.Add(New UnitBase(item, Me))
            ElseIf item.Contains("Drilled Pier Foundation") Then
                'Me.DrilledPiers.Add(New DrilledPier(item))
            ElseIf item.Contains("Guyed Anchor Block Foundation") Then
                'Me.GuyAnchorBlocks.Add(New GuyedAnchorBlock(item))
            End If
        Next
    End Sub

    Public Sub SaveTools(folderPath As String)
        'Uncomment your foundation type for testing when it's ready.
        Dim i As Integer
        Dim fileNum As String = ""

        If Me.tnx IsNot Nothing Then Me.tnx.GenerateERI(Path.Combine(folderPath, Me.bus_unit & ".eri"))

        For i = 0 To Me.PierandPads.Count - 1
            'I think we need a better way to get filename and maintain meaningful file names after they've gone through the database.
            'This works for now, just basing the name off the template name.
            fileNum = String.Format(" ({0})", i.ToString)
            PierandPads(i).workBookPath = Path.Combine(folderPath, Me.bus_unit & "_" & Path.GetFileNameWithoutExtension(PierandPads(i).templatePath) & "_EDS_" & fileNum & Path.GetExtension(PierandPads(i).templatePath))
            PierandPads(i).SavetoExcel()
        Next
        'For i = 0 To Me.Piles.Count - 1
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    Piles(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(Piles(i).templatePath) & fileNum)
        '    Piles(i).SavetoExcel()
        'Next
        For i = 0 To Me.UnitBases.Count - 1
            fileNum = String.Format(" ({0})", i.ToString)
            UnitBases(i).workBookPath = Path.Combine(folderPath, Me.bus_unit & "_" & Path.GetFileNameWithoutExtension(UnitBases(i).templatePath) & "_EDS_" & fileNum & Path.GetExtension(UnitBases(i).templatePath))
            UnitBases(i).SavetoExcel()
        Next
        'For i = 0 To Me.DrilledPiers.Count - 1
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    DrilledPiers(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(DrilledPiers(i).templatePath) & fileNum)
        '    DrilledPiers(i).SavetoExcel()
        'Next
        'For i = 0 To Me.GuyAnchorBlocks.Count - 1
        '    fileNum = If(i = 0, "", Format(" ({0})", i.ToString))
        '    GuyAnchorBlocks(i).workBookPath = Path.Combine(folderPath, Path.GetFileName(GuyAnchorBlocks(i).templatePath) & fileNum)
        '    GuyAnchorBlocks(i).SavetoExcel()
        'Next
    End Sub
#End Region

#Region "Check Changes"
    Public Function CompareEDS(other As EDSObject, Optional ByRef changes As List(Of AnalysisChange) = Nothing) As Boolean

    End Function

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As EDSStructure = TryCast(other, EDSStructure)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.tnx.CheckChange(otherToCompare.tnx, changes, categoryName, "TNX"), Equals, False)
        Equals = If(Me.connections.CheckChange(otherToCompare.connections, changes, categoryName, "Connections"), Equals, False)
        Equals = If(Me.pole.CheckChange(otherToCompare.pole, changes, categoryName, "Pole"), Equals, False)
        Equals = If(Me.structureCodeCriteria.CheckChange(otherToCompare.structureCodeCriteria, changes, categoryName, "Structure Code Criteria"), Equals, False)
        Equals = If(Me.PierandPads.CheckChange(otherToCompare.PierandPads, changes, categoryName, "Pier and Pads"), Equals, False)
        Equals = If(Me.Piles.CheckChange(otherToCompare.Piles, changes, categoryName, "Piles"), Equals, False)
        Equals = If(Me.UnitBases.CheckChange(otherToCompare.UnitBases, changes, categoryName, "Unit Bases"), Equals, False)
        Equals = If(Me.DrilledPiers.CheckChange(otherToCompare.DrilledPiers, changes, categoryName, "Drilled Piers"), Equals, False)
        Equals = If(Me.GuyAnchorBlocks.CheckChange(otherToCompare.GuyAnchorBlocks, changes, categoryName, "Guy Anchor Blocks"), Equals, False)
        Return Equals
    End Function

#End Region
End Class









