Imports System.ComponentModel
Imports System.Security.Principal
Imports DevExpress.Spreadsheet
Imports System.IO
Imports System.Reflection
Imports DevExpress.DataAccess.Excel
Imports System.Runtime.CompilerServices
Imports System.Data.SqlClient
Imports System.Runtime.Serialization
Imports DevExpress.Data.Helpers.SyncHelper.ZombieContextsDetector
Imports DevExpress.DataProcessing.InMemoryDataProcessor

<Serializable()>
<TypeConverterAttribute(GetType(ExpandableObjectConverter))>
<DataContract()>
Partial Public Class EDSStructure
    Inherits EDSObject

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Structure Model"
        End Get
    End Property


    'The structure class should return itself if the parent is requested
    Private _ParentStructure As EDSStructure
    Public Overrides ReadOnly Property ParentStructure As EDSStructure
        Get
            Return Me
        End Get
        'Set(value As EDSStructure)
        '    _ParentStructure = value
        'End Set
    End Property

    <DataMember()> Public Property tnx As tnxModel
    <DataMember()> Public Property CCIplates As New List(Of CCIplate)
    <DataMember()> Public Property structureCodeCriteria As SiteCodeCriteria
    <DataMember()> Public Property PierandPads As New List(Of PierAndPad)
    <DataMember()> Public Property Piles As New List(Of Pile)
    <DataMember()> Public Property Poles As New List(Of Pole)
    <DataMember()> Public Property UnitBases As New List(Of UnitBase)
    <DataMember()> Public Property DrilledPierTools As New List(Of DrilledPierFoundation)
    <DataMember()> Public Property GuyAnchorBlockTools As New List(Of AnchorBlockFoundation)
    <DataMember()> Public Property ReportOptions As ReportOptions
    <DataMember()> Public Property SiteInfo As SiteInfo
    <DataMember()> Public Property EDSMe As EDSStructure
    <DataMember()> Public Property WorkingDirectory As String
    <DataMember()> Public Property LegReinforcements As New List(Of LegReinforcement)
    <DataMember()> Public Property CCISeismics As New List(Of CCISeismic)

    Public Event MessageLogged As EventHandler(Of MessageLoggedEventArgs)

    Protected Overridable Sub OnMessageLogged(logMessage As LogMessage)
        If MessageLoggedEvent IsNot Nothing Then ''If the event has a handler this will be something.
            RaiseEvent MessageLogged(Me, New MessageLoggedEventArgs(logMessage))
        End If
    End Sub

    Public Delegate Sub MaestroProgressHandler(logMessage As LogMessage)

    Public Overrides Sub Clear()
        Me.CCIplates.Clear()
        Me.PierandPads.Clear()
        Me.Piles.Clear()
        Me.Poles.Clear()
        Me.UnitBases.Clear()
        Me.DrilledPierTools.Clear()
        Me.GuyAnchorBlockTools.Clear()
        Me.ReportOptions = Nothing
        Me.SiteInfo = Nothing
        Me.EDSMe = Nothing
        Me.WorkingDirectory = ""
        Me.LegReinforcements.Clear()
        Me.CCISeismics.Clear()
        Me.tnx.Clear()
    End Sub

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

    Public Sub New(ByVal BU As String, ByVal structureID As String, ByVal WorkOrder As String, ByVal workDirectory As String, filePaths As String(), ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Me.bus_unit = BU
        Me.structure_id = structureID
        Me.work_order_seq_num = WorkOrder
        Me.WorkingDirectory = workDirectory

        LoadFromFiles(filePaths)
    End Sub

    Public Sub New(ByVal BU As String, ByVal structureID As String, ByVal WorkOrder As String, ByVal workDirectory As String, ByVal reportDirectory As String, filePaths As String(), ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Me.bus_unit = BU
        Me.structure_id = structureID
        Me.work_order_seq_num = WorkOrder
        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase
        Me.WorkingDirectory = workDirectory
        Me.ReportOptions = New ReportOptions(workDirectory, reportDirectory, Me)
        Me.SiteInfo = New SiteInfo(WorkOrder)


        LoadFromFiles(filePaths)
    End Sub

    Public Sub New(ByVal BU As String, ByVal structureID As String, ByVal WorkOrder As String, ByVal Order As String, ByVal OrderRev As String, ByVal workDirectory As String, ByVal reportDirectory As String, filePaths As String(), ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Me.bus_unit = BU
        Me.structure_id = structureID
        Me.work_order_seq_num = WorkOrder
        Me.order = Order
        Me.orderRev = OrderRev
        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase
        Me.WorkingDirectory = workDirectory
        Me.ReportOptions = New ReportOptions(workDirectory, reportDirectory, Me)
        Me.SiteInfo = New SiteInfo(WorkOrder)

        LoadFromFiles(filePaths)
    End Sub

    Public Sub New(ByVal BU As String, ByVal structureID As String, ByVal WorkOrder As String,
                   ByVal Order As String, ByVal OrderRev As String,
                   ByVal EDSPersonID As Integer, ByVal stage As String,
                   ByVal workDirectory As String, ByVal reportDirectory As String,
                   ByVal filePaths As String(), ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Me.bus_unit = BU
        Me.structure_id = structureID
        Me.work_order_seq_num = WorkOrder
        Me.order = Order
        Me.orderRev = OrderRev
        Me.modified_person_id = EDSPersonID
        Me.process_stage = stage
        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase
        Me.WorkingDirectory = workDirectory
        Me.ReportOptions = New ReportOptions(workDirectory, reportDirectory, Me)
        Me.SiteInfo = New SiteInfo(WorkOrder)

        LoadFromFiles(filePaths)
    End Sub

    Public Sub New(ByVal BU As String, ByVal structureID As String, ByVal WorkOrder As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Me.bus_unit = BU
        Me.structure_id = structureID
        Me.work_order_seq_num = WorkOrder
        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase
        Me.ReportOptions = New ReportOptions()
        Me.ReportOptions.Absorb(Me)

        LoadFromEDS(BU, structureID, LogOnUser, ActiveDatabase)
    End Sub

    Public Sub New(ByVal BU As String, ByVal structureID As String, ByVal WorkOrder As String, ByVal Order As String, ByVal OrderRev As String, ByVal workDirectory As String, ByVal reportDirectory As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Me.bus_unit = BU
        Me.structure_id = structureID
        Me.work_order_seq_num = WorkOrder
        Me.order = Order
        Me.orderRev = OrderRev
        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase
        Me.WorkingDirectory = workDirectory
        Me.ReportOptions = New ReportOptions(workDirectory, reportDirectory, Me)
        Me.SiteInfo = New SiteInfo(WorkOrder)

        LoadFromEDS(BU, structureID, LogOnUser, ActiveDatabase)
    End Sub

    Public Sub New(ByVal BU As String, ByVal structureID As String, ByVal WorkOrder As String,
                   ByVal Order As String, ByVal OrderRev As String, ByVal EDSPersonID As Int32?, ByVal Stage As String,
                   ByVal workDirectory As String, ByVal reportDirectory As String,
                   ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Me.bus_unit = BU
        Me.structure_id = structureID
        Me.work_order_seq_num = WorkOrder
        Me.order = Order
        Me.orderRev = OrderRev
        Me.modified_person_id = EDSPersonID
        Me.process_stage = Stage
        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase
        Me.WorkingDirectory = workDirectory
        Me.ReportOptions = New ReportOptions(workDirectory, reportDirectory, Me)
        Me.SiteInfo = New SiteInfo(WorkOrder)

        LoadFromEDS(BU, structureID, LogOnUser, ActiveDatabase)
    End Sub

    Public Sub New(ByVal BU As String, ByVal structureID As String, ByVal WorkOrder As String, ByVal workDirectory As String, ByVal reportDirectory As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)
        Me.bus_unit = BU
        Me.structure_id = structureID
        Me.work_order_seq_num = WorkOrder
        Me.databaseIdentity = LogOnUser
        Me.activeDatabase = ActiveDatabase
        Me.WorkingDirectory = workDirectory
        Me.ReportOptions = New ReportOptions(workDirectory, reportDirectory, Me)
        Me.SiteInfo = New SiteInfo(WorkOrder)

        LoadFromEDS(BU, structureID, LogOnUser, ActiveDatabase)
    End Sub

    Public Sub CloneSettings(ByVal strctr As EDSStructure)
        Me.bus_unit = strctr.bus_unit
        Me.structure_id = strctr.structure_id
        Me.work_order_seq_num = strctr.work_order_seq_num
        Me.order = strctr.order
        Me.orderRev = strctr.orderRev
        Me.modified_person_id = strctr.modified_person_id
        Me.process_stage = strctr.process_stage
        Me.databaseIdentity = strctr.databaseIdentity
        Me.activeDatabase = strctr.activeDatabase
        Me.WorkingDirectory = strctr.WorkingDirectory
        Me.ReportOptions = strctr.ReportOptions
        Me.SiteInfo = strctr.SiteInfo
    End Sub

    Public Overrides Function ToString() As String
        Return String.Format("WO: {0}, BU: {1}, Structure: {2}", Me.EDSObjectName, Me.work_order_seq_num, Me.bus_unit, Me.structure_id)
    End Function
#End Region

#Region "EDS"
    Public Sub LoadFromEDS(ByVal BU As String, ByVal structureID As String, ByVal LogOnUser As WindowsIdentity, ByVal ActiveDatabase As String)

        ''Dim query As String = QueryBuilderFromFile(queryPath & "Structure\Structure (SELECT).sql").Replace("[BU]", BU.FormatDBValue()).Replace("[STRID]", structureID.FormatDBValue())
        Dim query As String = CCI_Engineering_Templates.My.Resources.Structure_SELECT.Replace("[BU]", BU.FormatDBValue()).Replace("[STRID]", structureID.FormatDBValue())
        Dim tableNames() As String = {"TNX",
                            "Base Structure",
                            "Upper Structure",
                            "Guys",
                            "Members",
                            "Materials",
                            "Pier and Pad",
                            "Unit Base",
                            "Pile",
                            "Pile Locations", '"Drilled Pier",'"Anchor Block",
                            "Soil Profiles",
                            "Soil Layers",
                            "CCIplates",
                            "Connections",
                            "Plate Details",
                            "Bolt Groups",
                            "Bolt Details",
                            "CCIplate Materials",
                            "Stiffener Groups",
                            "Stiffener Details",
                            "Bridge Stiffener Details",
                            "Pole General",
                            "Pole Unreinforced Sections",
                            "Pole Reinforced Sections",
                            "Pole Reinforcement Groups",
                            "Pole Reinforcement Details",
                            "Pole Interference Groups",
                            "Pole Interference Details", ' "Pole Results",
                            "Pole Custom Matls",
                            "Pole Custom Bolts",
                            "Pole Custom Reinfs",
                            "File Upload",
                            "Drilled Pier",
                            "Drilled Pier Profile",
                            "Drilled Pier Section",
                            "Drilled Pier Rebar",
                            "Belled Pier",
                            "Embedded Pole",
                            "Drilled Pier Foundation",
                            "Guy Anchor Block Tool",
                            "Guy Anchor Blocks",
                            "Guy Anchor Profiles",
                            "Leg Reinforcements",
                            "Leg Reinforcement Details",
                            "CCISeismics",
                            "Discretes",
                            "Dishes",
                            "User Forces",
                            "Lines"}

        'REMOVED site code criteria from EDS query 7/5/2023
        'It was between 'Pole Custom Reinfs' and 'File Upload'
        '"Site Code Criteria",

        Using strDS As New DataSet

            sqlLoader(query, strDS, ActiveDatabase, LogOnUser, 500)

            'name tables from tableNames list
            For i = 0 To strDS.Tables.Count - 1
                strDS.Tables(i).TableName = tableNames(i)
            Next

            LoadSiteCodeCriteria(strDS)

            'Load TNX Model
            If strDS.Tables("TNX").Rows.Count > 0 Then
                Me.tnx = New tnxModel(strDS, Me)
            End If

            'Pier and Pad
            For Each dr As DataRow In strDS.Tables("Pier and Pad").Rows
                Me.PierandPads.Add(New PierAndPad(dr, Me))
            Next

            'Unit Base
            For Each dr As DataRow In strDS.Tables("Unit Base").Rows
                Me.UnitBases.Add(New UnitBase(dr, Me))
            Next

            'Pile
            For Each dr As DataRow In strDS.Tables("Pile").Rows
                Me.Piles.Add(New Pile(dr, strDS, Me))
            Next

            'CCIplate
            For Each dr As DataRow In strDS.Tables("CCIplates").Rows
                Me.CCIplates.Add(New CCIplate(dr, strDS, Me))
            Next

            'CCIpole
            For Each dr As DataRow In strDS.Tables("Pole General").Rows
                Me.Poles.Add(New Pole(dr, strDS, Me))
            Next

            'For additional tools we'll need to update the constructor to use a datarow and pass through the dataset byref for sub tables (i.e. soil profiles)
            'That constructor will grab datarows from the sub data tables based on the foreign key in datarow


            'If a datarow is passed into the drilled pier foundation tool object then it will create new tools for every drilled pier that exists. 
            'Otherwise it will put all drilled piers into the same too. 
            'Basically all guyed towers would be in 1 tool.
            For Each dr As DataRow In strDS.Tables("Drilled Pier Foundation").Rows
                Me.DrilledPierTools.Add(New DrilledPierFoundation(strDS, Me, dr))
            Next

            'Guy Anchor Block
            For Each dr As DataRow In strDS.Tables("Guy Anchor Block Tool").Rows
                Me.GuyAnchorBlockTools.Add(New AnchorBlockFoundation(strDS, Me, dr))
            Next

            'Leg Reinforcement
            For Each dr As DataRow In strDS.Tables("Leg Reinforcements").Rows
                Me.LegReinforcements.Add(New LegReinforcement(dr, strDS, Me))
            Next

            'CCISeismic
            For Each dr As DataRow In strDS.Tables("CCISeismics").Rows
                Me.CCISeismics.Add(New CCISeismic(dr, strDS, Me))
            Next

        End Using

    End Sub

    Public Sub LoadSiteCodeCriteria(Optional ByVal strDS As DataSet = Nothing)
        If strDS Is Nothing Then strDS = New DataSet

        'If no site code criteria exists, fetch data from ORACLE to use for the first analysis. 
        'Still need to find all Topo inputs
        'Just set other parameters as default values 
        If Not strDS.Tables.Contains("Site Code Criteria") Then
            Dim sqlWhere As String = ""
            Dim sqlOrder As String = ""

            '''UNUSED PORTIONS OF THE QUERY BELOW'''
            '''WHERE STATEMENTS
            '''--wo.work_order_seqnum = 'XXXXXXX'
            '''--str.bus_unit = '" & bus_unit & "' --Comment out when switching to WO
            '''--AND str.structure_id = '" & structure_id & "' --Comment out when switching to WO" 
            '''--AND wo.bus_unit = str.bus_unit
            '''--AND wo.structure_id = str.structure_id
            '''--AND pi.eng_app_id = wo.eng_app_id(+)
            '''FROM STATEMENTS
            '''--,isit_aim.work_orders                 wo
            '''--,isit_isite.project_info              pi

            '''
            If Me.work_order_seq_num IsNot Nothing Then
                sqlWhere = vbCrLf & " ,isit_aim.work_orders wo
                            ,isit_isite.project_info pi 
                            WHERE 
                            wo.work_order_seq_num = '" & Me.work_order_seq_num.ToString & "'
                            AND wo.bus_unit = str.bus_unit 
                            AND wo.structure_id = str.structure_id
                            AND str.bus_unit = sit.bus_unit
                            AND str.bus_unit = tr.bus_unit
                            AND pi.eng_app_id(+) = wo.eng_app_id "
                sqlOrder = ",pi.eng_app_id
                            ,pi.crrnt_rvsn_num eng_app_id_revision"
            Else
                sqlWhere = vbCrLf & " WHERE 
                             str.bus_unit = '" & bus_unit & "' 
                             AND str.structure_id = '" & structure_id & "'
                             AND str.bus_unit = sit.bus_unit
                             AND str.bus_unit = tr.bus_unit "
            End If

            OracleLoader("
                        SELECT
                                str.bus_unit
                                ,str.structure_id
                                ,tr.standard_code tia_current
                                ,tr.bldg_code ibc_current
                                ,str.ground_elev elev_agl
                                ,str.hgt_no_appurt
                                ,str.crest_height
                                ,str.distance_from_crest
                                ,sit.site_name
                                ,'True' rev_h_section_15_5
                                ,0 tower_point_elev                                
                                ,str.structure_type
                                ,CAST(str.LAT_DEC as DECIMAL(12,8)) LAT_DEC
                                ,CAST(str.LONG_DEC as DECIMAL(12,8)) LONG_DEC " &
                                sqlOrder &
                            " FROM
                                isit_aim.structure                      str
                                ,isit_aim.site                          sit
                                ,rpt_appl.eng_tower_rating_vw           tr " &
                            sqlWhere, "Site Code Criteria", strDS, 3000, "ords")
        End If

        If strDS.Tables.Contains("Site Code Criteria") Then
            If strDS.Tables("Site Code Criteria").Rows.Count > 0 Then
                Me.structureCodeCriteria = New SiteCodeCriteria(strDS.Tables("Site Code Criteria").Rows(0)) 'Need to comment out when using dummy BU numbers - MRR
            End If
        End If

        SetStructureSiteInfo()
    End Sub

    Private Sub SetStructureSiteInfo()
        If Me.SiteInfo IsNot Nothing Then
            Me.order = Me.SiteInfo.app_num
            Me.orderRev = Me.SiteInfo.app_rev
            Me.work_order_seq_num = Me.SiteInfo.wo
        End If
    End Sub

    Public Function SavetoEDSQuery() As String

        If EDSMe Is Nothing Then
            Throw New Exception("EDS Structure object not set.")
        End If

        Dim structureQuery As String = ""


        For Each level In _SQLQueryVariables
            structureQuery.NewLine("DECLARE " & level & " TABLE(ID INT)")
            structureQuery.NewLine("DECLARE " & level & "ID INT")
        Next
        If Me.Poles.Count > 0 Then
            structureQuery.NewLine("DECLARE @TopBoltID INT" & vbCrLf & "DECLARE @BotBoltID INT")
        End If


        Dim transactionName As String = Me.process_stage & Me.structure_id & Me.bus_unit.NullableToString & Me.work_order_seq_num.NullableToString
        structureQuery.NewLine("BEGIN TRY")
        structureQuery.NewLine("BEGIN TRANSACTION " & transactionName)
        structureQuery.NewLine(Me.tnx?.EDSQueryBuilder(EDSMe.tnx))
        structureQuery.NewLine(Me.PierandPads.EDSListQueryBuilder(EDSMe.PierandPads))
        structureQuery.NewLine(Me.UnitBases.EDSListQueryBuilder(EDSMe.UnitBases))
        structureQuery.NewLine(Me.Piles.EDSListQueryBuilder(EDSMe.Piles))
        structureQuery.NewLine(Me.DrilledPierTools.EDSListQueryBuilder(EDSMe.DrilledPierTools))
        structureQuery.NewLine(Me.GuyAnchorBlockTools.EDSListQueryBuilder(EDSMe.GuyAnchorBlockTools))
        structureQuery.NewLine(Me.CCIplates.EDSListQueryBuilder(EDSMe.CCIplates))
        structureQuery.NewLine(Me.Poles.EDSListQueryBuilder(EDSMe.Poles))
        structureQuery.NewLine(Me.LegReinforcements.EDSListQueryBuilder(EDSMe.LegReinforcements))
        structureQuery.NewLine(Me.CCISeismics.EDSListQueryBuilder(EDSMe.CCISeismics))
        structureQuery.NewLine("COMMIT TRANSACTION " & transactionName)
        structureQuery.NewLine("SELECT '" & Me.bus_unit.NullableToString & "' bus_unit, '" & Me.structure_id & "' structure_id, '" & Me.work_order_seq_num.NullableToString & "' work_order_seq_num, '" & "Success" & "' result")
        structureQuery.NewLine("END TRY")
        structureQuery.NewLine("BEGIN CATCH")
        structureQuery.NewLine("ROLLBACK TRANSACTION " & transactionName)
        structureQuery.NewLine("EXECUTE dbo.usp_GetErrorInfo '" & Me.bus_unit.NullableToString & "', '" & Me.structure_id & "', '" & Me.work_order_seq_num.NullableToString & "'")
        structureQuery.NewLine("END CATCH")

        Return structureQuery


    End Function

    Public Function DeleteFromEDSQuery() As String

        Dim structureQuery As String = ""


        For Each level In _SQLQueryVariables
            structureQuery.NewLine("DECLARE " & level & " TABLE(ID INT)")
            structureQuery.NewLine("DECLARE " & level & "ID INT")
        Next
        If Me.Poles.Count > 0 Then
            structureQuery.NewLine("DECLARE @TopBoltID INT" & vbCrLf & "DECLARE @BotBoltID INT")
        End If


        Dim transactionName As String = "deleting" & Me.process_stage & Me.structure_id & Me.bus_unit.NullableToString & Me.work_order_seq_num.NullableToString
        structureQuery.NewLine("BEGIN TRY")
        structureQuery.NewLine("BEGIN TRANSACTION " & transactionName)
        structureQuery.NewLine(Me.tnx?.SQLDelete)
        structureQuery.NewLine(Me.PierandPads.EDSDELETEListQueryBuilder())
        structureQuery.NewLine(Me.UnitBases.EDSDELETEListQueryBuilder())
        structureQuery.NewLine(Me.Piles.EDSDELETEListQueryBuilder())
        structureQuery.NewLine(Me.DrilledPierTools.EDSDELETEListQueryBuilder())
        structureQuery.NewLine(Me.GuyAnchorBlockTools.EDSDELETEListQueryBuilder())
        structureQuery.NewLine(Me.CCIplates.EDSDELETEListQueryBuilder())
        structureQuery.NewLine(Me.Poles.EDSDELETEListQueryBuilder())
        structureQuery.NewLine(Me.LegReinforcements.EDSDELETEListQueryBuilder())
        structureQuery.NewLine(Me.CCISeismics.EDSDELETEListQueryBuilder())
        structureQuery.NewLine(Me.ReportOptions.SQLDelete)
        structureQuery.NewLine("COMMIT TRANSACTION " & transactionName)
        structureQuery.NewLine("SELECT '" & Me.bus_unit.NullableToString & "' bus_unit, '" & Me.structure_id & "' structure_id, '" & Me.work_order_seq_num.NullableToString & "' work_order_seq_num, '" & "Success" & "' result")
        structureQuery.NewLine("END TRY")
        structureQuery.NewLine("BEGIN CATCH")
        structureQuery.NewLine("ROLLBACK TRANSACTION " & transactionName)
        structureQuery.NewLine("EXECUTE dbo.usp_GetErrorInfo '" & Me.bus_unit.NullableToString & "', '" & Me.structure_id & "', '" & Me.work_order_seq_num.NullableToString & "'")
        structureQuery.NewLine("END CATCH")

        Return structureQuery


    End Function

    ''' <summary>
    ''' Save to EDS functionality to save the whole structure to EDS.
    ''' Returns a tuple of Boolean and Datatable
    ''' The datatable will contain the error or success information 
    ''' Boolean = true was a successful save --> Boolean = false was a failure.
    ''' </summary>
    ''' <param name="databaseID"></param>
    ''' <param name="ActiveDatabase"></param>
    ''' <param name="copyQueryToClipboard"></param>
    ''' <returns>Item1 = saveCheck, Item2 = Datatable of information</returns>
    Public Function SavetoEDS(ByVal Optional databaseID As WindowsIdentity = Nothing, ByVal Optional ActiveDatabase As String = Nothing, Optional ByVal copyQueryToClipboard As Boolean = False, Optional ByVal commitQuery As Boolean = True) As (success As Boolean, returner As DataTable, query As String)
        If databaseID Is Nothing Then databaseID = Me.databaseIdentity
        If ActiveDatabase Is Nothing Then ActiveDatabase = Me.activeDatabase

        If EDSMe Is Nothing Then EDSMe = New EDSStructure(Me.bus_unit, Me.structure_id, Me.work_order_seq_num, databaseID, ActiveDatabase)

        Dim resDS As New DataSet
        Dim myQuery As String = SavetoEDSQuery()

        If copyQueryToClipboard Then
            My.Computer.Clipboard.SetText(myQuery)
        End If

        If commitQuery Then
            sqlLoader(myQuery, resDS, ActiveDatabase, databaseID, 4051.ToString)

            If resDS.Tables.Count > 0 Then
                'Check for success or failure here. Changing this to a function to return as a datatable of informatoin 
                Dim saveCheck As Boolean = True
                If resDS.Tables(0).Rows(0).Item("Result").ToString = "Error" Then
                    saveCheck = False
                End If
                Return (saveCheck, resDS.Tables(0), myQuery)
            Else
                Return (False, Nothing, myQuery)
            End If
        Else
            Return (False, Nothing, myQuery)
        End If
    End Function

    Public Function DeleteFromEDS(ByVal Optional databaseID As WindowsIdentity = Nothing, ByVal Optional ActiveDatabase As String = Nothing, Optional ByVal copyQueryToClipboard As Boolean = False, Optional ByVal commitQuery As Boolean = True) As (result As Boolean, returner As DataTable, query As String)
        If databaseID Is Nothing Then databaseID = Me.databaseIdentity
        If ActiveDatabase Is Nothing Then ActiveDatabase = Me.activeDatabase

        Dim resDS As New DataSet
        Dim myQuery As String = DeleteFromEDSQuery()

        If copyQueryToClipboard Then
            My.Computer.Clipboard.SetText(myQuery)
        End If

        If commitQuery Then
            sqlLoader(myQuery, resDS, ActiveDatabase, databaseID, 4052.ToString)

            If resDS.Tables.Count > 0 Then
                'Check for success or failure here. Changing this to a function to return as a datatable of informatoin 
                Dim saveCheck As Boolean = True
                If resDS.Tables(0).Rows(0).Item("Result").ToString = "Error" Then
                    saveCheck = False
                End If
                Return (saveCheck, resDS.Tables(0), myQuery)
            Else
                Return (False, Nothing, myQuery)
            End If
        Else
            Return (False, Nothing, myQuery)
        End If
    End Function

    Public Sub RefreshReportOptions()
        Dim newRep As ReportOptions = New ReportOptions()
        newRep.WorkingDir = Me.ReportOptions.WorkingDir
        newRep.ReportDir = Me.ReportOptions.ReportDir
        newRep.Absorb(Me)
        newRep.Initialize()

        'If newRep.IsFromDB Then
        Me.ReportOptions = newRep
        'End If
    End Sub
#End Region

#Region "Files"
    Public Sub ResetTools()
        Me.tnx = Nothing
        Me.PierandPads = New List(Of PierAndPad)
        Me.Piles = New List(Of Pile)
        Me.UnitBases = New List(Of UnitBase)
        Me.DrilledPierTools = New List(Of DrilledPierFoundation)
        Me.CCIplates = New List(Of CCIplate)
        Me.Poles = New List(Of Pole)
        Me.LegReinforcements = New List(Of LegReinforcement)
        Me.CCISeismics = New List(Of CCISeismic)
    End Sub

    Public Sub LoadFromFiles(filePaths As String(), Optional LoadSiteData As Boolean = True)
        If LoadSiteData Then LoadSiteCodeCriteria()

        Me.PierandPads = New List(Of PierAndPad)
        Me.Piles = New List(Of Pile)
        Me.UnitBases = New List(Of UnitBase)
        Me.DrilledPierTools = New List(Of DrilledPierFoundation)
        Me.GuyAnchorBlockTools = New List(Of AnchorBlockFoundation)
        Me.CCIplates = New List(Of CCIplate)
        Me.Poles = New List(Of Pole)
        Me.LegReinforcements = New List(Of LegReinforcement)
        Me.CCISeismics = New List(Of CCISeismic)

        For Each item As String In filePaths
            Dim itemName As String = System.IO.Path.GetFileName(item)
            If itemName.EndsWith(".eri") Then
                Me.tnx = New tnxModel(item, Me)
            ElseIf itemName.Contains("Pier and Pad Foundation") And itemName.EndsWith(".xlsm") Then
                Me.PierandPads.Add(New PierAndPad(item, Me))
            ElseIf itemName.Contains("Pile Foundation") And itemName.EndsWith(".xlsm") Then
                Me.Piles.Add(New Pile(item, Me))
            ElseIf itemName.Contains("SST Unit Base Foundation") And itemName.EndsWith(".xlsm") Then
                Me.UnitBases.Add(New UnitBase(item, Me))
            ElseIf itemName.Contains("Drilled Pier Foundation") And itemName.EndsWith(".xlsm") Then
                Me.DrilledPierTools.Add(New DrilledPierFoundation(item, Me))
            ElseIf itemName.Contains("Guyed Anchor Block Foundation") And itemName.EndsWith(".xlsm") Then
                Me.GuyAnchorBlockTools.Add(New AnchorBlockFoundation(item, Me))
            ElseIf itemName.Contains("CCIplate") And itemName.EndsWith(".xlsm") Then
                Me.CCIplates.Add(New CCIplate(item, Me))
            ElseIf itemName.Contains("CCIpole") And itemName.EndsWith(".xlsm") Then
                Me.Poles.Add(New Pole(item, Me))
            ElseIf itemName.Contains("Leg Reinforcement Tool") And itemName.EndsWith(".xlsm") Then
                Me.LegReinforcements.Add(New LegReinforcement(item, Me))
            ElseIf itemName.Contains("CCISeismic") And itemName.EndsWith(".xlsm") Then
                Me.CCISeismics.Add(New CCISeismic(item, Me))
            End If
        Next

    End Sub


    ''' <summary>
    ''' Save all tools.
    ''' Will overwrite existing files by default.
    ''' </summary>
    ''' <param name="folderPath"></param>
    ''' <param name="replaceFiles">Determines if existing files are overwritten or skipped.</param>
    Public Sub SaveTools(Optional folderPath As String = "", Optional replaceFiles As Boolean = True)

        If folderPath = "" Then
            folderPath = Me.WorkingDirectory
        End If

        If Not Directory.Exists(folderPath) Then Exit Sub

        'Uncomment your foundation type for testing when it's ready.
        Dim i As Integer
        Dim fileNum As String = ""

        If Me.tnx IsNot Nothing Then Me.tnx.GenerateERI(Path.Combine(folderPath, Me.bus_unit & ".eri"), replaceFiles)

        For i = 0 To Me.PierandPads.Count - 1
            PierandPads(i).SavetoExcel(index:=i, replaceFiles:=replaceFiles, FolderPath:=folderPath)
        Next
        For i = 0 To Me.Piles.Count - 1
            Piles(i).SavetoExcel(index:=i, replaceFiles:=replaceFiles, FolderPath:=folderPath)
        Next
        For i = 0 To Me.UnitBases.Count - 1
            UnitBases(i).SavetoExcel(index:=i, replaceFiles:=replaceFiles, FolderPath:=folderPath)
        Next
        For i = 0 To Me.DrilledPierTools.Count - 1
            DrilledPierTools(i).SavetoExcel(index:=i, replaceFiles:=replaceFiles, FolderPath:=folderPath)
        Next
        For i = 0 To Me.GuyAnchorBlockTools.Count - 1
            GuyAnchorBlockTools(i).SavetoExcel(index:=i, replaceFiles:=replaceFiles, FolderPath:=folderPath)
        Next
        For i = 0 To Me.CCIplates.Count - 1
            CCIplates(i).SavetoExcel(index:=i, replaceFiles:=replaceFiles, FolderPath:=folderPath)
        Next
        For i = 0 To Me.Poles.Count - 1
            Poles(i).SavetoExcel(index:=i, replaceFiles:=replaceFiles, FolderPath:=folderPath)
        Next
        For i = 0 To Me.LegReinforcements.Count - 1
            LegReinforcements(i).SavetoExcel(index:=i, replaceFiles:=replaceFiles, FolderPath:=folderPath)
        Next
        For i = 0 To Me.CCISeismics.Count - 1
            CCISeismics(i).SavetoExcel(index:=i, replaceFiles:=replaceFiles, FolderPath:=folderPath)
        Next
    End Sub

    ''' <summary>
    ''' Save all tools using a delegate function to determine if files should be overwritten.
    ''' This allows the Dashboard UI to promt the user for each file replacement.
    ''' </summary>
    ''' <param name="overwriteFile">Delegate function which take a file path as a string and returns a boolean.</param>
    ''' <param name="folderPath">Save to folder path.</param>
    Public Sub SaveTools(overwriteFile As OverwriteFile, Optional folderPath As String = "")

        If folderPath = "" Then
            folderPath = Me.WorkingDirectory
        End If

        If Not Directory.Exists(folderPath) Then Exit Sub

        'Uncomment your foundation type for testing when it's ready.
        Dim i As Integer
        Dim fileNum As String = ""

        If Me.tnx IsNot Nothing Then Me.tnx.GenerateERI(overwriteFile, Path.Combine(folderPath, Me.bus_unit & ".eri"))

        For i = 0 To Me.PierandPads.Count - 1
            PierandPads(i).SavetoExcel(overwriteFile, index:=i)
        Next
        For i = 0 To Me.Piles.Count - 1
            Piles(i).SavetoExcel(overwriteFile, index:=i)
        Next
        For i = 0 To Me.UnitBases.Count - 1
            UnitBases(i).SavetoExcel(overwriteFile, index:=i)
        Next
        For i = 0 To Me.DrilledPierTools.Count - 1
            DrilledPierTools(i).SavetoExcel(overwriteFile, index:=i)
        Next
        For i = 0 To Me.GuyAnchorBlockTools.Count - 1
            GuyAnchorBlockTools(i).SavetoExcel(overwriteFile, index:=i)
        Next
        For i = 0 To Me.CCIplates.Count - 1
            CCIplates(i).SavetoExcel(overwriteFile, index:=i)
        Next
        For i = 0 To Me.Poles.Count - 1
            Poles(i).SavetoExcel(overwriteFile, index:=i)
        Next
        For i = 0 To Me.LegReinforcements.Count - 1
            LegReinforcements(i).SavetoExcel(overwriteFile, index:=i)
        Next
        For i = 0 To Me.CCISeismics.Count - 1
            CCISeismics(i).SavetoExcel(overwriteFile, index:=i)
        Next
    End Sub
#End Region

#Region "Check Changes"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As EDSStructure = TryCast(other, EDSStructure)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.tnx.CheckChange(otherToCompare.tnx, changes, categoryName, "TNX"), Equals, False)
        Equals = If(Me.CCIplates.CheckChange(otherToCompare.CCIplates, changes, categoryName, "Connections"), Equals, False)
        Equals = If(Me.Poles.CheckChange(otherToCompare.Poles, changes, categoryName, "Pole"), Equals, False)
        'Equals = If(Me.structureCodeCriteria.CheckChange(otherToCompare.structureCodeCriteria, changes, categoryName, "Structure Code Criteria"), Equals, False) 'Deactivated since causes errors
        Equals = If(Me.PierandPads.CheckChange(otherToCompare.PierandPads, changes, categoryName, "Pier and Pads"), Equals, False)
        Equals = If(Me.Piles.CheckChange(otherToCompare.Piles, changes, categoryName, "Piles"), Equals, False)
        Equals = If(Me.UnitBases.CheckChange(otherToCompare.UnitBases, changes, categoryName, "Unit Bases"), Equals, False)
        Equals = If(Me.DrilledPierTools.CheckChange(otherToCompare.DrilledPierTools, changes, categoryName, "Drilled Piers"), Equals, False)
        Equals = If(Me.GuyAnchorBlockTools.CheckChange(otherToCompare.GuyAnchorBlockTools, changes, categoryName, "Guy Anchor Blocks"), Equals, False)
        Equals = If(Me.LegReinforcements.CheckChange(otherToCompare.LegReinforcements, changes, categoryName, "LegReinforcements"), Equals, False)
        Equals = If(Me.CCISeismics.CheckChange(otherToCompare.CCISeismics, changes, categoryName, "CCISeismics"), Equals, False)

        Return Equals

    End Function

#End Region
End Class









