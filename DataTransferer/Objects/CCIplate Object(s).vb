Option Strict On

Imports System.ComponentModel
Imports System.Runtime.Serialization
Imports DevExpress.Spreadsheet
'Imports Microsoft.Office.Interop

<DataContractAttribute()>
Partial Public Class CCIplate
    Inherits EDSExcelObject


#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "CCIplate"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.connections"
        End Get
    End Property
    Public Overrides ReadOnly Property TemplatePath As String
        Get
            IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "CCIplate.xlsm")
        End Get
    End Property
    Public Overrides ReadOnly Property Template As Byte()
        Get
            Return CCI_Engineering_Templates.My.Resources.CCIplate
        End Get
    End Property
    Public Overrides ReadOnly Property ExcelDTParams As List(Of EXCELDTParameter)
        'Add additional sub table references here. Table names should be consistent with EDS table names. 
        Get
            Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("CCIplates", "A1:K2", "Details (SAPI)"),
                                                        New EXCELDTParameter("Connections", "C2:G18", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Plate Details", "I2:S33", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Bolt Groups", "U2:AC82", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Bolt Details", "AE2:AQ1602", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("CCIplate Materials", "AS2:BE55", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Plate Results", "F2:I499", "Results (SAPI)"),
                                                        New EXCELDTParameter("Bolt Results", "K2:O83", "Results (SAPI)"),
                                                        New EXCELDTParameter("Stiffener Groups", "BG2:BJ157", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Stiffener Details", "BL2:CB3102", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Bridge Stiffener Details", "CD2:DG52", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Connection Results", "A2:D52", "Results (SAPI)")}

            'note: Excel table names are consistent with EDS table names to limit work required within constructors

        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String

        If _Insert = "" Then
            '_Insert = QueryBuilderFromFile(queryPath & "CCIplate\CCIplate (INSERT).sql")
            _Insert = CCI_Engineering_Templates.My.Resources.CCIplate_INSERT
        End If
        SQLInsert = _Insert

        'Details
        SQLInsert = SQLInsert.Replace("[CCIPLATE VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[CCIPLATE FIELDS]", Me.SQLInsertFields)

        'Connection
        If Me.Connections.Count > 0 Then
            SQLInsert = SQLInsert.Replace("--BEGIN --[CONNECTION INSERT BEGIN]", "BEGIN --[CONNECTION INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[CONNECTION INSERT END]", "END --[CONNECTION INSERT END]")
            For Each row As Connection In Connections
                SQLInsert = SQLInsert.Replace("--[CONNECTION INSERT]", row.SQLInsert)
            Next
        End If

        'note: additional insert commands are imbedded within objects sharing similar relationships (e.g. plate details insert located within Connections Object)


        ''Results
        'If Me.Results.Count > 0 Then
        '    SQLInsert = SQLInsert.Replace("--BEGIN --[RESULTS INSERT BEGIN]", "BEGIN --[RESULTS INSERT BEGIN]")
        '    SQLInsert = SQLInsert.Replace("--END --[RESULTS INSERT END]", "END --[RESULTS INSERT END]")
        '    SQLInsert = SQLInsert.Replace("--[RESULTS INSERT]", Me.Results.EDSResultQuery)
        'End If

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String
        'This section not only needs to call update commands but also needs to call insert and delete commands since subtables may involve adding or deleting records

        If _Update = "" Then
            '_Update = QueryBuilderFromFile(queryPath & "CCIplate\CCIplate (UPDATE).sql")
            _Update = CCI_Engineering_Templates.My.Resources.CCIplate_UPDATE
        End If
        SQLUpdate = _Update

        'Details
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)

        'Connection
        If Me.Connections.Count > 0 Then
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[CONNECTION UPDATE BEGIN]", "BEGIN --[CONNECTION UPDATE BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[CONNECTION UPDATE END]", "END --[CONNECTION UPDATE END]")
            For Each row As Connection In Connections
                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                    If IsSomething(row.connection_elevation) Or IsSomethingString(row.connection_type) Or IsSomethingString(row.bolt_configuration) Then
                        SQLUpdate = SQLUpdate.Replace("--[CONNECTION INSERT]", row.SQLUpdate)
                    Else
                        SQLUpdate = SQLUpdate.Replace("--[CONNECTION INSERT]", row.SQLDelete)
                    End If
                Else
                    SQLUpdate = SQLUpdate.Replace("--[CONNECTION INSERT]", row.SQLInsert)
                End If
            Next
        End If

        'note: additional update commands are imbedded within objects sharing similar relationships (e.g. plate details update located within Connections Object)


        ''Results
        'If Me.Results.Count > 0 Then
        '    SQLUpdate = SQLUpdate.Replace("--BEGIN --[RESULTS UPDATE BEGIN]", "BEGIN --[RESULTS UPDATE BEGIN]")
        '    SQLUpdate = SQLUpdate.Replace("--END --[RESULTS UPDATE END]", "END --[RESULTS UPDATE END]")
        '    SQLUpdate = SQLUpdate.Replace("--[RESULTS INSERT]", Me.Results.EDSResultQuery)
        'End If

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        If _Delete = "" Then
            '_Delete = QueryBuilderFromFile(queryPath & "CCIplate\CCIplate (DELETE).sql")
            _Delete = CCI_Engineering_Templates.My.Resources.CCIplate_DELETE
        End If
        SQLDelete = _Delete
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)

        'Plate
        If Me.Connections.Count > 0 Then
            SQLDelete = SQLDelete.Replace("--BEGIN --[CONNECTION DELETE BEGIN]", "BEGIN --[CONNECTION DELETE BEGIN]")
            SQLDelete = SQLDelete.Replace("--END --[CONNECTION DELETE END]", "END --[CONNECTION DELETE END]")
            For Each row As Connection In Connections
                SQLDelete = SQLDelete.Replace("--[CONNECTION INSERT]", row.SQLDelete)
            Next
        End If

        'note: additional delete commands are imbedded within objects sharing similar relationships (e.g. plate details delete located within Connections Object)

        Return SQLDelete

    End Function

#End Region

#Region "Define"

    'Private _ID As Integer? 'Defined in EDSObject
    'Private _bus_unit As String 'Defined in EDSObject
    'Private _structure_id As String 'Defined in EDSObject
    Private _anchor_rod_spacing As Double?
    Private _clip_distance As Double?
    Private _barb_cl_elevation As Double?
    Private _include_pole_reactions As Boolean?
    Private _consider_ar_eccentricity As Boolean?
    Private _leg_mod_eccentricity As Double?
    Private _seismic As Boolean?
    Private _seismic_flanges As Boolean?
    Private _Structural_105 As Boolean?
    'Private _tool_version As String 'Defined in EDSExcelObject
    'Private _modified_person_id As Integer? 'Defined in EDSExcelObject
    'Private _process_stage As String 'Defined in EDSExcelObject

    <DataMember()> Public Property Connections As New List(Of Connection)
    <DataMember()> Public Property CCIplateMaterials As New List(Of CCIplateMaterial)

    <Category("CCIplate"), Description(""), DisplayName("Anchor Rod Spacing")>
    <DataMember()> Public Property anchor_rod_spacing() As Double?
        Get
            Return Me._anchor_rod_spacing
        End Get
        Set
            Me._anchor_rod_spacing = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Clip Distance")>
    <DataMember()> Public Property clip_distance() As Double?
        Get
            Return Me._clip_distance
        End Get
        Set
            Me._clip_distance = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Barb Cl Elevation")>
    <DataMember()> Public Property barb_cl_elevation() As Double?
        Get
            Return Me._barb_cl_elevation
        End Get
        Set
            Me._barb_cl_elevation = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Include Pole Reactions")>
    <DataMember()> Public Property include_pole_reactions() As Boolean?
        Get
            Return Me._include_pole_reactions
        End Get
        Set
            Me._include_pole_reactions = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Consider Ar Eccentricity")>
    <DataMember()> Public Property consider_ar_eccentricity() As Boolean?
        Get
            Return Me._consider_ar_eccentricity
        End Get
        Set
            Me._consider_ar_eccentricity = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Leg Mod Eccentricity")>
    <DataMember()> Public Property leg_mod_eccentricity() As Double?
        Get
            Return Me._leg_mod_eccentricity
        End Get
        Set
            Me._leg_mod_eccentricity = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Seismic")>
    <DataMember()> Public Property seismic() As Boolean?
        Get
            Return Me._seismic
        End Get
        Set
            Me._seismic = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Seismic Flanges")>
    <DataMember()> Public Property seismic_flanges() As Boolean?
        Get
            Return Me._seismic_flanges
        End Get
        Set
            Me._seismic_flanges = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Structural 105")>
    <DataMember()> Public Property Structural_105() As Boolean?
        Get
            Return Me._Structural_105
        End Get
        Set
            Me._Structural_105 = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    'Overriding the Results field to return a list of all connection, plate, and bolt results (casted as EDSResult objects)
    Private _results As List(Of EDSResult)
    <Category("Ratio"), Description("This rating takes into account TIA-222-H Annex S Section 15.5 when applicable."), DisplayName("Rating")>
    <DataMember()>
    Public Overrides Property Results As List(Of EDSResult)
        Get
            Dim returnThis As New List(Of EDSResult)()

            'Get all connections from all CCIPlates and order descending
            Dim AllConnections As List(Of Connection) = Me.Connections _
            .OrderByDescending(Function(c) c.connection_elevation) _
            .ToList()

            For Each connection As Connection In AllConnections
                'BoltGroups
                For Each boltGroup As BoltGroup In connection.BoltGroups
                    For Each boltResult As BoltResults In boltGroup.BoltResults
                        Dim edsResult As New EDSResult()
                        edsResult.result_lkup = boltResult.result_lkup
                        edsResult.rating = boltResult.rating
                        edsResult.modified_person_id = boltResult.modified_person_id
                        edsResult.process_stage = boltResult.process_stage
                        edsResult.EDSTableDepth = boltResult.EDSTableDepth + 1
                        edsResult.EDSTableName = "conn.bolt" & "_results"
                        edsResult.ForeignKeyName = "bolt" & "_id"
                        edsResult.foreign_key = boltResult.Parent.ID
                        returnThis.Add(edsResult)
                    Next
                Next

                'ConnectionResults
                For Each connectionResult As ConnectionResults In connection.ConnectionResults
                    Dim edsResult As New EDSResult()
                    edsResult.result_lkup = connectionResult.result_lkup
                    edsResult.rating = connectionResult.rating
                    edsResult.modified_person_id = connectionResult.modified_person_id
                    edsResult.process_stage = connectionResult.process_stage
                    edsResult.EDSTableDepth = connectionResult.EDSTableDepth + 1
                    edsResult.EDSTableName = "conn.connection" & "_results"
                    edsResult.ForeignKeyName = "plate" & "_id"
                    edsResult.foreign_key = connectionResult.Parent.ID
                    returnThis.Add(edsResult)
                Next

                'PlateDetails
                For Each plateDetail As PlateDetail In connection.PlateDetails
                    For Each plateResult As PlateResults In plateDetail.PlateResults
                        Dim edsResult As New EDSResult()
                        edsResult.result_lkup = plateResult.result_lkup
                        edsResult.rating = plateResult.rating
                        edsResult.modified_person_id = plateResult.modified_person_id
                        edsResult.process_stage = plateResult.process_stage
                        edsResult.EDSTableDepth = plateResult.EDSTableDepth + 1
                        edsResult.EDSTableName = "conn.plate" & "_results"
                        edsResult.ForeignKeyName = "plate_details" & "_id"
                        edsResult.foreign_key = plateResult.Parent.ID
                        returnThis.Add(edsResult)

                    Next
                Next
            Next
            Return returnThis
        End Get
        Set(value As List(Of EDSResult))
            Me._results = value
        End Set
    End Property

    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal dr As DataRow, ByRef strDS As DataSet, Optional ByVal Parent As EDSObject = Nothing) 'Added strDS in order to pull EDS data from subtables
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        'Following used to create dataset, regardless if source was EDS or Excel. Boolean used to identify source. EDS = True
        BuildFromDataset(dr, strDS, True, Me)

    End Sub 'Generate a CCIplate from EDS

    Public Sub New(ExcelFilePath As String, Optional ByVal Parent As EDSObject = Nothing)
        Me.WorkBookPath = ExcelFilePath
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)
        LoadFromExcel()

    End Sub 'Generate a CCIplate from Excel

    Private Sub BuildFromDataset(ByVal dr As DataRow, ByRef ds As DataSet, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing)
        'Dataset is pulled in from either EDS or Excel. True = EDS, False = Excel
        'If Parent IsNot Nothing Then Me.Absorb(Parent) 'Do not double absorb!!!

        'Not sure this is necessary, could just read the values from the structure code criteria when creating the Excel sheet (Added to Save to Excel Section)
        'Me.tia_current = Me.ParentStructure?.structureCodeCriteria?.tia_current
        'Me.rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
        'Me.seismic_design_category = Me.ParentStructure?.structureCodeCriteria?.seismic_design_category

        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.bus_unit = If(EDStruefalse, DBtoStr(dr.Item("bus_unit")), Me.bus_unit) 'Not provided in Excel
        Me.structure_id = If(EDStruefalse, DBtoStr(dr.Item("structure_id")), Me.structure_id) 'Not provided in Excel
        Me.anchor_rod_spacing = DBtoNullableDbl(dr.Item("anchor_rod_spacing"))
        Me.clip_distance = DBtoNullableDbl(dr.Item("clip_distance"))
        Me.barb_cl_elevation = DBtoNullableDbl(dr.Item("barb_cl_elevation"))
        Me.include_pole_reactions = DBtoNullableBool(dr.Item("include_pole_reactions"))
        Me.consider_ar_eccentricity = DBtoNullableBool(dr.Item("consider_ar_eccentricity"))
        Me.leg_mod_eccentricity = DBtoNullableDbl(dr.Item("leg_mod_eccentricity"))
        Me.seismic = DBtoNullableBool(dr.Item("seismic"))
        Me.seismic_flanges = DBtoNullableBool(dr.Item("seismic_flanges"))
        Me.Structural_105 = DBtoNullableBool(dr.Item("Structural_105"))
        Me.Version = DBtoStr(dr.Item("tool_version"))
        Me.modified_person_id = If(EDStruefalse, DBtoNullableInt(dr.Item("modified_person_id")), Me.modified_person_id) 'Not provided in Excel
        Me.process_stage = If(EDStruefalse, DBtoStr(dr.Item("process_stage")), Me.process_stage) 'Not provided in Excel

        Dim plConnection As New Connection 'Connection
        Dim plPlateDetail As New PlateDetail ' Plate Detail
        Dim plBoltGroup As New BoltGroup ' Bolt Group
        Dim plBoltDetail As New BoltDetail ' Bolt Detail
        Dim plCCIplateMaterial As New CCIplateMaterial 'CCIplate Material
        Dim plPlateResult As New PlateResults 'Plate Results
        Dim plBoltResult As New BoltResults 'Bolt Results
        Dim plStiffGroup As New StiffenerGroup ' Stiffener Group
        Dim plStiffDetail As New StiffenerDetail ' Stiffener Detail
        Dim plBridgeDetail As New BridgeStiffenerDetail 'Bridge Stiffener Detail
        Dim plConnectionResult As New ConnectionResults 'Connection Results (BARB & Bridge stiffeners)
        'Dim plStiffenerResult As New StiffenerResults 'Stiffener Results (not required, associated with plate details)

        'Storing all default materials to CCIplate object
        'added this to help determine whether or not materials need to be added to CCIplate when pulling in from CCIpole (When source is EDS)
        For Each mrow As DataRow In ds.Tables(plCCIplateMaterial.EDSObjectName).Rows
            plCCIplateMaterial = New CCIplateMaterial(mrow, EDStruefalse, Me)
            If If(EDStruefalse, plCCIplateMaterial.default_material = True, False) Then
                'plPlateDetail.plate_material = plConnectionMaterial
                CCIplateMaterials.Add(plCCIplateMaterial)
            End If
        Next


        For Each crow As DataRow In ds.Tables(plConnection.EDSObjectName).Rows
            'create a new connection based on the datarow from above
            plConnection = New Connection(crow, EDStruefalse, Me)
            'Check if the parent id, in the case cciplate id is equal to the original object id (Me)                    
            If If(EDStruefalse, plConnection.cciplate_id = Me.ID, True) Then 'If coming from Excel, all connections provided will be associated to CCIplate. 
                'If it is equal then add the newly created connection to the list of connections 
                Connections.Add(plConnection)

                'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.structure_type) Then
                '    If Me.ParentStructure?.structureCodeCriteria?.structure_type = "SELF SUPPORT" Then
                '        structure_type = "Self Suppot"
                '    ElseIf Me.ParentStructure?.structureCodeCriteria?.structure_type = "MONOPOLE" Then
                '        structure_type = "Monopole"
                '    Else
                '        structure_type = "Monopole"
                '    End If
                '    .Worksheets("Main").Range("tower_type").Value = CType(structure_type, String)
                'End If

                'Loop through all plates pulled from EDS and check if they are associated with the newly created connection
                For Each pdrow As DataRow In ds.Tables(plPlateDetail.EDSObjectName).Rows
                    'Create a new plate from the plate datarow from EDS
                    plPlateDetail = New PlateDetail(pdrow, EDStruefalse, Me)
                    If If(EDStruefalse, plPlateDetail.connection_id = plConnection.ID, plPlateDetail.local_connection_id = plConnection.local_id) Then
                        plConnection.PlateDetails.Add(plPlateDetail)
                        For Each mrow As DataRow In ds.Tables(plCCIplateMaterial.EDSObjectName).Rows
                            plCCIplateMaterial = New CCIplateMaterial(mrow, EDStruefalse, Me)
                            If If(EDStruefalse, plCCIplateMaterial.ID = plPlateDetail.plate_material, plCCIplateMaterial.local_id = plPlateDetail.plate_material) Then
                                'plPlateDetail.plate_material = plConnectionMaterial
                                plPlateDetail.CCIplateMaterials.Add(plCCIplateMaterial)
                                Exit For 'Once matched, don't need to continue checking. 
                            End If
                        Next
                        If IsSomething(ds.Tables("Plate Results")) Then
                            For Each prrow As DataRow In ds.Tables("Plate Results").Rows
                                plPlateResult = New PlateResults(prrow, EDStruefalse, Me)
                                If If(EDStruefalse, False, plPlateResult.local_plate_id = plPlateDetail.local_id) Then
                                    plPlateDetail.PlateResults.Add(plPlateResult)
                                End If
                            Next
                        End If

                        'Stiffeners
                        For Each sgrow As DataRow In ds.Tables(plStiffGroup.EDSObjectName).Rows
                            plStiffGroup = New StiffenerGroup(sgrow, EDStruefalse, Me)
                            If If(EDStruefalse, plStiffGroup.plate_details_id = plPlateDetail.ID, plStiffGroup.local_plate_id = plPlateDetail.local_id) Then
                                plPlateDetail.StiffenerGroups.Add(plStiffGroup)
                                For Each sdrow As DataRow In ds.Tables(plStiffDetail.EDSObjectName).Rows
                                    plStiffDetail = New StiffenerDetail(sdrow, EDStruefalse, Me)
                                    If If(EDStruefalse, plStiffDetail.stiffener_id = plStiffGroup.ID, If(plStiffDetail.local_group_id > 0, plStiffDetail.local_plate_id = plPlateDetail.local_id And plStiffDetail.local_group_id = plStiffGroup.local_id, plStiffDetail.local_plate_id = plPlateDetail.local_id And plStiffDetail.stiffener_id = plStiffGroup.ID)) Then
                                        plStiffGroup.StiffenerDetails.Add(plStiffDetail)
                                    End If
                                Next
                                'Stiffener results are currently pulled in with plate details.
                                'While user can specify multiple stiffener groups, CCIplate only reports summary of controlling ratings. 
                            End If
                        Next

                    End If
                Next

                'Bolts
                For Each bgrow As DataRow In ds.Tables(plBoltGroup.EDSObjectName).Rows
                    plBoltGroup = New BoltGroup(bgrow, EDStruefalse, Me)
                    If If(EDStruefalse, plBoltGroup.connection_id = plConnection.ID, plBoltGroup.local_connection_id = plConnection.local_id) Then
                        plConnection.BoltGroups.Add(plBoltGroup)
                        For Each bdrow As DataRow In ds.Tables(plBoltDetail.EDSObjectName).Rows
                            plBoltDetail = New BoltDetail(bdrow, EDStruefalse, Me)
                            'If If(EDStruefalse, plBoltDetail.bolt_id = plBoltGroup.ID, plBoltDetail.bolt_id = plBoltGroup.local_id) Then
                            'If If(EDStruefalse, plBoltDetail.bolt_id = plBoltGroup.ID, plBoltDetail.local_id = plConnection.local_id And plBoltDetail.bolt_id = plBoltGroup.local_id) Then
                            If If(EDStruefalse, plBoltDetail.bolt_group_id = plBoltGroup.ID, If(plBoltDetail.local_group_id > 0, plBoltDetail.local_connection_id = plConnection.local_id And plBoltDetail.local_group_id = plBoltGroup.local_id, plBoltDetail.local_connection_id = plConnection.local_id And plBoltDetail.bolt_group_id = plBoltGroup.ID)) Then
                                plBoltGroup.BoltDetails.Add(plBoltDetail)
                                For Each mrow As DataRow In ds.Tables(plCCIplateMaterial.EDSObjectName).Rows
                                    plCCIplateMaterial = New CCIplateMaterial(mrow, EDStruefalse, Me)
                                    If If(EDStruefalse, plCCIplateMaterial.ID = plBoltDetail.bolt_material, plCCIplateMaterial.local_id = plBoltDetail.bolt_material) Then
                                        'plPlateDetail.plate_material = plConnectionMaterial
                                        plBoltDetail.CCIplateMaterials.Add(plCCIplateMaterial)
                                        Exit For 'Once matched, don't need to continue checking. 
                                    End If
                                Next
                            End If
                        Next
                        If IsSomething(ds.Tables("Bolt Results")) Then
                            For Each brrow As DataRow In ds.Tables("Bolt Results").Rows
                                plBoltResult = New BoltResults(brrow, EDStruefalse, Me)
                                If If(EDStruefalse, False, plBoltResult.local_connection_id = plConnection.local_id And plBoltResult.local_bolt_group_id = plBoltGroup.local_id) Then
                                    plBoltGroup.BoltResults.Add(plBoltResult)
                                End If
                            Next
                        End If
                    End If
                Next

                'Bridge Stiffeners
                For Each bsdrow As DataRow In ds.Tables(plBridgeDetail.EDSObjectName).Rows
                    plBridgeDetail = New BridgeStiffenerDetail(bsdrow, EDStruefalse, Me)
                    If If(EDStruefalse, plBridgeDetail.connection_id = plConnection.ID, If(plBridgeDetail.local_connection_id > 0, plBridgeDetail.local_connection_id = plConnection.local_id, plBridgeDetail.connection_id = plConnection.ID)) Then
                        plConnection.BridgeStiffenerDetails.Add(plBridgeDetail)
                        For Each mrow As DataRow In ds.Tables(plCCIplateMaterial.EDSObjectName).Rows
                            plCCIplateMaterial = New CCIplateMaterial(mrow, EDStruefalse, Me)
                            If If(EDStruefalse, plCCIplateMaterial.ID = plBridgeDetail.bridge_stiffener_material, plCCIplateMaterial.local_id = plBridgeDetail.bridge_stiffener_material) Then
                                plBridgeDetail.CCIplateMaterials.Add(plCCIplateMaterial)
                                Exit For 'Once matched, don't need to continue checking. 
                            End If
                        Next
                    End If
                Next

                'BARB & Bridge Stiffener Results
                If IsSomething(ds.Tables("Connection Results")) Then
                    For Each crrow As DataRow In ds.Tables("Connection Results").Rows
                        plConnectionResult = New ConnectionResults(crrow, EDStruefalse, Me)
                        If If(EDStruefalse, False, plConnectionResult.local_connection_id = plConnection.local_id) Then
                            plConnection.ConnectionResults.Add(plConnectionResult)
                        End If
                    Next
                End If

            End If
        Next

        'End Function
    End Sub

#End Region

#Region "Load From Excel"
    Public Overrides Sub LoadFromExcel()
        Dim excelDS As New DataSet

        For Each item As EXCELDTParameter In ExcelDTParams
            'Get additional tables from excel file 
            Try
                excelDS.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(Me.WorkBookPath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
            Catch ex As Exception
                Debug.Print(String.Format("Failed to create datatable for: {0}, {1}, {2}", IO.Path.GetFileName(Me.WorkBookPath), item.xlsSheet, item.xlsRange))
            End Try
        Next

        'Specific requirements for setting CCIplate properties and properties of its children
        If excelDS.Tables.Contains("CCIplates") Then
            Dim dr = excelDS.Tables("CCIplates").Rows(0)

            'Following used to create dataset, regardless if source was EDS or Excel. Boolean used to identify source. Excel = False
            BuildFromDataset(dr, excelDS, False, Me)

        End If
    End Sub
#End Region

#Region "Save to Excel"

    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        '''''Customize for each excel tool'''''

        'Site Code Criteria
        Dim tia_current, site_name, structure_type As String
        Dim rev_h_section_15_5 As Boolean?
        Dim site_app, site_rev As Integer?

        With wb
            'Site Code Criteria
            'Site Name
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.site_name) Then
                site_name = Me.ParentStructure?.structureCodeCriteria?.site_name
                .Worksheets("Main").Range("C4").Value = CType(site_name, String)
            End If

            'App ID & Revision #
            .Worksheets("Main").Range("C5").Value = MyOrder()
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.eng_app_id) Then
            '    site_app = Me.ParentStructure?.structureCodeCriteria?.eng_app_id
            '    'Revision #
            '    If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.eng_app_id_revision) Then
            '        site_rev = Me.ParentStructure?.structureCodeCriteria?.eng_app_id_revision
            '        'fields are combined into 1 cell within CCIplate
            '        .Worksheets("Main").Range("C5").Value = CType(site_app, String) & " REV. " & CType(site_rev, String)
            '    Else
            '        .Worksheets("Main").Range("C5").Value = CType(site_app, String)
            '    End If
            'End If

            'Tower Type - Defaulting to Monopole if not one of the main tower types
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.structure_type) Then
            If Me.ParentStructure?.structureCodeCriteria?.structure_type = "SELF SUPPORT" Then
                structure_type = "Self Support"
            ElseIf Me.ParentStructure?.structureCodeCriteria?.structure_type = "MONOPOLE" Then
                structure_type = "Monopole"
            Else
                structure_type = "Monopole"
            End If
            .Worksheets("Main").Range("tower_type").Value = CType(structure_type, String)
            'End If

            'TIA Revision
            .Worksheets("Main").Range("C9").Value = MyTIA()

            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.tia_current) Then
            '    If Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-F" Then
            '        tia_current = "F"
            '    ElseIf Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-G" Then
            '        tia_current = "G"
            '    ElseIf Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-H" Then
            '        tia_current = "H"
            '    Else
            '        tia_current = "H"
            '    End If
            '    .Worksheets("Main").Range("C9").Value = CType(tia_current, String)
            'End If

            'H Section 15.5
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5) Then
                rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
                .Worksheets("Engine").Range("D5").Value = CType(rev_h_section_15_5, Boolean)
            End If
            'Load Z Normalization
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.load_z_norm) Then
            '    rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.load_z_norm
            '    .Worksheets("Engine").Range("G10").Value = CType(load_z_norm, Boolean)
            'End If

            'Loading
            'If structure_type = "Self-Suppot" Then
            '    .Worksheets("Input").Range("D13").Value = CType(uplift, Double)
            '    .Worksheets("Input").Range("D14").Value = CType(compression, Double)
            '    .Worksheets("Input").Range("D15").Value = CType(uplift_shear, Double)
            '    .Worksheets("Input").Range("D16").Value = CType(compression_shear, Double)
            'Else
            '    .Worksheets("Input").Range("D13").Value = CType(moment, Double)
            '    .Worksheets("Input").Range("D14").Value = CType(axial, Double)
            '    .Worksheets("Input").Range("D15").Value = CType(shear, Double)
            'End If

            .Worksheets("Sub Tables (SAPI)").Range("A3").Value = CType(True, Boolean) 'Flags if sheet was last touched by EDS. If true, worksheet change event upon opening tool. 

            .Worksheets("Sub Tables (SAPI)").Range("ID").Value = CType(Me.ID, Integer)
            'If Not IsNothing(Me.ID) Then
            '    .Worksheets("Sub Tables (SAPI)").Range("ID").Value = CType(Me.ID, Integer)
            'Else
            '    .Worksheets("Sub Tables (SAPI)").Range("ID").ClearContents
            'End If
            If Not IsNothing(Me.bus_unit) Then
                .Worksheets("Main").Range("C3").Value = CType(Me.bus_unit, String)
            End If
            'If Not IsNothing(Me.structure_id) Then
            '    .Worksheets("").Range("").Value = CType(Me.structure_id, String)
            'End If
            If Not IsNothing(Me.anchor_rod_spacing) Then
                .Worksheets("MP Connection Summary").Range("rod_spacing").Value = CType(Me.anchor_rod_spacing, Double)
            Else
                .Worksheets("MP Connection Summary").Range("rod_spacing").ClearContents
            End If
            If Not IsNothing(Me.clip_distance) Then
                .Worksheets("MP Connection Summary").Range("clip").Value = CType(Me.clip_distance, Double)
            Else
                .Worksheets("MP Connection Summary").Range("clip").ClearContents
            End If
            'If Not IsNothing(Me.barb_cl_elevation) Then
            '    .Worksheets("Custom Connection").Range("H8").Value = CType(Me.barb_cl_elevation, Double)
            'Else
            '    .Worksheets("Custom Connection").Range("H8").ClearContents
            'End If
            If Not IsNothing(Me.barb_cl_elevation) Then
                .Worksheets("Database").Range("G3833").Value = CType(Me.barb_cl_elevation, Double)
            Else
                .Worksheets("Database").Range("G3833").ClearContents
            End If
            If Not IsNothing(Me.include_pole_reactions) Then
                .Worksheets("BARB").Range("X2").Value = CType(Me.include_pole_reactions, Boolean)
            End If
            If Not IsNothing(Me.consider_ar_eccentricity) Then
                .Worksheets("Engine").Range("D17").Value = CType(Me.consider_ar_eccentricity, Boolean)
            End If
            If Not IsNothing(Me.leg_mod_eccentricity) Then
                .Worksheets("Custom Connection").Range("J8").Value = CType(Me.leg_mod_eccentricity, Double)
                If Me.leg_mod_eccentricity <> 0 Then
                    .Worksheets("Main").Range("D247").Value = "Yes"
                Else
                    .Worksheets("Main").Range("D247").Value = "No"
                End If
            Else
                .Worksheets("Custom Connection").Range("J8").ClearContents
            End If
            If Not IsNothing(Me.seismic) Then
                .Worksheets("Main").Range("J11").Value = CType(Me.seismic, Boolean)
            End If
            If Not IsNothing(Me.seismic_flanges) Then
                .Worksheets("Main").Range("K11").Value = CType(Me.seismic_flanges, Boolean)
            End If

            If Not IsNothing(Me.Structural_105) Then
                .Worksheets("Engine").Range("D5").Value = CType(Me.Structural_105, Boolean)
            End If
            'If Not IsNothing(Me.tool_version) Then
            '    .Worksheets("Revision History").Range("Revision").Value = CType(Me.tool_version, String)
            'End If
            'If Not IsNothing(Me.modified_person_id) Then
            '    .Worksheets("").Range("").Value = CType(Me.modified_person_id, Integer)
            'Else
            '    .Worksheets("").Range("").ClearContents
            'End If
            'If Not IsNothing(Me.process_stage) Then
            '    .Worksheets("").Range("").Value = CType(Me.process_stage, String)
            'End If

            If Me.Connections.Count > 0 Then
                'identify first row to copy data into Excel Sheet
                'Connection
                Dim PlateRow As Integer = 3 'SAPI Tab
                Dim PlateRow2 As Integer = 46 'MP Connection Summary Tab
                'Dim PlateRow3 As Integer = 19 'Main Tab (currently not used)
                'Plate Details
                Dim PlateDRow As Integer = 3 'SAPI Tab
                Dim i As Integer 'Row adjustment for top vs bottom plate
                'Materials
                Dim MatRow As Integer = 40 'Materials Tab; SAPI Tab - 2
                Dim tempMaterials As New List(Of CCIplateMaterial)
                Dim tempMaterial As New CCIplateMaterial 'Temp material object to determine if already added to Excel
                Dim matflag As Boolean = False 'determines whether or not to add to Excel based on temp list
                'Excel Database Reference
                Dim mycol As Integer = 6 'Bolt Group, Bolt Details, Stiffener Details
                Dim bump As Integer = 0 'Bolt Group
                Dim bump2 As Integer = 0 'Bolt Detail
                'Stiffener Group
                Dim StiffGRow As Integer = 3 'SAPI Tab
                'Dim StiffGRow2 As Integer = 10 'MP Connection Summary TAB
                'Stiffener Details
                Dim StiffDRow As Integer = 3 'SAPI Tab
                Dim StiffDRow2 As Integer = 85 ' MP Connection Summary Tab
                'Bridge Stiffener Details
                Dim BridgeDRow As Integer = 3 'SAPI Tab
                'Dim BridgeDRow2 As Integer = 167 'MP Connection Summary Tab

                For Each row As Connection In Connections

                    'Excel Database Reference (resets for each plate connection)
                    Dim myrow4 As Integer '= 1027 'Stiffener Details

                    If Not IsNothing(row.ID) Then
                        .Worksheets("Sub Tables (SAPI)").Range("D" & PlateRow).Value = CType(row.ID, Integer)
                    End If
                    If structure_type = "Self Support" Then
                        If Not IsNothing(row.connection_elevation) Then
                            .Worksheets("Custom Connection").Range("elevation").Value = CType(row.connection_elevation, Double)
                        End If
                        If Not IsNothing(row.bolt_configuration) Then
                            .Worksheets("Main").Range("D246").Value = CType(row.bolt_configuration, String)
                        Else
                            .Worksheets("Main").Range("D246").ClearContents
                        End If
                    ElseIf structure_type = "Monopole" Then
                        If Not IsNothing(row.connection_elevation) Then
                            .Worksheets("MP Connection Summary").Range("C" & PlateRow2).Value = CType(row.connection_elevation, Double)
                        End If
                        If Not IsNothing(row.bolt_configuration) Then
                            .Worksheets("MP Connection Summary").Range("K" & PlateRow2).Value = CType(row.bolt_configuration, String)
                        Else
                            .Worksheets("MP Connection Summary").Range("K" & PlateRow2).ClearContents
                        End If
                        'Need to report flags for proper worksheet change events
                        If row.bolt_configuration = "Custom" And row.connection_type = "Base" Then
                            .Worksheets("MP Connection Summary").Range("U" & PlateRow2).Value = CType(1, Integer)
                        ElseIf row.bolt_configuration = "Custom" And row.connection_type = "Flange" Then
                            .Worksheets("MP Connection Summary").Range("U" & PlateRow2).Value = CType(1, Integer)
                            .Worksheets("MP Connection Summary").Range("U" & PlateRow2 + 1).Value = CType(1, Integer)
                        End If
                    End If
                    'If Not IsNothing(row.connection_type) Then 'do not need, tool will autopopulate
                    '    .Worksheets("").Range("").Value = CType(row.connection_type, String)
                    'End If

                    'For i = 1 To 200 '200 possilbe rows for Pole Geometry (Need to figure out how to add. Need to reference pole geometry)
                    '    'If row.connection_elevation = .Worksheets("Main").Range("B" & PlateRow3).Value Then
                    '    If Not IsNothing(row.connection_type) Then
                    '        .Worksheets("Main").Range("B" & PlateRow3).Value = CType(row.connection_type, String)
                    '    Else
                    '        .Worksheets("Main").Range("B" & PlateRow3).ClearContents
                    '    End If
                    '    'End If
                    'Next i

                    'If Not IsNothing(row.bolt_configuration) Then
                    '    If structure_type = "Self Support" Then
                    '        .Worksheets("Main").Range("D246").Value = CType(row.bolt_configuration, String)
                    '    ElseIf structure_type = "Monopole" Then
                    '        .Worksheets("MP Connection Summary").Range("K" & PlateRow2).Value = CType(row.bolt_configuration, String)
                    '        'Need to report flags for proper worksheet change events
                    '        If row.bolt_configuration = "Custom" And row.connection_type = "Base" Then
                    '            .Worksheets("MP Connection Summary").Range("U" & PlateRow2).Value = CType(1, Integer)
                    '        ElseIf row.bolt_configuration = "Custom" And row.connection_type = "Flange" Then
                    '            .Worksheets("MP Connection Summary").Range("U" & PlateRow2).Value = CType(1, Integer)
                    '            .Worksheets("MP Connection Summary").Range("U" & PlateRow2 + 1).Value = CType(1, Integer)
                    '        End If
                    '    End If
                    'End If

                    For Each pdrow As PlateDetail In row.PlateDetails



                        If pdrow.connection_id = row.ID Then

                            If pdrow.plate_location = "Bottom" Then
                                i = 1
                                myrow4 = 2427
                            Else
                                i = 0
                                myrow4 = 1027
                            End If

                            If Not IsNothing(pdrow.ID) Then
                                .Worksheets("Sub Tables (SAPI)").Range("K" & PlateDRow).Value = CType(pdrow.ID, Integer)
                            End If
                            'If Not IsNothing(row.plate_location) Then 'do not need, tool will autopopulate
                            '    .Worksheets("MP Connection Summary").Range("D" & PlateRow2).Value = CType(row.plate_location, String)
                            'End If
                            If Not IsNothing(pdrow.plate_type) Then
                                .Worksheets("MP Connection Summary").Range("E" & PlateRow2 + i).Value = CType(pdrow.plate_type, String)
                            End If
                            If Not IsNothing(pdrow.plate_diameter) Then
                                .Worksheets("MP Connection Summary").Range("F" & PlateRow2 + i).Value = CType(pdrow.plate_diameter, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("F" & PlateRow2 + i).ClearContents
                            End If
                            If Not IsNothing(pdrow.plate_thickness) Then
                                .Worksheets("MP Connection Summary").Range("G" & PlateRow2 + i).Value = CType(pdrow.plate_thickness, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("G" & PlateRow2 + i).ClearContents
                            End If
                            'If Not IsNothing(pdrow.plate_material) Then
                            '    .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).Value = CType(pdrow.plate_material, Integer)
                            'Else
                            '    .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).ClearContents
                            'End If
                            For Each mrow As CCIplateMaterial In pdrow.CCIplateMaterials
                                If mrow.default_material = True Then
                                    If Not IsNothing(mrow.name) Then
                                        .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).Value = CType(mrow.name, String)
                                    Else
                                        .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).ClearContents
                                    End If
                                Else 'After adding new materail, save material name in a list to reference for other plates to see if materail was already added. 
                                    For Each tmrow In tempMaterials
                                        If mrow.ID = tmrow.ID Then
                                            matflag = True 'don't add to excel
                                            Exit For
                                        End If
                                    Next
                                    If matflag = False Then
                                        'tempMaterial = New CCIplateMaterial(mrow.ID)
                                        tempMaterial = New CCIplateMaterial(mrow.ID, mrow.name, mrow.fy_0, mrow.fu_0)
                                        tempMaterials.Add(tempMaterial)

                                        '.Worksheets("Sub Tables (SAPI)").Range("AR").Count
                                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Columns("AR").Count 'counts total columns in excel
                                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Cells("AR:AR").Count
                                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").GetDataRange("AR:AR")

                                        'If Not IsNothing(mrow.ID) Then
                                        '    .Worksheets("Sub Tables (SAPI)").Range("AR" & MatRow2).Value = CType(mrow.ID, Integer)
                                        'End If
                                        If Not IsNothing(mrow.name) Then
                                            .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).Value = CType(mrow.name, String)
                                        Else
                                            .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).ClearContents
                                        End If

                                        SaveMaterial(wb, mrow, MatRow)
                                        MatRow += 1
                                    Else
                                        If Not IsNothing(mrow.name) Then
                                            .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).Value = CType(mrow.name, String)
                                        Else
                                            .Worksheets("MP Connection Summary").Range("H" & PlateRow2 + i).ClearContents
                                        End If
                                    End If
                                End If
                                matflag = False 'reset flag
                            Next

                            If Not IsNothing(pdrow.stiffener_configuration) Then
                                If pdrow.stiffener_configuration = 4 Then
                                    .Worksheets("MP Connection Summary").Range("I" & PlateRow2 + i).Value = "Custom"
                                    'Need to report flags for proper worksheet change events
                                    .Worksheets("MP Connection Summary").Range("T" & PlateRow2).Value = CType(1, Integer)
                                Else
                                    .Worksheets("MP Connection Summary").Range("I" & PlateRow2 + i).Value = CType(pdrow.stiffener_configuration, Integer)
                                End If
                            Else
                                .Worksheets("MP Connection Summary").Range("I" & PlateRow2 + i).ClearContents
                            End If
                            If Not IsNothing(pdrow.stiffener_clear_space) Then
                                .Worksheets("MP Connection Summary").Range("D" & PlateRow2 + i + 39).Value = CType(pdrow.stiffener_clear_space, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("D" & PlateRow2 + i + 39).ClearContents
                            End If
                            If Not IsNothing(pdrow.plate_check) Then
                                If pdrow.plate_check = True Then
                                    .Worksheets("MP Connection Summary").Range("J" & PlateRow2 + i).Value = "Yes"
                                Else
                                    .Worksheets("MP Connection Summary").Range("J" & PlateRow2 + i).Value = "No"
                                End If
                            End If

                            Dim sgid As Integer = 1 'Stiffener Group names per CCIplate are integers 1-5
                            Dim FirstRowStiff As Boolean = True
                            'Stiffener Group
                            StiffGRow = 3 'SAPI Tab -reset for each plate detail
                            'Stiffener Details
                            StiffDRow = 3 'SAPI Tab -reset for each plate detail

                            For Each sgrow As StiffenerGroup In pdrow.StiffenerGroups
                                If sgrow.plate_details_id = pdrow.ID Then

                                    If Not IsNothing(sgrow.ID) Then
                                        .Worksheets("Sub Tables (SAPI)").Range("BI" & StiffGRow + (PlateDRow - 3) * 5).Value = CType(sgrow.ID, Integer)
                                    End If
                                    'If Not IsNothing(Me.stiffener_name) Then
                                    '    .Worksheets("Database").Range("G").Value = CType(Me.stiffener_name, String)
                                    'End If

                                    For Each sdrow As StiffenerDetail In sgrow.StiffenerDetails
                                        If sdrow.stiffener_id = sgrow.ID Then

                                            'Save stiffener data to MP Connection Summary when connection is symmetrical
                                            If FirstRowStiff And pdrow.stiffener_configuration > 0 And pdrow.stiffener_configuration <> 4 Then
                                                'If Not IsNothing(sdrow.stiffener_location) Then
                                                '    .Worksheets("Database").Cells(myrow4 + 1, mycol).Value = CType(sdrow.stiffener_location, Double)
                                                'End If
                                                If Not IsNothing(sdrow.stiffener_width) Then
                                                    .Worksheets("MP Connection Summary").Range("E" & StiffDRow2).Value = CType(sdrow.stiffener_width, Double)
                                                End If
                                                If Not IsNothing(sdrow.stiffener_height) Then
                                                    .Worksheets("MP Connection Summary").Range("F" & StiffDRow2).Value = CType(sdrow.stiffener_height, Double)
                                                End If
                                                If Not IsNothing(sdrow.stiffener_thickness) Then
                                                    .Worksheets("MP Connection Summary").Range("G" & StiffDRow2).Value = CType(sdrow.stiffener_thickness, Double)
                                                End If
                                                If Not IsNothing(sdrow.stiffener_h_notch) Then
                                                    .Worksheets("MP Connection Summary").Range("H" & StiffDRow2).Value = CType(sdrow.stiffener_h_notch, Double)
                                                End If
                                                If Not IsNothing(sdrow.stiffener_v_notch) Then
                                                    .Worksheets("MP Connection Summary").Range("I" & StiffDRow2).Value = CType(sdrow.stiffener_v_notch, Double)
                                                End If
                                                If Not IsNothing(sdrow.stiffener_grade) Then
                                                    .Worksheets("MP Connection Summary").Range("J" & StiffDRow2).Value = CType(sdrow.stiffener_grade, Double)
                                                End If
                                                If Not IsNothing(sdrow.weld_type) Then
                                                    .Worksheets("MP Connection Summary").Range("K" & StiffDRow2).Value = CType(sdrow.weld_type, String)
                                                End If
                                                If Not IsNothing(sdrow.groove_depth) Then
                                                    .Worksheets("MP Connection Summary").Range("L" & StiffDRow2).Value = CType(sdrow.groove_depth, Double)
                                                End If
                                                If Not IsNothing(sdrow.groove_angle) Then
                                                    .Worksheets("MP Connection Summary").Range("M" & StiffDRow2).Value = CType(sdrow.groove_angle, Double)
                                                End If
                                                If Not IsNothing(sdrow.h_fillet_weld) Then
                                                    .Worksheets("MP Connection Summary").Range("N" & StiffDRow2).Value = CType(sdrow.h_fillet_weld, Double)
                                                End If
                                                If Not IsNothing(sdrow.v_fillet_weld) Then
                                                    .Worksheets("MP Connection Summary").Range("O" & StiffDRow2).Value = CType(sdrow.v_fillet_weld, Double)
                                                End If
                                                If Not IsNothing(sdrow.weld_strength) Then
                                                    .Worksheets("MP Connection Summary").Range("P" & StiffDRow2).Value = CType(sdrow.weld_strength, Double)
                                                End If

                                                FirstRowStiff = False
                                            End If


                                            If Not IsNothing(sdrow.ID) Then
                                                .Worksheets("Sub Tables (SAPI)").Range("BN" & StiffDRow + (PlateDRow - 3) * 100).Value = CType(sdrow.ID, Integer)
                                            Else
                                                .Worksheets("Sub Tables (SAPI)").Range("BN" & StiffDRow + (PlateDRow - 3) * 100).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.stiffener_id) Then
                                                .Worksheets("Sub Tables (SAPI)").Range("BO" & StiffDRow + (PlateDRow - 3) * 100).Value = CType(sdrow.stiffener_id, Integer)
                                                .Worksheets("Database").Cells(myrow4, mycol).Value = CType(sgid, Double)
                                            Else
                                                .Worksheets("Sub Tables (SAPI)").Range("BO" & StiffDRow + (PlateDRow - 3) * 100).ClearContents
                                                .Worksheets("Database").Cells(myrow4, mycol).ClearContents
                                            End If



                                            If Not IsNothing(sdrow.stiffener_location) Then
                                                .Worksheets("Database").Cells(myrow4 + 1, mycol).Value = CType(sdrow.stiffener_location, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 1, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.stiffener_width) Then
                                                .Worksheets("Database").Cells(myrow4 + 2, mycol).Value = CType(sdrow.stiffener_width, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 2, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.stiffener_height) Then
                                                .Worksheets("Database").Cells(myrow4 + 3, mycol).Value = CType(sdrow.stiffener_height, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 3, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.stiffener_thickness) Then
                                                .Worksheets("Database").Cells(myrow4 + 4, mycol).Value = CType(sdrow.stiffener_thickness, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 4, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.stiffener_h_notch) Then
                                                .Worksheets("Database").Cells(myrow4 + 5, mycol).Value = CType(sdrow.stiffener_h_notch, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 5, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.stiffener_v_notch) Then
                                                .Worksheets("Database").Cells(myrow4 + 6, mycol).Value = CType(sdrow.stiffener_v_notch, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 6, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.stiffener_grade) Then
                                                .Worksheets("Database").Cells(myrow4 + 7, mycol).Value = CType(sdrow.stiffener_grade, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 7, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.weld_type) Then
                                                .Worksheets("Database").Cells(myrow4 + 8, mycol).Value = CType(sdrow.weld_type, String)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 8, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.groove_depth) Then
                                                .Worksheets("Database").Cells(myrow4 + 9, mycol).Value = CType(sdrow.groove_depth, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 9, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.groove_angle) Then
                                                .Worksheets("Database").Cells(myrow4 + 10, mycol).Value = CType(sdrow.groove_angle, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 10, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.h_fillet_weld) Then
                                                .Worksheets("Database").Cells(myrow4 + 11, mycol).Value = CType(sdrow.h_fillet_weld, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 11, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.v_fillet_weld) Then
                                                .Worksheets("Database").Cells(myrow4 + 12, mycol).Value = CType(sdrow.v_fillet_weld, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 12, mycol).ClearContents
                                            End If
                                            If Not IsNothing(sdrow.weld_strength) Then
                                                .Worksheets("Database").Cells(myrow4 + 13, mycol).Value = CType(sdrow.weld_strength, Double)
                                            Else
                                                .Worksheets("Database").Cells(myrow4 + 13, mycol).ClearContents
                                            End If

                                            myrow4 += 14
                                            StiffDRow += 1
                                            'FirstRow = False 'Turns off saving bolt information to MP Connection Summary Tab if not first row

                                        End If
                                    Next
                                    sgid += 1
                                    StiffGRow += 1
                                End If
                            Next


                            PlateDRow += 1

                        End If
                        StiffDRow2 -= 1
                    Next

                    'Excel Database Reference (resets for each plate connection)
                    Dim myrow As Integer = 7 'Bolt Group & Bolt Details
                    Dim bgid As Integer = 1 'Bolt Group names per CCIplate are integers 1-5
                    'Bolt Group
                    Dim myrow3 As Integer = 3827 'Bolt Group BARB Elevation
                    Dim BoltGRow As Integer = 3 'SAPI Tab
                    'Bolt Detail
                    Dim myrow2 As Integer = 27 'Excel Database
                    Dim BoltDRow As Integer = 3 'SAPI Tab
                    Dim FirstRow As Boolean = True 'Apllies for only 1 bolt group that is symmetrical. Saves data to MP Connection Summary Tab

                    For Each bgrow As BoltGroup In row.BoltGroups
                        If bgrow.connection_id = row.ID Then

                            If FirstRow And row.bolt_configuration = "Symmetrical" Then
                                If structure_type = "Self Support" Then
                                    If Not IsNothing(bgrow.grout_considered) Then
                                        If bgrow.grout_considered = True Then
                                            .Worksheets("Main").Range("D251").Value = "Yes"
                                        Else
                                            .Worksheets("Main").Range("D251").Value = "No"
                                        End If
                                    End If
                                ElseIf structure_type = "Monopole" And row.connection_type = "Base" Then
                                    If Not IsNothing(bgrow.grout_considered) Then
                                        If bgrow.grout_considered = True Then
                                            .Worksheets("MP Connection Summary").Range("N9").Value = "Yes"
                                        Else
                                            .Worksheets("MP Connection Summary").Range("N9").Value = "No"
                                        End If
                                    End If
                                End If
                            ElseIf FirstRow And row.bolt_configuration = "Custom" Then
                                If structure_type = "Self Support" Then
                                    .Worksheets("Main").Range("D251").ClearContents
                                End If
                            End If

                            If Not IsNothing(bgrow.ID) Then
                                .Worksheets("Sub Tables (SAPI)").Range("W" & BoltGRow + bump).Value = CType(bgrow.ID, Integer)
                            End If

                            If Not IsNothing(bgrow.resist_axial) Then
                                If bgrow.resist_axial = True Then
                                    .Worksheets("Database").Cells(myrow, mycol).Value = "Yes"
                                Else
                                    .Worksheets("Database").Cells(myrow, mycol).Value = "No"
                                End If
                            End If
                            If Not IsNothing(bgrow.resist_shear) Then
                                If bgrow.resist_shear = True Then
                                    .Worksheets("Database").Cells(myrow + 1, mycol).Value = "Yes"
                                Else
                                    .Worksheets("Database").Cells(myrow + 1, mycol).Value = "No"
                                End If
                            End If
                            If Not IsNothing(bgrow.plate_bending) Then
                                If bgrow.plate_bending = True Then
                                    .Worksheets("Database").Cells(myrow + 2, mycol).Value = "Yes"
                                Else
                                    .Worksheets("Database").Cells(myrow + 2, mycol).Value = "No"
                                End If
                            End If
                            If Not IsNothing(bgrow.grout_considered) Then
                                If bgrow.grout_considered = True Then
                                    .Worksheets("Database").Cells(myrow + 3, mycol).Value = "Yes"
                                Else
                                    .Worksheets("Database").Cells(myrow + 3, mycol).Value = "No"
                                End If
                            End If
                            If Not IsNothing(bgrow.apply_barb_elevation) Then
                                If bgrow.apply_barb_elevation = True Then
                                    .Worksheets("Database").Cells(myrow3, mycol).Value = "Yes"
                                Else
                                    .Worksheets("Database").Cells(myrow3, mycol).Value = "No"
                                End If
                            End If
                            'Bolt Group names in CCIplate are named 1 through 5. 
                            'If Not IsNothing(bgrow.bolt_name) Then
                            '    .Worksheets("Database").Cells(myrow + 5, mycol).Value = CType(bgrow.bolt_name, String)
                            'End If

                            For Each bdrow As BoltDetail In bgrow.BoltDetails
                                If bdrow.bolt_group_id = bgrow.ID Then

                                    'Save bolt data to MP Connection Summary when connection is symmetrical
                                    If FirstRow And row.bolt_configuration = "Symmetrical" Then
                                        If structure_type = "Self Support" Then
                                            If Not IsNothing(bgrow.BoltDetails.Count) Then
                                                .Worksheets("Main").Range("D248").Value = CType(bgrow.BoltDetails.Count, Integer)
                                            End If
                                            If Not IsNothing(bdrow.bolt_diameter) Then
                                                .Worksheets("Main").Range("D249").Value = CType(bdrow.bolt_diameter, Double)
                                            End If
                                            '.Worksheets("MP Connection Summary").Range("N" & PlateRow2).Value = CType(bdrow.bolt_material, String)
                                            If Not IsNothing(bdrow.bolt_thread_type) Then
                                                .Worksheets("Main").Range("D254").Value = CType(bdrow.bolt_thread_type, String)
                                            End If
                                            'If Not IsNothing(bdrow.bolt_circle) Then
                                            '    .Worksheets("Main").Range("P" & PlateRow2).Value = CType(bdrow.bolt_circle, Double)
                                            'End If

                                            If Not IsNothing(bdrow.eta_factor) Then
                                                .Worksheets("Main").Range("D253").Value = CType(bdrow.eta_factor, Double)
                                            End If
                                            If Not IsNothing(bdrow.lar) Then
                                                .Worksheets("Main").Range("D252").Value = CType(bdrow.lar, Double)
                                            End If

                                            'Need to store elevation also within Database Tab 
                                            If Not IsNothing(row.connection_elevation) Then
                                                .Worksheets("Database").Cells(myrow2 - 24, mycol).Value = CType(row.connection_elevation, Double)
                                            End If
                                        ElseIf structure_type = "Monopole" Then
                                            If Not IsNothing(bgrow.BoltDetails.Count) Then
                                                .Worksheets("MP Connection Summary").Range("L" & PlateRow2).Value = CType(bgrow.BoltDetails.Count, Integer)
                                            End If
                                            If Not IsNothing(bdrow.bolt_diameter) Then
                                                .Worksheets("MP Connection Summary").Range("M" & PlateRow2).Value = CType(bdrow.bolt_diameter, Double)
                                            End If
                                            '.Worksheets("MP Connection Summary").Range("N" & PlateRow2).Value = CType(bdrow.bolt_material, String)
                                            If Not IsNothing(bdrow.bolt_thread_type) Then
                                                .Worksheets("MP Connection Summary").Range("O" & PlateRow2).Value = CType(bdrow.bolt_thread_type, String)
                                            End If
                                            If Not IsNothing(bdrow.bolt_circle) Then
                                                .Worksheets("MP Connection Summary").Range("P" & PlateRow2).Value = CType(bdrow.bolt_circle, Double)
                                            End If
                                            If row.connection_type = "Base" Then
                                                If Not IsNothing(bdrow.eta_factor) Then
                                                    .Worksheets("MP Connection Summary").Range("N11").Value = CType(bdrow.eta_factor, Double)
                                                End If
                                                If Not IsNothing(bdrow.lar) Then
                                                    .Worksheets("MP Connection Summary").Range("N10").Value = CType(bdrow.lar, Double)
                                                End If
                                            End If
                                            'Need to store elevation also within Database Tab 
                                            If Not IsNothing(row.connection_elevation) Then
                                                .Worksheets("Database").Cells(myrow2 - 24, mycol).Value = CType(row.connection_elevation, Double)
                                            End If
                                        End If
                                    ElseIf FirstRow And row.bolt_configuration = "Custom" Then
                                        If structure_type = "Self Support" Then
                                            .Worksheets("Main").Range("D253").ClearContents
                                            .Worksheets("Main").Range("D252").ClearContents
                                            .Worksheets("Main").Range("D254").ClearContents
                                        End If
                                    End If


                                    If Not IsNothing(bdrow.ID) Then
                                        .Worksheets("Sub Tables (SAPI)").Range("AG" & BoltDRow + bump2).Value = CType(bdrow.ID, Integer)
                                    End If
                                    If Not IsNothing(bdrow.bolt_group_id) Then
                                        .Worksheets("Sub Tables (SAPI)").Range("AH" & BoltDRow + bump2).Value = CType(bdrow.bolt_group_id, Integer)
                                    End If

                                    'If Not IsNothing(bdrow.bolt_id) Then
                                    '    .Worksheets("Database").Cells(myrow, mycol).Value = CType(bdrow.bolt_id, Integer)
                                    'End If
                                    .Worksheets("Database").Cells(myrow2, mycol).Value = CType(bgid, Integer)
                                    If Not IsNothing(bdrow.bolt_location) Then
                                        .Worksheets("Database").Cells(myrow2 + 4, mycol).Value = CType(bdrow.bolt_location, Double)
                                    End If
                                    If Not IsNothing(bdrow.bolt_diameter) Then
                                        .Worksheets("Database").Cells(myrow2 + 1, mycol).Value = CType(bdrow.bolt_diameter, Double)
                                    End If
                                    'If Not IsNothing(bdrow.bolt_material) Then
                                    '    .Worksheets("Database").Cells(myrow2 + 2, mycol).Value = CType(bdrow.bolt_material, Integer)
                                    'End If

                                    For Each mrow As CCIplateMaterial In bdrow.CCIplateMaterials

                                        If mrow.default_material = True Then
                                            If FirstRow And row.bolt_configuration = "Symmetrical" Then
                                                If structure_type = "Self Support" Then
                                                    If Not IsNothing(mrow.name) Then
                                                        .Worksheets("Main").Range("D250").Value = CType(mrow.name, String)
                                                    End If
                                                ElseIf structure_type = "Monopole" Then
                                                    If Not IsNothing(mrow.name) Then
                                                        .Worksheets("MP Connection Summary").Range("N" & PlateRow2).Value = CType(mrow.name, String)
                                                    End If
                                                End If
                                            End If
                                            If Not IsNothing(mrow.name) Then
                                                .Worksheets("Database").Cells(myrow2 + 2, mycol).Value = CType(mrow.name, String)
                                            End If
                                        Else 'After adding new materail, save material name in a list to reference for other plates to see if materail was already added. 
                                            For Each tmrow In tempMaterials
                                                If mrow.ID = tmrow.ID Then
                                                    matflag = True 'don't add to excel
                                                    Exit For
                                                End If
                                            Next
                                            If matflag = False Then
                                                tempMaterial = New CCIplateMaterial(mrow.ID, mrow.name, mrow.fy_0, mrow.fu_0)
                                                tempMaterials.Add(tempMaterial)

                                                ''.Worksheets("Sub Tables (SAPI)").Range("AR").Count
                                                'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Columns("AR").Count 'counts total columns in excel
                                                'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Cells("AR:AR").Count
                                                'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").GetDataRange("AR:AR")

                                                If FirstRow And row.bolt_configuration = "Symmetrical" Then
                                                    If structure_type = "Self Support" Then
                                                        If Not IsNothing(mrow.name) Then
                                                            .Worksheets("Main").Range("D250").Value = CType(mrow.name, String)
                                                        End If
                                                    ElseIf structure_type = "Monopole" Then
                                                        If Not IsNothing(mrow.name) Then
                                                            .Worksheets("MP Connection Summary").Range("N" & PlateRow2).Value = CType(mrow.name, String)
                                                        End If
                                                    End If
                                                End If
                                                If Not IsNothing(mrow.name) Then
                                                    .Worksheets("Database").Cells(myrow2 + 2, mycol).Value = CType(mrow.name, String)
                                                End If
                                                'planning to reference SaveMaterial here since variables will be similar between sources
                                                SaveMaterial(wb, mrow, MatRow)
                                                MatRow += 1
                                            Else
                                                'If Not IsNothing(mrow.name) Then
                                                '    .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).Value = CType(mrow.name, String)
                                                'Else
                                                '    .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).ClearContents
                                                'End If

                                                If FirstRow And row.bolt_configuration = "Symmetrical" Then
                                                    If structure_type = "Self Support" Then
                                                        If Not IsNothing(mrow.name) Then
                                                            .Worksheets("Main").Range("D250").Value = CType(mrow.name, String)
                                                        End If
                                                    ElseIf structure_type = "Monopole" Then
                                                        If Not IsNothing(mrow.name) Then
                                                            .Worksheets("MP Connection Summary").Range("N" & PlateRow2).Value = CType(mrow.name, String)
                                                        End If
                                                    End If
                                                End If
                                                If Not IsNothing(mrow.name) Then
                                                    .Worksheets("Database").Cells(myrow2 + 2, mycol).Value = CType(mrow.name, String)
                                                End If

                                            End If
                                        End If
                                        matflag = False 'reset material flag
                                    Next

                                    If Not IsNothing(bdrow.bolt_circle) Then
                                        .Worksheets("Database").Cells(myrow2 + 3, mycol).Value = CType(bdrow.bolt_circle, Double)
                                    End If
                                    If Not IsNothing(bdrow.eta_factor) Then
                                        .Worksheets("Database").Cells(myrow2 + 5, mycol).Value = CType(bdrow.eta_factor, Double)
                                    End If
                                    If Not IsNothing(bdrow.lar) Then
                                        .Worksheets("Database").Cells(myrow2 + 6, mycol).Value = CType(bdrow.lar, Double)
                                    End If
                                    If Not IsNothing(bdrow.bolt_thread_type) Then
                                        .Worksheets("Database").Cells(myrow2 + 7, mycol).Value = CType(bdrow.bolt_thread_type, String)
                                    End If
                                    If Not IsNothing(bdrow.area_override) Then
                                        .Worksheets("Database").Cells(myrow2 + 8, mycol).Value = CType(bdrow.area_override, Double)
                                    End If
                                    If Not IsNothing(bdrow.tension_only) Then
                                        If bdrow.tension_only = True Then
                                            .Worksheets("Database").Cells(myrow2 + 9, mycol).Value = "Yes"
                                        Else
                                            .Worksheets("Database").Cells(myrow2 + 9, mycol).Value = "No"
                                        End If
                                    End If
                                    myrow2 += 10
                                    BoltDRow += 1
                                    FirstRow = False 'Turns off saving bolt information to MP Connection Summary Tab if not first row
                                End If
                            Next

                            myrow += 4
                            myrow3 += 1
                            bgid += 1
                            BoltGRow += 1

                        End If

                    Next

                    For Each bsdrow As BridgeStiffenerDetail In row.BridgeStiffenerDetails
                        If bsdrow.connection_id = row.ID Then

                            If Not IsNothing(bsdrow.ID) Then
                                .Worksheets("Sub Tables (SAPI)").Range("CF" & BridgeDRow).Value = CType(bsdrow.ID, Integer)
                                .Worksheets("Sub Tables (SAPI)").Range("CG" & BridgeDRow).Value = CType(row.ID, Integer) 'connetion id req. when deleting
                            End If
                            If Not IsNothing(bsdrow.connection_id) Then
                                '.Worksheets("MP Connection Summary").Range("B" & BridgeDRow + 164).Value = CType(bsdrow.plate_id, Integer)
                                .Worksheets("MP Connection Summary").Range("B" & BridgeDRow + 164).Value = CType(row.connection_elevation, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("B" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.stiffener_type) Then
                                .Worksheets("MP Connection Summary").Range("C" & BridgeDRow + 164).Value = CType(bsdrow.stiffener_type, String)
                            End If
                            If Not IsNothing(bsdrow.analysis_type) Then
                                .Worksheets("MP Connection Summary").Range("D" & BridgeDRow + 164).Value = CType(bsdrow.analysis_type, String)
                            End If
                            If Not IsNothing(bsdrow.quantity) Then
                                .Worksheets("MP Connection Summary").Range("E" & BridgeDRow + 164).Value = CType(bsdrow.quantity, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("E" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.bridge_stiffener_width) Then
                                .Worksheets("MP Connection Summary").Range("F" & BridgeDRow + 164).Value = CType(bsdrow.bridge_stiffener_width, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("F" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.bridge_stiffener_thickness) Then
                                .Worksheets("MP Connection Summary").Range("G" & BridgeDRow + 164).Value = CType(bsdrow.bridge_stiffener_thickness, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("G" & BridgeDRow + 164).ClearContents
                            End If
                            'If Not IsNothing(bsdrow.bridge_stiffener_material) Then
                            '    .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).Value = CType(bsdrow.bridge_stiffener_material, Integer)
                            'Else
                            '    .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).ClearContents
                            'End If
                            For Each mrow As CCIplateMaterial In bsdrow.CCIplateMaterials
                                If mrow.default_material = True Then
                                    If Not IsNothing(mrow.name) Then
                                        .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).Value = CType(mrow.name, String)
                                    Else
                                        .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).ClearContents
                                    End If
                                Else 'After adding new materail, save material name in a list to reference for other plates to see if materail was already added. 
                                    For Each tmrow In tempMaterials
                                        If mrow.ID = tmrow.ID Then
                                            matflag = True 'don't add to excel
                                            Exit For
                                        End If
                                    Next
                                    If matflag = False Then
                                        tempMaterial = New CCIplateMaterial(mrow.ID, mrow.name, mrow.fy_0, mrow.fu_0)
                                        tempMaterials.Add(tempMaterial)

                                        '.Worksheets("Sub Tables (SAPI)").Range("AR").Count
                                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Columns("AR").Count 'counts total columns in excel
                                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Cells("AR:AR").Count
                                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").GetDataRange("AR:AR")

                                        'If Not IsNothing(mrow.ID) Then
                                        '    .Worksheets("Sub Tables (SAPI)").Range("AR" & MatRow2).Value = CType(mrow.ID, Integer)
                                        'End If
                                        If Not IsNothing(mrow.name) Then
                                            .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).Value = CType(mrow.name, String)
                                        Else
                                            .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).ClearContents
                                        End If

                                        SaveMaterial(wb, mrow, MatRow)
                                        MatRow += 1
                                    Else
                                        If Not IsNothing(mrow.name) Then
                                            .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).Value = CType(mrow.name, String)
                                        Else
                                            .Worksheets("MP Connection Summary").Range("H" & BridgeDRow + 164).ClearContents
                                        End If
                                    End If
                                End If
                                matflag = False 'reset flag
                            Next
                            If Not IsNothing(bsdrow.unbraced_length) Then
                                .Worksheets("MP Connection Summary").Range("I" & BridgeDRow + 164).Value = CType(bsdrow.unbraced_length, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("I" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.total_length) Then
                                .Worksheets("MP Connection Summary").Range("J" & BridgeDRow + 164).Value = CType(bsdrow.total_length, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("J" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.weld_size) Then
                                .Worksheets("MP Connection Summary").Range("K" & BridgeDRow + 164).Value = CType(bsdrow.weld_size, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("K" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.exx) Then
                                .Worksheets("MP Connection Summary").Range("L" & BridgeDRow + 164).Value = CType(bsdrow.exx, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("L" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.upper_weld_length) Then
                                .Worksheets("MP Connection Summary").Range("M" & BridgeDRow + 164).Value = CType(bsdrow.upper_weld_length, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("M" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.lower_weld_length) Then
                                .Worksheets("MP Connection Summary").Range("N" & BridgeDRow + 164).Value = CType(bsdrow.lower_weld_length, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("N" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.upper_plate_width) Then
                                .Worksheets("MP Connection Summary").Range("O" & BridgeDRow + 164).Value = CType(bsdrow.upper_plate_width, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("O" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.lower_plate_width) Then
                                .Worksheets("MP Connection Summary").Range("P" & BridgeDRow + 164).Value = CType(bsdrow.lower_plate_width, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("P" & BridgeDRow + 164).ClearContents
                            End If
                            If Not IsNothing(bsdrow.neglect_flange_connection) Then
                                If bsdrow.neglect_flange_connection = True Then
                                    .Worksheets("MP Connection Summary").Range("R" & BridgeDRow + 164).Value = "Yes"
                                Else
                                    .Worksheets("MP Connection Summary").Range("R" & BridgeDRow + 164).Value = "No"
                                End If
                            End If
                            If Not IsNothing(bsdrow.bolt_hole_diameter) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("AQ" & BridgeDRow + 25).Value = CType(bsdrow.bolt_hole_diameter, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("AQ" & BridgeDRow + 25).ClearContents
                            End If
                            If Not IsNothing(bsdrow.bolt_qty_eccentric) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("DE" & BridgeDRow + 25).Value = CType(bsdrow.bolt_qty_eccentric, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("DE" & BridgeDRow + 25).ClearContents
                            End If
                            If Not IsNothing(bsdrow.bolt_qty_shear) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("DF" & BridgeDRow + 25).Value = CType(bsdrow.bolt_qty_shear, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("DF" & BridgeDRow + 25).ClearContents
                            End If
                            If Not IsNothing(bsdrow.intermediate_bolt_spacing) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("DG" & BridgeDRow + 25).Value = CType(bsdrow.intermediate_bolt_spacing, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("DG" & BridgeDRow + 25).ClearContents
                            End If
                            If Not IsNothing(bsdrow.bolt_diameter) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("DH" & BridgeDRow + 25).Value = CType(bsdrow.bolt_diameter, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("DH" & BridgeDRow + 25).ClearContents
                            End If
                            If Not IsNothing(bsdrow.bolt_sleeve_diameter) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("DJ" & BridgeDRow + 25).Value = CType(bsdrow.bolt_sleeve_diameter, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("DJ" & BridgeDRow + 25).ClearContents
                            End If
                            If Not IsNothing(bsdrow.washer_diameter) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("DL" & BridgeDRow + 25).Value = CType(bsdrow.washer_diameter, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("DL" & BridgeDRow + 25).ClearContents
                            End If
                            If Not IsNothing(bsdrow.bolt_tensile_strength) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("DN" & BridgeDRow + 25).Value = CType(bsdrow.bolt_tensile_strength, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("DN" & BridgeDRow + 25).ClearContents
                            End If
                            If Not IsNothing(bsdrow.bolt_allowable_shear) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("DP" & BridgeDRow + 25).Value = CType(bsdrow.bolt_allowable_shear, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("DP" & BridgeDRow + 25).ClearContents
                            End If
                            If Not IsNothing(bsdrow.exx_shim_plate) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("ES" & BridgeDRow + 25).Value = CType(bsdrow.exx_shim_plate, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("ES" & BridgeDRow + 25).ClearContents
                            End If
                            If Not IsNothing(bsdrow.filler_shim_thickness) Then
                                .Worksheets("Bridge Stiffener Calcs").Range("ET" & BridgeDRow + 25).Value = CType(bsdrow.filler_shim_thickness, Double)
                            Else
                                .Worksheets("Bridge Stiffener Calcs").Range("ET" & BridgeDRow + 25).ClearContents
                            End If

                            BridgeDRow += 1
                            'BridgeDRow2 += 1

                        End If
                    Next

                    PlateRow += 1
                    PlateRow2 -= 2
                    'PlateRow3 -= 1
                    mycol += 1
                    bump += 5
                    bump2 += 100

                Next

                Dim polmatflag As Boolean = False
                Dim poltempMaterials As New List(Of CCIplateMaterial)
                Dim poltempMaterial As New CCIplateMaterial 'Temp material object to determine if already added to Excel
                'Pole Geometry (when CCIpole exists)
                'This is to ensure that the unreinforced geometry is always referenced in CCIplate. 
                'Sometimes the reinforced geometry is required depending on the type of connection and therefore a warning will be logged when CCIpole exists)
                If Me.ParentStructure.Poles().Count > 0 Then
                    If Me.ParentStructure.Poles(0).unreinf_sections.Count > 0 Then
                        Dim col, GeoRow As Integer
                        GeoRow = 18
                        .Worksheets("Sub Tables (SAPI)").Range("A4").Value = CType(True, Boolean) 'Flags if geometry was produced by CCIpole. If true, geometry won't pull in from tnx file path.
                        For Each ps As PoleSection In Me.ParentStructure.Poles(0).unreinf_sections
                            col = 3
                            If Not IsNothing(ps.length_section) Then .Worksheets("Main").Cells(GeoRow, col).Value = CType(ps.length_section, Double)
                            col += 1
                            If Not IsNothing(ps.length_splice) Then .Worksheets("Main").Cells(GeoRow, col).Value = CType(ps.length_splice, Double)
                            col += 1
                            If Not IsNothing(ps.num_sides) Then .Worksheets("Main").Cells(GeoRow, col).Value = CType(ps.num_sides, Integer)
                            col += 1
                            If Not IsNothing(ps.diam_top) Then .Worksheets("Main").Cells(GeoRow, col).Value = CType(ps.diam_top, Double)
                            col += 1
                            If Not IsNothing(ps.diam_bot) Then .Worksheets("Main").Cells(GeoRow, col).Value = CType(ps.diam_bot, Double)
                            col += 1
                            If Not IsNothing(ps.wall_thickness) Then .Worksheets("Main").Cells(GeoRow, col).Value = CType(ps.wall_thickness, Double)
                            col += 1
                            If Not IsNothing(ps.matl_id) Then
                                For Each matl As PoleMatlProp In Me.ParentStructure.Poles(0).matls
                                    If matl.matl_id = ps.matl_id Then
                                        .Worksheets("Main").Cells(GeoRow, col).Value = CType(matl.name, String)
                                        'Determine if material needs to be added to CCIplate's material database
                                        'check and see if material matches default materials in CCIplate.
                                        For Each mrow As CCIplateMaterial In CCIplateMaterials
                                            If mrow.name = matl.name And mrow.fy_0 = matl.fy And mrow.fu_0 = matl.fu Then
                                                polmatflag = True 'don't add to materials database, already exists
                                                Exit For
                                            End If
                                        Next
                                        If polmatflag = False Then
                                            'Check and see if material matches temp materials (nondefault associated to site)
                                            For Each tmrow In tempMaterials
                                                If tmrow.name = matl.name And tmrow.fy_0 = matl.fy And tmrow.fu_0 = matl.fu Then
                                                    polmatflag = True 'don't add to materials database, already exists
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                        If polmatflag = False Then
                                            'check and see if ccipole material already added
                                            For Each ptmrow In poltempMaterials
                                                If ptmrow.ID = ps.matl_id Then
                                                    polmatflag = True 'don't add to excel
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                        'If false, add to materials database. 
                                        If polmatflag = False Then
                                            poltempMaterial = New CCIplateMaterial(matl.matl_id)
                                            poltempMaterials.Add(poltempMaterial)
                                            If Not IsNothing(matl.name) Then
                                                .Worksheets("Materials").Range("B" & MatRow).Value = CType(matl.name, String)
                                            End If
                                            If Not IsNothing(matl.fy) Then
                                                .Worksheets("Materials").Range("C" & MatRow).Value = CType(matl.fy, Double)
                                            Else
                                                .Worksheets("Materials").Range("C" & MatRow).ClearContents
                                            End If
                                            If Not IsNothing(matl.fu) Then
                                                .Worksheets("Materials").Range("K" & MatRow).Value = CType(matl.fu, Double)
                                            Else
                                                .Worksheets("Materials").Range("K" & MatRow).ClearContents
                                            End If
                                            MatRow += 1
                                        End If

                                        Exit For
                                    End If

                                Next
                            End If

                            polmatflag = False 'reset flag
                            GeoRow += 1
                        Next
                    End If
                End If

                'Update named ranges so dropdown seletions work correctly
                'Materials
                Dim qty, polqty As Integer
                qty = tempMaterials.Count
                polqty = poltempMaterials.Count
                'Dim definedName As DefinedName = .DefinedNames.Add("Materials", "Materials!$B$5:$D$" & qty + 39)
                'Dim definedName As DefinedName = .DefinedNames.
                ''Dim definedName As DefinedName = .DefinedNames.scope()
                'Dim rangeC2D3 As CellRange = .Range(definedName.Name)
                ''IWorkbook.DefinedNames

                Dim definedName As DefinedName = .DefinedNames.GetDefinedName("Materials")
                definedName.RefersTo = "Materials!$B$5:$B$" & qty + polqty + 39


            End If







            ''Hiding/unhiding specific tabs
            'If Me.pile_group_config = "Circular" Then
            '    .Worksheets("Moment of Inertia").Visible = False
            '    .Worksheets("Moment of Inertia (Circle)").Visible = True
            'Else
            '    .Worksheets("Moment of Inertia").Visible = True
            '    .Worksheets("Moment of Inertia (Circle)").Visible = False
            'End If

            ''Resizing Image 'User is currently running solution which will resize image within tool
            ''Try
            ''    With .Worksheets("Input").Charts(0)
            ''        .Width = (300 / Math.Max(CType(pf.pad_width_dir1, Double), CType(pf.pad_width_dir2, Double))) * CType(pf.pad_width_dir1, Double) * 4.19 '4.19 multiplier determined through trial and error. 
            ''        .Height = (300 / Math.Max(CType(pf.pad_width_dir1, Double), CType(pf.pad_width_dir2, Double))) * CType(pf.pad_width_dir2, Double) * 4.19
            ''    End With
            ''Catch
            ''    'error handling to avoid dividing by zero
            ''End Try


        End With

    End Sub

    Private Sub SaveMaterial(ByRef wb As Workbook, ByVal mrow As CCIplateMaterial, MatRow As Integer)

        With wb

            If Not IsNothing(mrow.ID) Then
                .Worksheets("Sub Tables (SAPI)").Range("AT" & MatRow - 2).Value = CType(mrow.ID, Integer)
            End If
            If Not IsNothing(mrow.name) Then
                .Worksheets("Materials").Range("B" & MatRow).Value = CType(mrow.name, String)
            End If

            If Not IsNothing(mrow.fy_0) Then
                .Worksheets("Materials").Range("C" & MatRow).Value = CType(mrow.fy_0, Double)
            Else
                .Worksheets("Materials").Range("C" & MatRow).ClearContents
            End If
            If Not IsNothing(mrow.fy_1_125) Then
                .Worksheets("Materials").Range("D" & MatRow).Value = CType(mrow.fy_1_125, Double)
            Else
                .Worksheets("Materials").Range("D" & MatRow).ClearContents
            End If
            If Not IsNothing(mrow.fy_1_625) Then
                .Worksheets("Materials").Range("E" & MatRow).Value = CType(mrow.fy_1_625, Double)
            Else
                .Worksheets("Materials").Range("E" & MatRow).ClearContents
            End If
            If Not IsNothing(mrow.fy_2_625) Then
                .Worksheets("Materials").Range("F" & MatRow).Value = CType(mrow.fy_2_625, Double)
            Else
                .Worksheets("Materials").Range("F" & MatRow).ClearContents
            End If
            If Not IsNothing(mrow.fy_4_125) Then
                .Worksheets("Materials").Range("G" & MatRow).Value = CType(mrow.fy_4_125, Double)
            Else
                .Worksheets("Materials").Range("G" & MatRow).ClearContents
            End If
            If Not IsNothing(mrow.fu_0) Then
                .Worksheets("Materials").Range("K" & MatRow).Value = CType(mrow.fu_0, Double)
            Else
                .Worksheets("Materials").Range("K" & MatRow).ClearContents
            End If
            If Not IsNothing(mrow.fu_1_125) Then
                .Worksheets("Materials").Range("L" & MatRow).Value = CType(mrow.fu_1_125, Double)
            Else
                .Worksheets("Materials").Range("L" & MatRow).ClearContents
            End If
            If Not IsNothing(mrow.fu_1_625) Then
                .Worksheets("Materials").Range("M" & MatRow).Value = CType(mrow.fu_1_625, Double)
            Else
                .Worksheets("Materials").Range("M" & MatRow).ClearContents
            End If
            If Not IsNothing(mrow.fu_2_625) Then
                .Worksheets("Materials").Range("N" & MatRow).Value = CType(mrow.fu_2_625, Double)
            Else
                .Worksheets("Materials").Range("N" & MatRow).ClearContents
            End If
            If Not IsNothing(mrow.fu_4_125) Then
                .Worksheets("Materials").Range("O" & MatRow).Value = CType(mrow.fu_4_125, Double)
            Else
                .Worksheets("Materials").Range("O" & MatRow).ClearContents
            End If

        End With

    End Sub

#End Region

#Region "Save to EDS"

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.anchor_rod_spacing.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.clip_distance.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.barb_cl_elevation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.include_pole_reactions.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.consider_ar_eccentricity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.leg_mod_eccentricity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.seismic.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.seismic_flanges.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Structural_105.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Version.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("anchor_rod_spacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("clip_distance")
        SQLInsertFields = SQLInsertFields.AddtoDBString("barb_cl_elevation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("include_pole_reactions")
        SQLInsertFields = SQLInsertFields.AddtoDBString("consider_ar_eccentricity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("leg_mod_eccentricity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("seismic")
        SQLInsertFields = SQLInsertFields.AddtoDBString("seismic_flanges")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Structural_105")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tool_version")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        'SQLUpdate = SQLUpdate.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit = " & Me.bus_unit.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id = " & Me.structure_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("anchor_rod_spacing = " & Me.anchor_rod_spacing.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("clip_distance = " & Me.clip_distance.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("barb_cl_elevation = " & Me.barb_cl_elevation.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("include_pole_reactions = " & Me.include_pole_reactions.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("consider_ar_eccentricity = " & Me.consider_ar_eccentricity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("leg_mod_eccentricity = " & Me.leg_mod_eccentricity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("seismic = " & Me.seismic.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("seismic_flanges = " & Me.seismic_flanges.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("Structural_105 = " & Me.Structural_105.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tool_version = " & Me.Version.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function

#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As CCIplate = TryCast(other, CCIplate)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        'Equals = If(Me.bus_unit.CheckChange(otherToCompare.bus_unit, changes, categoryName, "Bus Unit"), Equals, False)
        'Equals = If(Me.structure_id.CheckChange(otherToCompare.structure_id, changes, categoryName, "Structure Id"), Equals, False)
        Equals = If(Me.anchor_rod_spacing.CheckChange(otherToCompare.anchor_rod_spacing, changes, categoryName, "Anchor Rod Spacing"), Equals, False)
        Equals = If(Me.clip_distance.CheckChange(otherToCompare.clip_distance, changes, categoryName, "Clip Distance"), Equals, False)
        Equals = If(Me.barb_cl_elevation.CheckChange(otherToCompare.barb_cl_elevation, changes, categoryName, "Barb Cl Elevation"), Equals, False)
        Equals = If(Me.include_pole_reactions.CheckChange(otherToCompare.include_pole_reactions, changes, categoryName, "Include Pole Reactions"), Equals, False)
        Equals = If(Me.consider_ar_eccentricity.CheckChange(otherToCompare.consider_ar_eccentricity, changes, categoryName, "Consider Ar Eccentricity"), Equals, False)
        Equals = If(Me.leg_mod_eccentricity.CheckChange(otherToCompare.leg_mod_eccentricity, changes, categoryName, "Leg Mod Eccentricity"), Equals, False)
        Equals = If(Me.seismic.CheckChange(otherToCompare.seismic, changes, categoryName, "Seismic"), Equals, False)
        Equals = If(Me.seismic_flanges.CheckChange(otherToCompare.seismic_flanges, changes, categoryName, "Seismic Flanges"), Equals, False)
        Equals = If(Me.Structural_105.CheckChange(otherToCompare.Structural_105, changes, categoryName, "Structural 105"), Equals, False)
        Equals = If(Me.Version.CheckChange(otherToCompare.Version, changes, categoryName, "Tool Version"), Equals, False)
        'Equals = If(Me.modified_person_id.CheckChange(otherToCompare.modified_person_id, changes, categoryName, "Modified Person Id"), Equals, False)
        'Equals = If(Me.process_stage.CheckChange(otherToCompare.process_stage, changes, categoryName, "Process Stage"), Equals, False)

        'Connection
        If Me.Connections.Count > 0 Then
            Equals = If(Me.Connections.CheckChange(otherToCompare.Connections, changes, categoryName, "Plates"), Equals, False)
        End If

        Return Equals

    End Function
#End Region

    Public Overrides Sub Clear()
        Me.Connections.Clear()
        Me.Results.Clear()
    End Sub
End Class

<DataContractAttribute()>
Partial Public Class Connection
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Connections"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.plates"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Connection (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Connection_INSERT
        SQLInsert = SQLInsert.Replace("[CONNECTION VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[CONNECTION FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        'Plate Detail
        If Me.PlateDetails.Count > 0 Then
            SQLInsert = SQLInsert.Replace("--BEGIN --[PLATE DETAIL INSERT BEGIN]", "BEGIN --[PLATE DETAIL INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[PLATE DETAIL INSERT END]", "END --[PLATE DETAIL INSERT END]")
            For Each row As PlateDetail In PlateDetails
                SQLInsert = SQLInsert.Replace("--[PLATE DETAIL INSERT]", row.SQLInsert)
            Next
        End If

        'Bolt Group
        If Me.BoltGroups.Count > 0 Then
            SQLInsert = SQLInsert.Replace("--BEGIN --[BOLT GROUP INSERT BEGIN]", "BEGIN --[BOLT GROUP INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[BOLT GROUP INSERT END]", "END --[BOLT GROUP INSERT END]")
            For Each row As BoltGroup In BoltGroups
                SQLInsert = SQLInsert.Replace("--[BOLT GROUP INSERT]", row.SQLInsert)
            Next
        End If

        'Bridge Stiffener Details
        If Me.BridgeStiffenerDetails.Count > 0 Then
            SQLInsert = SQLInsert.Replace("--BEGIN --[BRIDGE STIFFENER DETAIL INSERT BEGIN]", "BEGIN --[BRIDGE STIFFENER DETAIL INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[BRIDGE STIFFENER DETAIL INSERT END]", "END --[BRIDGE STIFFENER DETAIL INSERT END]")
            For Each row As BridgeStiffenerDetail In BridgeStiffenerDetails
                SQLInsert = SQLInsert.Replace("--[BRIDGE STIFFENER DETAIL INSERT]", row.SQLInsert)
            Next
        End If

        'Connection Results
        For Each row As ConnectionResults In ConnectionResults
            SQLInsert = SQLInsert.Replace("--BEGIN --[CONNECTION RESULTS INSERT BEGIN]", "BEGIN --[CONNECTION RESULTS INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[CONNECTION RESULTS INSERT END]", "END --[CONNECTION RESULTS INSERT END]")
            SQLInsert = SQLInsert.Replace("--[CONNECTION RESULTS INSERT]", row.SQLInsert)
        Next

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Connection (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Connection_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        'Plate Detail
        If Me.PlateDetails.Count > 0 Then
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[PLATE DETAIL UPDATE BEGIN]", "BEGIN --[PLATE DETAIL UPDATE BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[PLATE DETAIL UPDATE END]", "END --[PLATE DETAIL UPDATE END]")
            For Each row As PlateDetail In PlateDetails
                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                    If IsSomethingString(row.plate_location) Or IsSomethingString(row.plate_type) Or IsSomething(row.plate_diameter) Or IsSomething(row.plate_thickness) Or IsSomething(row.plate_material) Or IsSomething(row.stiffener_configuration) Or IsSomething(row.stiffener_clear_space) Or IsSomething(row.plate_check) Then
                        SQLUpdate = SQLUpdate.Replace("--[PLATE DETAIL INSERT]", row.SQLUpdate)
                    Else
                        SQLUpdate = SQLUpdate.Replace("--[PLATE DETAIL INSERT]", row.SQLDelete)
                    End If
                Else
                    SQLUpdate = SQLUpdate.Replace("--[PLATE DETAIL INSERT]", row.SQLInsert)
                End If
            Next
        End If

        'Bolt Group
        If Me.BoltGroups.Count > 0 Then
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[BOLT GROUP UPDATE BEGIN]", "BEGIN --[BOLT GROUP UPDATE BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[BOLT GROUP UPDATE END]", "END --[BOLT GROUP UPDATE END]")
            For Each row As BoltGroup In BoltGroups
                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                    If IsSomething(row.resist_axial) Or IsSomething(row.resist_shear) Or IsSomething(row.plate_bending) Or IsSomething(row.grout_considered) Or IsSomething(row.apply_barb_elevation) Or IsSomething(row.bolt_name) Then
                        SQLUpdate = SQLUpdate.Replace("--[BOLT GROUP INSERT]", row.SQLUpdate)
                    Else
                        SQLUpdate = SQLUpdate.Replace("--[BOLT GROUP INSERT]", row.SQLDelete)
                    End If
                Else
                    SQLUpdate = SQLUpdate.Replace("--[BOLT GROUP INSERT]", row.SQLInsert)
                End If
            Next
        End If

        'Bridge Stiffener Details
        If Me.BridgeStiffenerDetails.Count > 0 Then
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[BRIDGE STIFFENER DETAIL UPDATE BEGIN]", "BEGIN --[BRIDGE STIFFENER DETAIL UPDATE BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[BRIDGE STIFFENER DETAIL UPDATE END]", "END --[BRIDGE STIFFENER DETAIL UPDATE END]")
            For Each row As BridgeStiffenerDetail In BridgeStiffenerDetails
                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                    If IsSomething(row.local_connection_id) Or IsSomethingString(row.stiffener_type) Or IsSomethingString(row.analysis_type) Or IsSomething(row.quantity) Or IsSomething(row.bridge_stiffener_width) _
                        Or IsSomething(row.bridge_stiffener_thickness) Or IsSomething(row.bridge_stiffener_material) Or IsSomething(row.unbraced_length) Or IsSomething(row.total_length) _
                        Or IsSomething(row.weld_size) Or IsSomething(row.exx) Or IsSomething(row.upper_weld_length) Or IsSomething(row.lower_weld_length) _
                        Or IsSomething(row.upper_plate_width) Or IsSomething(row.lower_plate_width) Or IsSomething(row.neglect_flange_connection) Then
                        'not including below since a user typically won't delete associated fields since on a seperate window in tool. 
                        'Or IsSomething(row.bolt_hole_diameter) _
                        'Or IsSomething(row.bolt_qty_eccentric) Or IsSomething(row.bolt_qty_shear) Or IsSomething(row.intermediate_bolt_spacing) Or IsSomething(row.bolt_diameter) _
                        'Or IsSomething(row.bolt_sleeve_diameter) Or IsSomething(row.washer_diameter) Or IsSomething(row.bolt_tensile_strength) Or IsSomething(row.bolt_allowable_shear) _
                        'Or IsSomething(row.exx_shim_plate) Or IsSomething(row.filler_shim_thickness)
                        SQLUpdate = SQLUpdate.Replace("--[BRIDGE STIFFENER DETAIL INSERT]", row.SQLUpdate)
                    Else
                        SQLUpdate = SQLUpdate.Replace("--[BRIDGE STIFFENER DETAIL INSERT]", row.SQLDelete)
                    End If
                Else
                    SQLUpdate = SQLUpdate.Replace("--[BRIDGE STIFFENER DETAIL INSERT]", row.SQLInsert)
                End If
            Next
        End If

        'Connection Results
        For Each row As ConnectionResults In ConnectionResults
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[CONNECTION RESULTS INSERT BEGIN]", "BEGIN --[CONNECTION RESULTS INSERT BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[CONNECTION RESULTS INSERT END]", "END --[CONNECTION RESULTS INSERT END]")
            SQLUpdate = SQLUpdate.Replace("--[CONNECTION RESULTS INSERT]", row.SQLInsert)
        Next

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Connection (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Connection_DELETE
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        'Plate Details
        If Me.PlateDetails.Count > 0 Then
            SQLDelete = SQLDelete.Replace("--BEGIN --[PLATE DETAIL DELETE BEGIN]", "BEGIN --[PLATE DETAIL DELETE BEGIN]")
            SQLDelete = SQLDelete.Replace("--END --[PLATE DETAIL DELETE END]", "END --[PLATE DETAIL DELETE END]")
            For Each row As PlateDetail In PlateDetails
                SQLDelete = SQLDelete.Replace("--[PLATE DETAIL INSERT]", row.SQLDelete)
            Next
        End If

        'Bolt Groups
        If Me.BoltGroups.Count > 0 Then
            SQLDelete = SQLDelete.Replace("--BEGIN --[BOLT GROUP DELETE BEGIN]", "BEGIN --[BOLT GROUP DELETE BEGIN]")
            SQLDelete = SQLDelete.Replace("--END --[BOLT GROUP DELETE END]", "END --[BOLT GROUP DELETE END]")
            For Each row As BoltGroup In BoltGroups
                SQLDelete = SQLDelete.Replace("--[BOLT GROUP INSERT]", row.SQLDelete)
            Next
        End If

        'Bridge Stiffener Details
        If Me.BridgeStiffenerDetails.Count > 0 Then
            SQLDelete = SQLDelete.Replace("--BEGIN --[BRIDGE STIFFENER DETAIL DELETE BEGIN]", "BEGIN --[BRIDGE STIFFENER DETAIL DELETE BEGIN]")
            SQLDelete = SQLDelete.Replace("--END --[BRIDGE STIFFENER DETAIL DELETE END]", "END --[BRIDGE STIFFENER DETAIL DELETE END]")
            For Each row As BridgeStiffenerDetail In BridgeStiffenerDetails
                SQLDelete = SQLDelete.Replace("--[BRIDGE STIFFENER DETAIL INSERT]", row.SQLDelete)
            Next
        End If

        Return SQLDelete

    End Function

#End Region

#Region "Define"
    Private _local_id As Integer?
    Private _connection_elevation As Double?
    Private _connection_type As String
    Private _bolt_configuration As String
    Private _cciplate_id As Integer?

    <DataMember()> Public Property PlateDetails As New List(Of PlateDetail)
    <DataMember()> Public Property BoltGroups As New List(Of BoltGroup)
    <DataMember()> Public Property BridgeStiffenerDetails As New List(Of BridgeStiffenerDetail)
    <DataMember()> Public Property ConnectionResults As New List(Of ConnectionResults)

    <Category("Connection"), Description(""), DisplayName("Local Id")>
    <DataMember()> Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property
    <Category("Connection"), Description(""), DisplayName("Connection Elevation")>
    <DataMember()> Public Property connection_elevation() As Double?
        Get
            Return Me._connection_elevation
        End Get
        Set
            Me._connection_elevation = Value
        End Set
    End Property
    <Category("Connection"), Description(""), DisplayName("Connection Type")>
    <DataMember()> Public Property connection_type() As String
        Get
            Return Me._connection_type
        End Get
        Set
            Me._connection_type = Value
        End Set
    End Property
    <Category("Connection"), Description(""), DisplayName("Bolt Configuration")>
    <DataMember()> Public Property bolt_configuration() As String
        Get
            Return Me._bolt_configuration
        End Get
        Set
            Me._bolt_configuration = Value
        End Set
    End Property
    <Category("Connection"), Description(""), DisplayName("CCIplate Id")>
    <DataMember()> Public Property cciplate_id() As Integer?
        Get
            Return Me._cciplate_id
        End Get
        Set
            Me._cciplate_id = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal row As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing) '(ByVal prow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = row
        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_id"))) '-Must only associate this to EDS since 0 vs. >0 triggers different functions (e.g. update vs. delete)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_id = DBtoNullableInt(dr.Item("local_connection_id"))
        End If
        Me.connection_elevation = DBtoNullableDbl(dr.Item("connection_elevation"))
        Me.connection_type = DBtoStr(dr.Item("connection_type"))
        Me.bolt_configuration = DBtoStr(dr.Item("bolt_configuration"))
        'Me.cciplate_id = If(EDStruefalse, DBtoNullableInt(dr.Item("connection_id")), Me.cciplate_id) 'Not provided in Excel
        If EDStruefalse = True Then 'Only Pull in when referencing EDS
            Me.cciplate_id = DBtoNullableInt(dr.Item("connection_id"))
        End If

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@TopLevelID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.connection_elevation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.connection_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_configuration.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_elevation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("connection_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_configuration")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_id = " & "@TopLevelID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_elevation = " & Me.connection_elevation.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_type = " & Me.connection_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_configuration = " & Me.bolt_configuration.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As Connection = TryCast(other, Connection)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.connection_elevation.CheckChange(otherToCompare.connection_elevation, changes, categoryName, "Connection Elevation"), Equals, False)
        Equals = If(Me.connection_type.CheckChange(otherToCompare.connection_type, changes, categoryName, "Connection Type"), Equals, False)
        Equals = If(Me.bolt_configuration.CheckChange(otherToCompare.bolt_configuration, changes, categoryName, "Bolt Configuration"), Equals, False)

        'Plate Details
        If Me.PlateDetails.Count > 0 Then
            Equals = If(Me.PlateDetails.CheckChange(otherToCompare.PlateDetails, changes, categoryName, "Plate Details"), Equals, False)
        End If

        'Bolt Groups
        If Me.BoltGroups.Count > 0 Then
            Equals = If(Me.BoltGroups.CheckChange(otherToCompare.BoltGroups, changes, categoryName, "Bolt Groups"), Equals, False)
        End If

        'Bridge Stiffeners
        If Me.BridgeStiffenerDetails.Count > 0 Then
            Equals = If(Me.BridgeStiffenerDetails.CheckChange(otherToCompare.BridgeStiffenerDetails, changes, categoryName, "Bridge Stiffener Details"), Equals, False)
        End If

    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class PlateDetail
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Plate Details"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.plate_details"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Plate Detail (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Plate_Detail_INSERT
        SQLInsert = SQLInsert.Replace("[PLATE DETAIL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[PLATE DETAIL FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        'Plate Material
        For Each row As CCIplateMaterial In CCIplateMaterials
            SQLInsert = SQLInsert.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert)
        Next

        'Results
        'If Me.Results.Count > 0 Then
        For Each row As PlateResults In PlateResults
            SQLInsert = SQLInsert.Replace("--BEGIN --[PLATE RESULTS INSERT BEGIN]", "BEGIN --[PLATE RESULTS INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[PLATE RESULTS INSERT END]", "END --[PLATE RESULTS INSERT END]")
            SQLInsert = SQLInsert.Replace("--[PLATE RESULTS INSERT]", row.SQLInsert)
        Next
        'End If

        'Stiffener Group
        If Me.StiffenerGroups.Count > 0 Then
            SQLInsert = SQLInsert.Replace("--BEGIN --[STIFFENER GROUP INSERT BEGIN]", "BEGIN --[STIFFENER GROUP INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[STIFFENER GROUP INSERT END]", "END --[STIFFENER GROUP INSERT END]")
            For Each row As StiffenerGroup In StiffenerGroups
                SQLInsert = SQLInsert.Replace("--[STIFFENER GROUP INSERT]", row.SQLInsert)
            Next
        End If

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Plate Detail (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Plate_Detail_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        'Plate Material
        For Each row As CCIplateMaterial In CCIplateMaterials
            SQLUpdate = SQLUpdate.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert) 'Can only insert materials, no deleting or updating since database is referenced by all BUs. 
        Next

        'Results
        'If Me.Results.Count > 0 Then
        '    SQLUpdate = SQLUpdate.Replace("--BEGIN --[RESULTS UPDATE BEGIN]", "BEGIN --[RESULTS UPDATE BEGIN]")
        '    SQLUpdate = SQLUpdate.Replace("--END --[RESULTS UPDATE END]", "END --[RESULTS UPDATE END]")
        '    SQLUpdate = SQLUpdate.Replace("--[RESULTS INSERT]", Me.Results.EDSResultQuery)
        'End If
        'Insert is always performed for results
        For Each row As PlateResults In PlateResults
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[PLATE RESULTS INSERT BEGIN]", "BEGIN --[PLATE RESULTS INSERT BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[PLATE RESULTS INSERT END]", "END --[PLATE RESULTS INSERT END]")
            SQLUpdate = SQLUpdate.Replace("--[PLATE RESULTS INSERT]", row.SQLInsert)
        Next

        'Stiffener Groups
        If Me.StiffenerGroups.Count > 0 Then
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[STIFFENER GROUP UPDATE BEGIN]", "BEGIN --[STIFFENER GROUP UPDATE BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[STIFFENER GROUP UPDATE END]", "END --[STIFFENER GROUP UPDATE END]")
            For Each row As StiffenerGroup In StiffenerGroups
                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                    If IsSomethingString(row.stiffener_name) Then
                        SQLUpdate = SQLUpdate.Replace("--[STIFFENER GROUP INSERT]", row.SQLUpdate)
                    Else
                        SQLUpdate = SQLUpdate.Replace("--[STIFFENER GROUP INSERT]", row.SQLDelete)
                    End If
                Else
                    SQLUpdate = SQLUpdate.Replace("--[STIFFENER GROUP INSERT]", row.SQLInsert)
                End If
            Next
        End If

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Plate Detail (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Plate_Detail_DELETE
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        'Stiffener Groups
        If Me.StiffenerGroups.Count > 0 Then
            SQLDelete = SQLDelete.Replace("--BEGIN --[STIFFENER GROUP DELETE BEGIN]", "BEGIN --[STIFFENER GROUP DELETE BEGIN]")
            SQLDelete = SQLDelete.Replace("--END --[STIFFENER GROUP DELETE END]", "END --[STIFFENER GROUP DELETE END]")
            For Each row As StiffenerGroup In StiffenerGroups
                SQLDelete = SQLDelete.Replace("--[STIFFENER GROUP INSERT]", row.SQLDelete)
            Next
        End If

        Return SQLDelete

    End Function

#End Region

#Region "Define"
    Private _local_id As Integer?
    Private _local_connection_id As Integer?
    Private _connection_id As Integer? 'currently called plate_id in EDS
    Private _plate_location As String
    Private _plate_type As String
    Private _plate_diameter As Double?
    Private _plate_thickness As Double?
    Private _plate_material As Integer?
    Private _stiffener_configuration As Integer?
    Private _stiffener_clear_space As Double?
    Private _plate_check As Boolean?

    <DataMember()> Public Property CCIplateMaterials As New List(Of CCIplateMaterial)
    <DataMember()> Public Property PlateResults As New List(Of PlateResults)
    <DataMember()> Public Property StiffenerGroups As New List(Of StiffenerGroup)

    <Category("Plate Details"), Description(""), DisplayName("Local Id")>
    <DataMember()> Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Local Connection Id")>
    <DataMember()> Public Property local_connection_id() As Integer?
        Get
            Return Me._local_connection_id
        End Get
        Set
            Me._local_connection_id = Value
        End Set
    End Property

    <Category("Plate Details"), Description(""), DisplayName("Connection Id")>
    <DataMember()> Public Property connection_id() As Integer?
        Get
            Return Me._connection_id
        End Get
        Set
            Me._connection_id = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Location")>
    <DataMember()> Public Property plate_location() As String
        Get
            Return Me._plate_location
        End Get
        Set
            Me._plate_location = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Type")>
    <DataMember()> Public Property plate_type() As String
        Get
            Return Me._plate_type
        End Get
        Set
            Me._plate_type = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Diameter")>
    <DataMember()> Public Property plate_diameter() As Double?
        Get
            Return Me._plate_diameter
        End Get
        Set
            Me._plate_diameter = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Thickness")>
    <DataMember()> Public Property plate_thickness() As Double?
        Get
            Return Me._plate_thickness
        End Get
        Set
            Me._plate_thickness = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Material")>
    <DataMember()> Public Property plate_material() As Integer?
        Get
            Return Me._plate_material
        End Get
        Set
            Me._plate_material = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Stiffener Configuration")>
    <DataMember()> Public Property stiffener_configuration() As Integer?
        Get
            Return Me._stiffener_configuration
        End Get
        Set
            Me._stiffener_configuration = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Stiffener Clear Space")>
    <DataMember()> Public Property stiffener_clear_space() As Double?
        Get
            Return Me._stiffener_clear_space
        End Get
        Set
            Me._stiffener_clear_space = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Check")>
    <DataMember()> Public Property plate_check() As Boolean?
        Get
            Return Me._plate_check
        End Get
        Set
            Me._plate_check = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal pdrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = pdrow
        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_plate_id")))
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_id = DBtoNullableInt(dr.Item("local_plate_id"))
            Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id"))
        End If
        'Me.connection_id = If(EDStruefalse, DBtoNullableInt(dr.Item("plate_id")), DBtoNullableInt(dr.Item("local_connection_id"))) 'ME.plate_id '-pulls in null when Excel is referenced. 
        If EDStruefalse = True Then 'only pull in when referencing EDS
            Me.connection_id = DBtoNullableInt(dr.Item("plate_id"))
        End If
        Me.plate_location = DBtoStr(dr.Item("plate_location"))
        Me.plate_type = DBtoStr(dr.Item("plate_type"))
        Me.plate_diameter = DBtoNullableDbl(dr.Item("plate_diameter"))
        Me.plate_thickness = DBtoNullableDbl(dr.Item("plate_thickness"))
        Me.plate_material = DBtoNullableInt(dr.Item("plate_material"))
        Me.stiffener_configuration = If(DBtoStr(dr.Item("stiffener_configuration")) = "Custom", 4, DBtoNullableInt(dr.Item("stiffener_configuration"))) 'Stiffener configuration is 0 through 3 plus 'custom'. Custom will report as option 4. 
        Me.stiffener_clear_space = DBtoNullableDbl(dr.Item("stiffener_clear_space"))
        Me.plate_check = If(EDStruefalse, DBtoNullableBool(dr.Item("plate_check")), If(DBtoStr(dr.Item("plate_check")) = "Yes", True, If(DBtoStr(dr.Item("plate_check")) = "No", False, DBtoNullableBool(dr.Item("plate_check"))))) 'Listed as a string and need to convert to Boolean

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_location.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_thickness.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_material.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_configuration.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_clear_space.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_check.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_location")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_material")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_configuration")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_clear_space")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_check")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_id = " & "@SubLevel1ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_location = " & Me.plate_location.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_type = " & Me.plate_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_diameter = " & Me.plate_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_thickness = " & Me.plate_thickness.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_material = " & Me.plate_material.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_material = " & "@SubLevel3ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_configuration = " & Me.stiffener_configuration.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_clear_space = " & Me.stiffener_clear_space.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_check = " & Me.plate_check.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'plate_material references the local id when coming from Excel. Need to convert to EDS ID when performing Equals function
        Dim material As Integer?
        For Each row As CCIplateMaterial In CCIplateMaterials
            If Me.plate_material = row.local_id And row.ID > 0 Then
                material = row.ID
                Exit For
            End If
        Next


        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As PlateDetail = TryCast(other, PlateDetail)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        'Equals = If(Me.plate_id.CheckChange(otherToCompare.plate_id, changes, categoryName, "Plate Id"), Equals, False)
        Equals = If(Me.plate_location.CheckChange(otherToCompare.plate_location, changes, categoryName, "Plate Location"), Equals, False)
        Equals = If(Me.plate_type.CheckChange(otherToCompare.plate_type, changes, categoryName, "Plate Type"), Equals, False)
        Equals = If(Me.plate_diameter.CheckChange(otherToCompare.plate_diameter, changes, categoryName, "Plate Diameter"), Equals, False)
        Equals = If(Me.plate_thickness.CheckChange(otherToCompare.plate_thickness, changes, categoryName, "Plate Thickness"), Equals, False)
        'Equals = If(Me.plate_material.CheckChange(otherToCompare.plate_material, changes, categoryName, "Plate Material"), Equals, False)
        Equals = If(material.CheckChange(otherToCompare.plate_material, changes, categoryName, "Plate Material"), Equals, False)
        Equals = If(Me.stiffener_configuration.CheckChange(otherToCompare.stiffener_configuration, changes, categoryName, "Stiffener Configuration"), Equals, False)
        Equals = If(Me.stiffener_clear_space.CheckChange(otherToCompare.stiffener_clear_space, changes, categoryName, "Stiffener Clear Space"), Equals, False)
        Equals = If(Me.plate_check.CheckChange(otherToCompare.plate_check, changes, categoryName, "Plate Check"), Equals, False)

        'Materials
        If Me.CCIplateMaterials.Count > 0 Then
            Equals = If(Me.CCIplateMaterials.CheckChange(otherToCompare.CCIplateMaterials, changes, categoryName, "CCIplate Materials"), Equals, False)
        End If

        'Stiffener Groups
        If Me.StiffenerGroups.Count > 0 Then
            Equals = If(Me.StiffenerGroups.CheckChange(otherToCompare.StiffenerGroups, changes, categoryName, "Stiffener Groups"), Equals, False)
        End If

    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class BoltGroup
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Bolt Groups"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.bolts"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Group (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Group_INSERT
        SQLInsert = SQLInsert.Replace("[BOLT GROUP VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[BOLT GROUP FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        'Bolt Detail
        If Me.BoltDetails.Count > 0 Then
            SQLInsert = SQLInsert.Replace("--BEGIN --[BOLT DETAIL INSERT BEGIN]", "BEGIN --[BOLT DETAIL INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[BOLT DETAIL INSERT END]", "END --[BOLT DETAIL INSERT END]")
            For Each row As BoltDetail In BoltDetails
                SQLInsert = SQLInsert.Replace("--[BOLT DETAIL INSERT]", row.SQLInsert)
            Next
        End If

        'Results
        For Each row As BoltResults In BoltResults
            SQLInsert = SQLInsert.Replace("--BEGIN --[BOLT RESULTS INSERT BEGIN]", "BEGIN --[BOLT RESULTS INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[BOLT RESULTS INSERT END]", "END --[BOLT RESULTS INSERT END]")
            SQLInsert = SQLInsert.Replace("--[BOLT RESULTS INSERT]", row.SQLInsert)
        Next

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Group (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Group_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        'Bolt Detail
        If Me.BoltDetails.Count > 0 Then
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[BOLT DETAIL UPDATE BEGIN]", "BEGIN --[BOLT DETAIL UPDATE BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[BOLT DETAIL UPDATE END]", "END --[BOLT DETAIL UPDATE END]")
            For Each row As BoltDetail In BoltDetails
                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                    If IsSomething(row.bolt_location) Or IsSomething(row.bolt_diameter) Or IsSomething(row.bolt_material) Or IsSomething(row.bolt_circle) Or IsSomething(row.eta_factor) Or IsSomething(row.lar) Or IsSomethingString(row.bolt_thread_type) Or IsSomething(row.area_override) Or IsSomething(row.tension_only) Then
                        SQLUpdate = SQLUpdate.Replace("--[BOLT DETAIL INSERT]", row.SQLUpdate)
                    Else
                        SQLUpdate = SQLUpdate.Replace("--[BOLT DETAIL INSERT]", row.SQLDelete)
                    End If
                Else
                    SQLUpdate = SQLUpdate.Replace("--[BOLT DETAIL INSERT]", row.SQLInsert)
                End If
            Next
        End If

        For Each row As BoltResults In BoltResults
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[BOLT RESULTS INSERT BEGIN]", "BEGIN --[BOLT RESULTS INSERT BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[BOLT RESULTS INSERT END]", "END --[BOLT RESULTS INSERT END]")
            SQLUpdate = SQLUpdate.Replace("--[BOLT RESULTS INSERT]", row.SQLInsert)
        Next

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Group (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Group_DELETE
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        'Bolt Details
        If Me.BoltDetails.Count > 0 Then
            SQLDelete = SQLDelete.Replace("--BEGIN --[BOLT DETAIL DELETE BEGIN]", "BEGIN --[BOLT DETAIL DELETE BEGIN]")
            SQLDelete = SQLDelete.Replace("--END --[BOLT DETAIL DELETE END]", "END --[BOLT DETAIL DELETE END]")
            For Each row As BoltDetail In BoltDetails
                SQLDelete = SQLDelete.Replace("--[BOLT DETAIL INSERT]", row.SQLDelete)
            Next
        End If

        Return SQLDelete

    End Function

#End Region

#Region "Define"
    Private _local_id As Integer?
    Private _local_connection_id As Integer?
    Private _connection_id As Integer?
    Private _resist_axial As Boolean?
    Private _resist_shear As Boolean?
    Private _plate_bending As Boolean?
    Private _grout_considered As Boolean?
    Private _apply_barb_elevation As Boolean?
    Private _bolt_name As String


    <DataMember()> Public Property BoltDetails As New List(Of BoltDetail)
    <DataMember()> Public Property BoltResults As New List(Of BoltResults)

    <Category("Bolt Groups"), Description(""), DisplayName("Local Id")>
    <DataMember()> Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property
    <Category("Bolt Groups"), Description(""), DisplayName("Local Connection Id")>
    <DataMember()> Public Property local_connection_id() As Integer?
        Get
            Return Me._local_connection_id
        End Get
        Set
            Me._local_connection_id = Value
        End Set
    End Property

    <Category("Bolt Groups"), Description(""), DisplayName("Connection Id")>
    <DataMember()> Public Property connection_id() As Integer?
        Get
            Return Me._connection_id
        End Get
        Set
            Me._connection_id = Value
        End Set
    End Property
    <Category("Bolt Groups"), Description(""), DisplayName("Resist Axial")>
    <DataMember()> Public Property resist_axial() As Boolean?
        Get
            Return Me._resist_axial
        End Get
        Set
            Me._resist_axial = Value
        End Set
    End Property
    <Category("Bolt Groups"), Description(""), DisplayName("Resist Shear")>
    <DataMember()> Public Property resist_shear() As Boolean?
        Get
            Return Me._resist_shear
        End Get
        Set
            Me._resist_shear = Value
        End Set
    End Property
    <Category("Bolt Groups"), Description(""), DisplayName("Plate Bending")>
    <DataMember()> Public Property plate_bending() As Boolean?
        Get
            Return Me._plate_bending
        End Get
        Set
            Me._plate_bending = Value
        End Set
    End Property
    <Category("Bolt Groups"), Description(""), DisplayName("Grout Considered")>
    <DataMember()> Public Property grout_considered() As Boolean?
        Get
            Return Me._grout_considered
        End Get
        Set
            Me._grout_considered = Value
        End Set
    End Property
    <Category("Bolt Groups"), Description(""), DisplayName("Apply Barb Elevation")>
    <DataMember()> Public Property apply_barb_elevation() As Boolean?
        Get
            Return Me._apply_barb_elevation
        End Get
        Set
            Me._apply_barb_elevation = Value
        End Set
    End Property
    <Category("Bolt Groups"), Description(""), DisplayName("Bolt Name")>
    <DataMember()> Public Property bolt_name() As String
        Get
            Return Me._bolt_name
        End Get
        Set
            Me._bolt_name = Value
        End Set
    End Property


#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal bgrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = bgrow
        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_plate_id")))
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_id = DBtoNullableInt(dr.Item("local_bolt_group_id"))
            Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id"))
        End If
        'Me.connection_id = If(EDStruefalse, DBtoNullableInt(dr.Item("plate_id")), DBtoNullableInt(dr.Item("local_connection_id")))
        If EDStruefalse = True Then 'only pull in when referencing EDS
            Me.connection_id = DBtoNullableInt(dr.Item("plate_id"))
        End If
        Me.resist_axial = If(EDStruefalse, DBtoNullableBool(dr.Item("resist_axial")), If(DBtoStr(dr.Item("resist_axial")) = "Yes", True, If(DBtoStr(dr.Item("resist_axial")) = "No", False, DBtoNullableBool(dr.Item("resist_axial")))))
        Me.resist_shear = If(EDStruefalse, DBtoNullableBool(dr.Item("resist_shear")), If(DBtoStr(dr.Item("resist_shear")) = "Yes", True, If(DBtoStr(dr.Item("resist_shear")) = "No", False, DBtoNullableBool(dr.Item("resist_shear")))))
        Me.plate_bending = If(EDStruefalse, DBtoNullableBool(dr.Item("plate_bending")), If(DBtoStr(dr.Item("plate_bending")) = "Yes", True, If(DBtoStr(dr.Item("plate_bending")) = "No", False, DBtoNullableBool(dr.Item("plate_bending")))))
        Me.grout_considered = If(EDStruefalse, DBtoNullableBool(dr.Item("grout_considered")), If(DBtoStr(dr.Item("grout_considered")) = "Yes", True, If(DBtoStr(dr.Item("grout_considered")) = "No", False, DBtoNullableBool(dr.Item("grout_considered")))))
        Me.apply_barb_elevation = If(EDStruefalse, DBtoNullableBool(dr.Item("apply_barb_elevation")), If(DBtoStr(dr.Item("apply_barb_elevation")) = "Yes", True, If(DBtoStr(dr.Item("apply_barb_elevation")) = "No", False, DBtoNullableBool(dr.Item("apply_barb_elevation")))))
        Me.bolt_name = If(EDStruefalse, DBtoStr(dr.Item("bolt_name")), DBtoStr(dr.Item("group_name")))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.resist_axial.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.resist_shear.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_bending.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.grout_considered.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.apply_barb_elevation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_name.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("resist_axial")
        SQLInsertFields = SQLInsertFields.AddtoDBString("resist_shear")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_bending")
        SQLInsertFields = SQLInsertFields.AddtoDBString("grout_considered")
        SQLInsertFields = SQLInsertFields.AddtoDBString("apply_barb_elevation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_name")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_id = " & "@SubLevel1ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("resist_axial = " & Me.resist_axial.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("resist_shear = " & Me.resist_shear.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_bending = " & Me.plate_bending.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("grout_considered = " & Me.grout_considered.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("apply_barb_elevation = " & Me.apply_barb_elevation.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_name = " & Me.bolt_name.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As BoltGroup = TryCast(other, BoltGroup)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        'Equals = If(Me.plate_id.CheckChange(otherToCompare.plate_id, changes, categoryName, "Plate Id"), Equals, False)
        Equals = If(Me.resist_axial.CheckChange(otherToCompare.resist_axial, changes, categoryName, "Resist Axial"), Equals, False)
        Equals = If(Me.resist_shear.CheckChange(otherToCompare.resist_shear, changes, categoryName, "Resist Shear"), Equals, False)
        Equals = If(Me.plate_bending.CheckChange(otherToCompare.plate_bending, changes, categoryName, "Plate Bending"), Equals, False)
        Equals = If(Me.grout_considered.CheckChange(otherToCompare.grout_considered, changes, categoryName, "Grout Considered"), Equals, False)
        Equals = If(Me.apply_barb_elevation.CheckChange(otherToCompare.apply_barb_elevation, changes, categoryName, "Apply Barb Elevation"), Equals, False)
        Equals = If(Me.bolt_name.CheckChange(otherToCompare.bolt_name, changes, categoryName, "Bolt Name"), Equals, False)

        'Bolt Details
        If Me.BoltDetails.Count > 0 Then
            Equals = If(Me.BoltDetails.CheckChange(otherToCompare.BoltDetails, changes, categoryName, "Bolt Details"), Equals, False)
        End If

    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class BoltDetail
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Bolt Details"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.bolt_details"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Detail (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Detail_INSERT
        SQLInsert = SQLInsert.Replace("[BOLT DETAIL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[BOLT DETAIL FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        'Plate Material
        For Each row As CCIplateMaterial In CCIplateMaterials
            SQLInsert = SQLInsert.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert)
        Next

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Detail (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Detail_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        'Plate Material
        For Each row As CCIplateMaterial In CCIplateMaterials
            SQLUpdate = SQLUpdate.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert) 'Can only insert materials, no deleting or updating since database is referenced by all BUs. 
        Next

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Detail (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Detail_DELETE
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete

    End Function

#End Region

#Region "Define"
    'Private _local_id As Integer?
    Private _local_connection_id As Integer?
    Private _local_group_id As Integer?
    Private _bolt_group_id As Integer?
    Private _bolt_location As Double?
    Private _bolt_diameter As Double?
    Private _bolt_material As Integer?
    Private _bolt_circle As Double?
    Private _eta_factor As Double?
    Private _lar As Double?
    Private _bolt_thread_type As String
    Private _area_override As Double?
    Private _tension_only As Boolean?

    <DataMember()> Public Property CCIplateMaterials As New List(Of CCIplateMaterial)

    '<Category("Bolt Details"), Description(""), DisplayName("Local Id")>
    '<DataMember()> Public Property local_id() As Integer?
    '    Get
    '        Return Me._local_id
    '    End Get
    '    Set
    '        Me._local_id = Value
    '    End Set
    'End Property
    <Category("Bolt Details"), Description(""), DisplayName("Local Connection Id")>
    <DataMember()> Public Property local_connection_id() As Integer?
        Get
            Return Me._local_connection_id
        End Get
        Set
            Me._local_connection_id = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Local Group Id")>
    <DataMember()> Public Property local_group_id() As Integer?
        Get
            Return Me._local_group_id
        End Get
        Set
            Me._local_group_id = Value
        End Set
    End Property

    <Category("Bolt Details"), Description(""), DisplayName("Bolt Group Id")>
    <DataMember()> Public Property bolt_group_id() As Integer?
        Get
            Return Me._bolt_group_id
        End Get
        Set
            Me._bolt_group_id = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Bolt Location")>
    <DataMember()> Public Property bolt_location() As Double?
        Get
            Return Me._bolt_location
        End Get
        Set
            Me._bolt_location = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Bolt Diameter")>
    <DataMember()> Public Property bolt_diameter() As Double?
        Get
            Return Me._bolt_diameter
        End Get
        Set
            Me._bolt_diameter = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Bolt Material")>
    <DataMember()> Public Property bolt_material() As Integer?
        Get
            Return Me._bolt_material
        End Get
        Set
            Me._bolt_material = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Bolt Circle")>
    <DataMember()> Public Property bolt_circle() As Double?
        Get
            Return Me._bolt_circle
        End Get
        Set
            Me._bolt_circle = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Eta Factor")>
    <DataMember()> Public Property eta_factor() As Double?
        Get
            Return Me._eta_factor
        End Get
        Set
            Me._eta_factor = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Lar")>
    <DataMember()> Public Property lar() As Double?
        Get
            Return Me._lar
        End Get
        Set
            Me._lar = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Bolt Thread Type")>
    <DataMember()> Public Property bolt_thread_type() As String
        Get
            Return Me._bolt_thread_type
        End Get
        Set
            Me._bolt_thread_type = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Area Override")>
    <DataMember()> Public Property area_override() As Double?
        Get
            Return Me._area_override
        End Get
        Set
            Me._area_override = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Tension Only")>
    <DataMember()> Public Property tension_only() As Boolean?
        Get
            Return Me._tension_only
        End Get
        Set
            Me._tension_only = Value
        End Set
    End Property


#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal bdrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = bdrow
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            'Me.local_id = DBtoNullableInt(dr.Item("local_connection_id")) 'nothing references bolt details so deactivating
            Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id"))
            Me.local_group_id = DBtoNullableInt(dr.Item("local_group_id"))
        End If
        'Me.bolt_id = If(EDStruefalse, DBtoNullableInt(dr.Item("bolt_id")), DBtoNullableInt(dr.Item("local_group_id")))
        Me.bolt_group_id = If(EDStruefalse, DBtoNullableInt(dr.Item("bolt_id")), DBtoNullableInt(dr.Item("group_id")))
        Me.bolt_location = DBtoNullableDbl(dr.Item("bolt_location"))
        Me.bolt_diameter = DBtoNullableDbl(dr.Item("bolt_diameter"))
        Me.bolt_material = DBtoNullableInt(dr.Item("bolt_material"))
        Me.bolt_circle = DBtoNullableDbl(dr.Item("bolt_circle"))
        Me.eta_factor = DBtoNullableDbl(dr.Item("eta_factor"))
        Me.lar = DBtoNullableDbl(dr.Item("lar"))
        Me.bolt_thread_type = DBtoStr(dr.Item("bolt_thread_type"))
        Me.area_override = DBtoNullableDbl(dr.Item("area_override"))
        'When data is coming from Excel, blank data will report nothing. 
        Me.tension_only = If(EDStruefalse, DBtoNullableBool(dr.Item("tension_only")), If(DBtoStr(dr.Item("tension_only")) = "Yes", True, If(DBtoStr(dr.Item("tension_only")) = "No", False, DBtoNullableBool(dr.Item("tension_only")))))


    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_location.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_diameter.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_material.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_circle.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.eta_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.lar.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_thread_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.area_override.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tension_only.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_location")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_material")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_circle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("eta_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("lar")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_thread_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("area_override")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tension_only")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_id = " & "@SubLevel2ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_location = " & Me.bolt_location.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_diameter = " & Me.bolt_diameter.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_material = " & Me.bolt_material.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_material = " & "@SubLevel3ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_circle = " & Me.bolt_circle.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("eta_factor = " & Me.eta_factor.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("lar = " & Me.lar.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_thread_type = " & Me.bolt_thread_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("area_override = " & Me.area_override.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tension_only = " & Me.tension_only.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As BoltDetail = TryCast(other, BoltDetail)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.bolt_group_id.CheckChange(otherToCompare.bolt_group_id, changes, categoryName, "Bolt Group Id"), Equals, False)
        Equals = If(Me.bolt_location.CheckChange(otherToCompare.bolt_location, changes, categoryName, "Bolt Location"), Equals, False)
        Equals = If(Me.bolt_diameter.CheckChange(otherToCompare.bolt_diameter, changes, categoryName, "Bolt Diameter"), Equals, False)
        Equals = If(Me.bolt_material.CheckChange(otherToCompare.bolt_material, changes, categoryName, "Bolt Material"), Equals, False)
        Equals = If(Me.bolt_circle.CheckChange(otherToCompare.bolt_circle, changes, categoryName, "Bolt Circle"), Equals, False)
        Equals = If(Me.eta_factor.CheckChange(otherToCompare.eta_factor, changes, categoryName, "Eta Factor"), Equals, False)
        Equals = If(Me.lar.CheckChange(otherToCompare.lar, changes, categoryName, "Lar"), Equals, False)
        Equals = If(Me.bolt_thread_type.CheckChange(otherToCompare.bolt_thread_type, changes, categoryName, "Bolt Thread Type"), Equals, False)
        Equals = If(Me.area_override.CheckChange(otherToCompare.area_override, changes, categoryName, "Area Override"), Equals, False)
        Equals = If(Me.tension_only.CheckChange(otherToCompare.tension_only, changes, categoryName, "Tension Only"), Equals, False)

        'Materials
        If Me.CCIplateMaterials.Count > 0 Then
            Equals = If(Me.CCIplateMaterials.CheckChange(otherToCompare.CCIplateMaterials, changes, categoryName, "CCIplate Materials"), Equals, False)
        End If

    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class CCIplateMaterial
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "CCIplate Materials"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "gen.connection_material_properties"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\CCIplate Material (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Material_INSERT
        SQLInsert = SQLInsert.Replace("[MATERIAL PROPERTY ID]", Me.ID.ToString.FormatDBValue)
        SQLInsert = SQLInsert.Replace("[SELECT]", Me.SQLUpdateFieldsandValues)
        SQLInsert = SQLInsert.Replace("[CCIPLATE MATERIAL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[CCIPLATE MATERIAL FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert

    End Function

#End Region

#Region "Define"
    Private _local_id As Integer? 'removed
    Private _name As String
    Private _fy_0 As Double?
    Private _fy_1_125 As Double?
    Private _fy_1_625 As Double?
    Private _fy_2_625 As Double?
    Private _fy_4_125 As Double?
    Private _fu_0 As Double?
    Private _fu_1_125 As Double?
    Private _fu_1_625 As Double?
    Private _fu_2_625 As Double?
    Private _fu_4_125 As Double?
    Private _default_material As Boolean?

    <Category("Connection Material Properties"), Description(""), DisplayName("Local Id")>
    <DataMember()> Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Name")>
    <DataMember()> Public Property name() As String
        Get
            Return Me._name
        End Get
        Set
            Me._name = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 0")>
    <DataMember()> Public Property fy_0() As Double?
        Get
            Return Me._fy_0
        End Get
        Set
            Me._fy_0 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 1 125")>
    <DataMember()> Public Property fy_1_125() As Double?
        Get
            Return Me._fy_1_125
        End Get
        Set
            Me._fy_1_125 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 1 625")>
    <DataMember()> Public Property fy_1_625() As Double?
        Get
            Return Me._fy_1_625
        End Get
        Set
            Me._fy_1_625 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 2 625")>
    <DataMember()> Public Property fy_2_625() As Double?
        Get
            Return Me._fy_2_625
        End Get
        Set
            Me._fy_2_625 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 4 125")>
    <DataMember()> Public Property fy_4_125() As Double?
        Get
            Return Me._fy_4_125
        End Get
        Set
            Me._fy_4_125 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 0")>
    <DataMember()> Public Property fu_0() As Double?
        Get
            Return Me._fu_0
        End Get
        Set
            Me._fu_0 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 1 125")>
    <DataMember()> Public Property fu_1_125() As Double?
        Get
            Return Me._fu_1_125
        End Get
        Set
            Me._fu_1_125 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 1 625")>
    <DataMember()> Public Property fu_1_625() As Double?
        Get
            Return Me._fu_1_625
        End Get
        Set
            Me._fu_1_625 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 2 625")>
    <DataMember()> Public Property fu_2_625() As Double?
        Get
            Return Me._fu_2_625
        End Get
        Set
            Me._fu_2_625 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 4 125")>
    <DataMember()> Public Property fu_4_125() As Double?
        Get
            Return Me._fu_4_125
        End Get
        Set
            Me._fu_4_125 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Default Material")>
    <DataMember()> Public Property default_material() As Boolean?
        Get
            Return Me._default_material
        End Get
        Set
            Me._default_material = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal mrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing) '(ByVal mrow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = mrow
        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_id")))
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_id = DBtoNullableInt(dr.Item("local_material_id"))
        End If
        Me.name = DBtoStr(dr.Item("name"))
        Me.fy_0 = DBtoNullableDbl(dr.Item("fy_0"))
        Me.fy_1_125 = DBtoNullableDbl(dr.Item("fy_1_125"))
        Me.fy_1_625 = DBtoNullableDbl(dr.Item("fy_1_625"))
        Me.fy_2_625 = DBtoNullableDbl(dr.Item("fy_2_625"))
        Me.fy_4_125 = DBtoNullableDbl(dr.Item("fy_4_125"))
        Me.fu_0 = DBtoNullableDbl(dr.Item("fu_0"))
        Me.fu_1_125 = DBtoNullableDbl(dr.Item("fu_1_125"))
        Me.fu_1_625 = DBtoNullableDbl(dr.Item("fu_1_625"))
        Me.fu_2_625 = DBtoNullableDbl(dr.Item("fu_2_625"))
        Me.fu_4_125 = DBtoNullableDbl(dr.Item("fu_4_125"))
        Me.default_material = If(EDStruefalse, DBtoNullableBool(dr.Item("default_material")), False)


    End Sub

    Public Sub New(ByVal ID As Integer?, Optional ByVal name As String = Nothing, Optional ByVal fy_0 As Double? = Nothing, Optional ByVal fu_0 As Double? = Nothing)
        'This is used to store a temp list of new materials to add to the Excel tool
        Me.ID = ID
        Me.name = name
        Me.fy_0 = fy_0
        Me.fu_0 = fu_0
    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.name.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy_0.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy_1_125.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy_1_625.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy_2_625.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fy_4_125.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_0.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_1_125.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_1_625.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_2_625.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.fu_4_125.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.default_material.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("local_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("name")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fy_0")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fy_1_125")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fy_1_625")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fy_2_625")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fy_4_125")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_0")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_1_125")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_1_625")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_2_625")
        SQLInsertFields = SQLInsertFields.AddtoDBString("fu_4_125")
        SQLInsertFields = SQLInsertFields.AddtoDBString("default_material")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("local_id = " & Me.local_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("name = " & Me.name.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fy_0), "fy_0 = " & Me.fy_0.ToString.FormatDBValue, "fy_0 IS NULL "))
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fy_1_125), "fy_1_125 = " & Me.fy_1_125.ToString.FormatDBValue, "fy_1_125 IS NULL "))
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fy_1_625), "fy_1_625 = " & Me.fy_1_625.ToString.FormatDBValue, "fy_1_625 IS NULL "))
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fy_2_625), "fy_2_625 = " & Me.fy_2_625.ToString.FormatDBValue, "fy_2_625 IS NULL "))
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fy_4_125), "fy_4_125 = " & Me.fy_4_125.ToString.FormatDBValue, "fy_4_125 IS NULL "))
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fu_0), "fu_0 = " & Me.fu_0.ToString.FormatDBValue, "fu_0 IS NULL "))
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fu_1_125), "fu_1_125 = " & Me.fu_1_125.ToString.FormatDBValue, "fu_1_125 IS NULL "))
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fu_1_625), "fu_1_625 = " & Me.fu_1_625.ToString.FormatDBValue, "fu_1_625 IS NULL "))
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fu_2_625), "fu_2_625 = " & Me.fu_2_625.ToString.FormatDBValue, "fu_2_625 IS NULL "))
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString(If(IsSomething(Me.fu_4_125), "fu_4_125 = " & Me.fu_4_125.ToString.FormatDBValue, "fu_4_125 IS NULL "))

        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("default_material = " & Me.default_material.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As CCIplateMaterial = TryCast(other, CCIplateMaterial)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        'Equals = If(Me.local_id.CheckChange(otherToCompare.local_id, changes, categoryName, "Local Id"), Equals, False)
        Equals = If(Me.name.CheckChange(otherToCompare.name, changes, categoryName, "Name"), Equals, False)
        Equals = If(Me.fy_0.CheckChange(otherToCompare.fy_0, changes, categoryName, "Fy 0"), Equals, False)
        Equals = If(Me.fy_1_125.CheckChange(otherToCompare.fy_1_125, changes, categoryName, "Fy 1 125"), Equals, False)
        Equals = If(Me.fy_1_625.CheckChange(otherToCompare.fy_1_625, changes, categoryName, "Fy 1 625"), Equals, False)
        Equals = If(Me.fy_2_625.CheckChange(otherToCompare.fy_2_625, changes, categoryName, "Fy 2 625"), Equals, False)
        Equals = If(Me.fy_4_125.CheckChange(otherToCompare.fy_4_125, changes, categoryName, "Fy 4 125"), Equals, False)
        Equals = If(Me.fu_0.CheckChange(otherToCompare.fu_0, changes, categoryName, "Fu 0"), Equals, False)
        Equals = If(Me.fu_1_125.CheckChange(otherToCompare.fu_1_125, changes, categoryName, "Fu 1 125"), Equals, False)
        Equals = If(Me.fu_1_625.CheckChange(otherToCompare.fu_1_625, changes, categoryName, "Fu 1 625"), Equals, False)
        Equals = If(Me.fu_2_625.CheckChange(otherToCompare.fu_2_625, changes, categoryName, "Fu 2 625"), Equals, False)
        Equals = If(Me.fu_4_125.CheckChange(otherToCompare.fu_4_125, changes, categoryName, "Fu 4 125"), Equals, False)
        'Equals = If(Me.default_material.CheckChange(otherToCompare.default_material, changes, categoryName, "Default Material"), Equals, False)


    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class PlateResults
    Inherits EDSObjectWithQueries
    'Inherits EDSResult

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Plate Results"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.plate_results"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Plate Result (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Plate_Result_INSERT
        SQLInsert = SQLInsert.Replace("[PLATE RESULT VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[PLATE RESULT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert

    End Function

#End Region

#Region "Define"
    Private _plate_details_id As Integer?
    Private _local_plate_id As Integer?
    'Private _work_order_seq_num As Double? 'not provided in Excel
    Private _rating As Decimal?
    Private _result_lkup As String
    'Private _modified_person_id As Integer? 'not provided in Excel
    'Private _process_stage As String 'not provided in Excel
    'Private _modified_date As DateTime? 'not provided in Excel

    <Category("Plate Results"), Description(""), DisplayName("Plate Details Id")>
    <DataMember()> Public Property plate_details_id() As Integer?
        Get
            Return Me._plate_details_id
        End Get
        Set
            Me._plate_details_id = Value
        End Set
    End Property
    <Category("Plate Results"), Description(""), DisplayName("Local Plate Id")>
    <DataMember()> Public Property local_plate_id() As Integer?
        Get
            Return Me._local_plate_id
        End Get
        Set
            Me._local_plate_id = Value
        End Set
    End Property
    '<Category("Plate Results"), Description(""), DisplayName("Work Order Seq Num")>
    '<DataMember()> Public Property work_order_seq_num() As Double?
    '    Get
    '        Return Me._work_order_seq_num
    '    End Get
    '    Set
    '        Me._work_order_seq_num = Value
    '    End Set
    'End Property
    <Category("Plate Results"), Description(""), DisplayName("Rating")>
    <DataMember()> Public Property rating() As Decimal?
        Get
            Return Me._rating
        End Get
        Set
            Me._rating = Value
        End Set
    End Property
    <Category("Plate Results"), Description(""), DisplayName("Result Lkup")>
    <DataMember()> Public Property result_lkup() As String
        Get
            Return Me._result_lkup
        End Get
        Set
            Me._result_lkup = Value
        End Set
    End Property
    '<Category("Plate Results"), Description(""), DisplayName("Modified Person Id")>
    '<DataMember()> Public Property modified_person_id() As Integer?
    '    Get
    '        Return Me._modified_person_id
    '    End Get
    '    Set
    '        Me._modified_person_id = Value
    '    End Set
    'End Property
    '<Category("Plate Results"), Description(""), DisplayName("Process Stage")>
    '<DataMember()> Public Property process_stage() As String
    '    Get
    '        Return Me._process_stage
    '    End Get
    '    Set
    '        Me._process_stage = Value
    '    End Set
    'End Property
    '<Category("Plate Results"), Description(""), DisplayName("Modified Date")>
    '<DataMember()> Public Property modified_date() As DateTime?
    '    Get
    '        Return Me._modified_date
    '    End Get
    '    Set
    '        Me._modified_date = Value
    '    End Set
    'End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal prrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = prrow

        Me.plate_details_id = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_plate_id = DBtoNullableInt(dr.Item("local_plate_id"))
        End If
        'Me.work_order_seq_num = DBtoNullableDbl(dr.Item("work_order_seq_num"))
        Me.rating = DBtoNullableDec(dr.Item("rating"))
        Me.result_lkup = DBtoStr(dr.Item("result_lkup"))
        'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
        'Me.process_stage = DBtoStr(dr.Item("process_stage"))
        'Me.modified_date = DBtoStr(dr.Item("modified_date"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_details_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_date.ToString.FormatDBValue)


        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_details_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_date")


        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_details_id = " & Me.plate_details_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rating = " & Me.rating.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_lkup = " & Me.result_lkup.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_date = " & Me.modified_date.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As PlateResults = TryCast(other, PlateResults)
        If otherToCompare Is Nothing Then Return False


    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class BoltResults
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Bolt Results"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.bolt_results"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Bolt Result (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Bolt_Result_INSERT
        SQLInsert = SQLInsert.Replace("[BOLT RESULT VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[BOLT RESULT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert

    End Function

#End Region

#Region "Define"
    Private _bolt_id As Integer?
    Private _local_connection_id As Integer?
    Private _local_bolt_group_id As Integer?
    'Private _work_order_seq_num As Double? 'not provided in Excel
    Private _rating As Decimal?
    Private _result_lkup As String
    'Private _modified_person_id As Integer? 'not provided in Excel
    'Private _process_stage As String 'not provided in Excel
    'Private _modified_date As DateTime? 'not provided in Excel

    <Category("Bolt Results"), Description(""), DisplayName("Bolt Id")>
    <DataMember()> Public Property bolt_id() As Integer?
        Get
            Return Me._bolt_id
        End Get
        Set
            Me._bolt_id = Value
        End Set
    End Property
    <Category("Bolt Results"), Description(""), DisplayName("Local Connection Id")>
    <DataMember()> Public Property local_connection_id() As Integer?
        Get
            Return Me._local_connection_id
        End Get
        Set
            Me._local_connection_id = Value
        End Set
    End Property
    <Category("Bolt Results"), Description(""), DisplayName("Local Bolt Group Id")>
    <DataMember()> Public Property local_bolt_group_id() As Integer?
        Get
            Return Me._local_bolt_group_id
        End Get
        Set
            Me._local_bolt_group_id = Value
        End Set
    End Property
    '<Category("Bolt Results"), Description(""), DisplayName("Work Order Seq Num")>
    '<DataMember()> Public Property work_order_seq_num() As Double?
    '    Get
    '        Return Me._work_order_seq_num
    '    End Get
    '    Set
    '        Me._work_order_seq_num = Value
    '    End Set
    'End Property
    <Category("Bolt Results"), Description(""), DisplayName("Rating")>
    <DataMember()> Public Property rating() As Decimal?
        Get
            Return Me._rating
        End Get
        Set
            Me._rating = Value
        End Set
    End Property
    <Category("Bolt Results"), Description(""), DisplayName("Result Lkup")>
    <DataMember()> Public Property result_lkup() As String
        Get
            Return Me._result_lkup
        End Get
        Set
            Me._result_lkup = Value
        End Set
    End Property
    '<Category("Bolt Results"), Description(""), DisplayName("Modified Person Id")>
    '<DataMember()> Public Property modified_person_id() As Integer?
    '    Get
    '        Return Me._modified_person_id
    '    End Get
    '    Set
    '        Me._modified_person_id = Value
    '    End Set
    'End Property
    '<Category("Bolt Results"), Description(""), DisplayName("Process Stage")>
    '<DataMember()> Public Property process_stage() As String
    '    Get
    '        Return Me._process_stage
    '    End Get
    '    Set
    '        Me._process_stage = Value
    '    End Set
    'End Property
    '<Category("Bolt Results"), Description(""), DisplayName("Modified Date")>
    '<DataMember()> Public Property modified_date() As DateTime?
    '    Get
    '        Return Me._modified_date
    '    End Get
    '    Set
    '        Me._modified_date = Value
    '    End Set
    'End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal brrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = brrow

        Me.bolt_id = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id"))
            Me.local_bolt_group_id = DBtoNullableInt(dr.Item("local_bolt_group_id"))
        End If
        'Me.work_order_seq_num = DBtoNullableDbl(dr.Item("work_order_seq_num"))
        Me.rating = DBtoNullableDec(dr.Item("rating")) 'same in all 
        Me.result_lkup = DBtoStr(dr.Item("result_lkup")) 'same in all
        'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
        'Me.process_stage = DBtoStr(dr.Item("process_stage"))
        'Me.modified_date = DBtoStr(dr.Item("modified_date"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_details_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_date.ToString.FormatDBValue)


        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_date")


        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_id = " & Me.bolt_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rating = " & Me.rating.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_lkup = " & Me.result_lkup.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_date = " & Me.modified_date.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As BoltResults = TryCast(other, BoltResults)
        If otherToCompare Is Nothing Then Return False


    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class StiffenerGroup
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Stiffener Groups"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.stiffeners"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Group (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Group_INSERT
        SQLInsert = SQLInsert.Replace("[STIFFENER GROUP VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[STIFFENER GROUP FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        'Bolt Detail
        If Me.StiffenerDetails.Count > 0 Then
            SQLInsert = SQLInsert.Replace("--BEGIN --[STIFFENER DETAIL INSERT BEGIN]", "BEGIN --[STIFFENER DETAIL INSERT BEGIN]")
            SQLInsert = SQLInsert.Replace("--END --[STIFFENER DETAIL INSERT END]", "END --[STIFFENER DETAIL INSERT END]")
            For Each row As StiffenerDetail In StiffenerDetails
                SQLInsert = SQLInsert.Replace("--[STIFFENER DETAIL INSERT]", row.SQLInsert)
            Next
        End If

        ''Results
        'For Each row As StiffenerResults In StiffenerResults
        '    SQLInsert = SQLInsert.Replace("--BEGIN --[STIFFENER RESULTS INSERT BEGIN]", "BEGIN --[STIFFENER RESULTS INSERT BEGIN]")
        '    SQLInsert = SQLInsert.Replace("--END --[STIFFENER RESULTS INSERT END]", "END --[STIFFENER RESULTS INSERT END]")
        '    SQLInsert = SQLInsert.Replace("--[STIFFENER RESULTS INSERT]", row.SQLInsert)
        'Next

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Group (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Group_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        'Stiffener Detail
        If Me.StiffenerDetails.Count > 0 Then
            SQLUpdate = SQLUpdate.Replace("--BEGIN --[STIFFENER DETAIL UPDATE BEGIN]", "BEGIN --[STIFFENER DETAIL UPDATE BEGIN]")
            SQLUpdate = SQLUpdate.Replace("--END --[STIFFENER DETAIL UPDATE END]", "END --[STIFFENER DETAIL UPDATE END]")
            For Each row As StiffenerDetail In StiffenerDetails
                If IsSomething(row.ID) Then 'If ID exists within Excel, layer exists in EDS and either update or delete should be performed. Otherwise, insert new record. 
                    If IsSomething(row.stiffener_location) Or IsSomething(row.stiffener_width) Or IsSomething(row.stiffener_height) _
                        Or IsSomething(row.stiffener_thickness) Or IsSomething(row.stiffener_h_notch) Or IsSomething(row.stiffener_v_notch) _
                        Or IsSomething(row.stiffener_grade) Or IsSomethingString(row.weld_type) Or IsSomething(row.groove_depth) _
                        Or IsSomething(row.groove_angle) Or IsSomething(row.h_fillet_weld) Or IsSomething(row.v_fillet_weld) _
                        Or IsSomething(row.weld_strength) Then
                        SQLUpdate = SQLUpdate.Replace("--[STIFFENER DETAIL INSERT]", row.SQLUpdate)
                    Else
                        SQLUpdate = SQLUpdate.Replace("--[STIFFENER DETAIL INSERT]", row.SQLDelete)
                    End If
                Else
                    SQLUpdate = SQLUpdate.Replace("--[STIFFENER DETAIL INSERT]", row.SQLInsert)
                End If
            Next
        End If

        'For Each row As StiffenerResults In StiffenerResults
        '    SQLUpdate = SQLUpdate.Replace("--BEGIN --[STIFFENER RESULTS INSERT BEGIN]", "BEGIN --[STIFFENER RESULTS INSERT BEGIN]")
        '    SQLUpdate = SQLUpdate.Replace("--END --[STIFFENER RESULTS INSERT END]", "END --[STIFFENER RESULTS INSERT END]")
        '    SQLUpdate = SQLUpdate.Replace("--[STIFFENER RESULTS INSERT]", row.SQLInsert)
        'Next

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Group (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Group_DELETE
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        'Stiffener Details
        If Me.StiffenerDetails.Count > 0 Then
            SQLDelete = SQLDelete.Replace("--BEGIN --[STIFFENER DETAIL DELETE BEGIN]", "BEGIN --[STIFFENER DETAIL DELETE BEGIN]")
            SQLDelete = SQLDelete.Replace("--END --[STIFFENER DETAIL DELETE END]", "END --[STIFFENER DETAIL DELETE END]")
            For Each row As StiffenerDetail In StiffenerDetails
                SQLDelete = SQLDelete.Replace("--[STIFFENER DETAIL INSERT]", row.SQLDelete)
            Next
        End If

        Return SQLDelete

    End Function

#End Region

#Region "Define"
    Private _local_id As Integer?
    Private _local_plate_id As Integer?
    Private _plate_details_id As Integer?
    Private _stiffener_name As String

    <DataMember()> Public Property StiffenerDetails As New List(Of StiffenerDetail)
    '<DataMember()> Public Property StiffenerResults As New List(Of StiffenerResults)

    <Category("Stiffener Groups"), Description(""), DisplayName("Local Id")>
    <DataMember()> Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property
    <Category("Stiffener Groups"), Description(""), DisplayName("Local Plate Id")>
    <DataMember()> Public Property local_plate_id() As Integer?
        Get
            Return Me._local_plate_id
        End Get
        Set
            Me._local_plate_id = Value
        End Set
    End Property

    <Category("Stiffener Groups"), Description(""), DisplayName("Plate Details Id")>
    <DataMember()> Public Property plate_details_id() As Integer?
        Get
            Return Me._plate_details_id
        End Get
        Set
            Me._plate_details_id = Value
        End Set
    End Property
    <Category("Stiffener Groups"), Description(""), DisplayName("Stiffener Name")>
    <DataMember()> Public Property stiffener_name() As String
        Get
            Return Me._stiffener_name
        End Get
        Set
            Me._stiffener_name = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal sgrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = sgrow
        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_plate_id")))
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_id = DBtoNullableInt(dr.Item("local_stiffener_group_id"))
            Me.local_plate_id = DBtoNullableInt(dr.Item("local_plate_id"))
        End If
        'Me.plate_details_id = If(EDStruefalse, DBtoNullableInt(dr.Item("plate_details_id")), DBtoNullableInt(dr.Item("local_plate_id")))
        If EDStruefalse = True Then 'Only pull in when referencing EDS
            Me.plate_details_id = DBtoNullableInt(dr.Item("plate_details_id"))
        End If
        Me.stiffener_name = If(EDStruefalse, DBtoStr(dr.Item("stiffener_name")), DBtoStr(dr.Item("group_name")))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel2ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_details_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_name.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_details_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_name")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_details_id = " & "@SubLevel2ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_name = " & Me.stiffener_name.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As StiffenerGroup = TryCast(other, StiffenerGroup)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        'Equals = If(Me.plate_details_id.CheckChange(otherToCompare.plate_details_id, changes, categoryName, "Plate Details Id"), Equals, False)
        Equals = If(Me.stiffener_name.CheckChange(otherToCompare.stiffener_name, changes, categoryName, "Stiffener Name"), Equals, False)

        'Stiffener Details
        If Me.StiffenerDetails.Count > 0 Then
            Equals = If(Me.StiffenerDetails.CheckChange(otherToCompare.StiffenerDetails, changes, categoryName, "Stiffener Details"), Equals, False)
        End If

    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class StiffenerDetail
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Stiffener Details"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.stiffener_details"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Detail (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Detail_INSERT
        SQLInsert = SQLInsert.Replace("[STIFFENER DETAIL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[STIFFENER DETAIL FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Detail (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Detail_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Detail (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Stiffener_Detail_DELETE
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete

    End Function

#End Region

#Region "Define"
    'Private _local_id As Integer?
    Private _local_plate_id As Integer?
    Private _local_group_id As Integer?
    Private _stiffener_id As Integer?
    Private _stiffener_location As Double?
    Private _stiffener_width As Double?
    Private _stiffener_height As Double?
    Private _stiffener_thickness As Double?
    Private _stiffener_h_notch As Double?
    Private _stiffener_v_notch As Double?
    Private _stiffener_grade As Double?
    Private _weld_type As String
    Private _groove_depth As Double?
    Private _groove_angle As Double?
    Private _h_fillet_weld As Double?
    Private _v_fillet_weld As Double?
    Private _weld_strength As Double?


    <DataMember()> Public Property CCIplateMaterials As New List(Of CCIplateMaterial)

    '<Category("Bolt Details"), Description(""), DisplayName("Local Id")>
    '<DataMember()> Public Property local_id() As Integer?
    '    Get
    '        Return Me._local_id
    '    End Get
    '    Set
    '        Me._local_id = Value
    '    End Set
    'End Property
    <Category("Bolt Details"), Description(""), DisplayName("Local Plate Id")>
    <DataMember()> Public Property local_plate_id() As Integer?
        Get
            Return Me._local_plate_id
        End Get
        Set
            Me._local_plate_id = Value
        End Set
    End Property
    <Category("Bolt Details"), Description(""), DisplayName("Local Group Id")>
    <DataMember()> Public Property local_group_id() As Integer?
        Get
            Return Me._local_group_id
        End Get
        Set
            Me._local_group_id = Value
        End Set
    End Property

    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Id")>
    <DataMember()> Public Property stiffener_id() As Integer?
        Get
            Return Me._stiffener_id
        End Get
        Set
            Me._stiffener_id = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Location")>
    <DataMember()> Public Property stiffener_location() As Double?
        Get
            Return Me._stiffener_location
        End Get
        Set
            Me._stiffener_location = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Width")>
    <DataMember()> Public Property stiffener_width() As Double?
        Get
            Return Me._stiffener_width
        End Get
        Set
            Me._stiffener_width = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Height")>
    <DataMember()> Public Property stiffener_height() As Double?
        Get
            Return Me._stiffener_height
        End Get
        Set
            Me._stiffener_height = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Thickness")>
    <DataMember()> Public Property stiffener_thickness() As Double?
        Get
            Return Me._stiffener_thickness
        End Get
        Set
            Me._stiffener_thickness = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener H Notch")>
    <DataMember()> Public Property stiffener_h_notch() As Double?
        Get
            Return Me._stiffener_h_notch
        End Get
        Set
            Me._stiffener_h_notch = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener V Notch")>
    <DataMember()> Public Property stiffener_v_notch() As Double?
        Get
            Return Me._stiffener_v_notch
        End Get
        Set
            Me._stiffener_v_notch = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Stiffener Grade")>
    <DataMember()> Public Property stiffener_grade() As Double?
        Get
            Return Me._stiffener_grade
        End Get
        Set
            Me._stiffener_grade = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Weld Type")>
    <DataMember()> Public Property weld_type() As String
        Get
            Return Me._weld_type
        End Get
        Set
            Me._weld_type = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Groove Depth")>
    <DataMember()> Public Property groove_depth() As Double?
        Get
            Return Me._groove_depth
        End Get
        Set
            Me._groove_depth = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Groove Angle")>
    <DataMember()> Public Property groove_angle() As Double?
        Get
            Return Me._groove_angle
        End Get
        Set
            Me._groove_angle = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("H Fillet Weld")>
    <DataMember()> Public Property h_fillet_weld() As Double?
        Get
            Return Me._h_fillet_weld
        End Get
        Set
            Me._h_fillet_weld = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("V Fillet Weld")>
    <DataMember()> Public Property v_fillet_weld() As Double?
        Get
            Return Me._v_fillet_weld
        End Get
        Set
            Me._v_fillet_weld = Value
        End Set
    End Property
    <Category("Stiffener Details"), Description(""), DisplayName("Weld Strength")>
    <DataMember()> Public Property weld_strength() As Double?
        Get
            Return Me._weld_strength
        End Get
        Set
            Me._weld_strength = Value
        End Set
    End Property



#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal bdrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = bdrow
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            'Me.local_id = DBtoNullableInt(dr.Item("local_id")) 'nothing references stiffener details so deactivating
            Me.local_plate_id = DBtoNullableInt(dr.Item("local_plate_id"))
            Me.local_group_id = DBtoNullableInt(dr.Item("local_group_id"))
        End If
        Me.stiffener_id = If(EDStruefalse, DBtoNullableInt(dr.Item("stiffener_id")), DBtoNullableInt(dr.Item("group_id")))
        Me.stiffener_location = DBtoNullableDbl(dr.Item("stiffener_location"))
        Me.stiffener_width = DBtoNullableDbl(dr.Item("stiffener_width"))
        Me.stiffener_height = DBtoNullableDbl(dr.Item("stiffener_height"))
        Me.stiffener_thickness = DBtoNullableDbl(dr.Item("stiffener_thickness"))
        Me.stiffener_h_notch = DBtoNullableDbl(dr.Item("stiffener_h_notch"))
        Me.stiffener_v_notch = DBtoNullableDbl(dr.Item("stiffener_v_notch"))
        Me.stiffener_grade = DBtoNullableDbl(dr.Item("stiffener_grade"))
        Me.weld_type = DBtoStr(dr.Item("weld_type"))
        Me.groove_depth = DBtoNullableDbl(dr.Item("groove_depth"))
        Me.groove_angle = DBtoNullableDbl(dr.Item("groove_angle"))
        Me.h_fillet_weld = DBtoNullableDbl(dr.Item("h_fillet_weld"))
        Me.v_fillet_weld = DBtoNullableDbl(dr.Item("v_fillet_weld"))
        Me.weld_strength = DBtoNullableDbl(dr.Item("weld_strength"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_location.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_width.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_height.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_h_notch.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_v_notch.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_grade.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.groove_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.groove_angle.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.h_fillet_weld.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.v_fillet_weld.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_strength.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_location")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_width")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_height")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_h_notch")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_v_notch")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_grade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("groove_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("groove_angle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("h_fillet_weld")
        SQLInsertFields = SQLInsertFields.AddtoDBString("v_fillet_weld")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_strength")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_id = " & "@SubLevel3ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_location = " & Me.stiffener_location.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_width = " & Me.stiffener_width.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_height = " & Me.stiffener_height.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_thickness = " & Me.stiffener_thickness.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_h_notch = " & Me.stiffener_h_notch.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_v_notch = " & Me.stiffener_v_notch.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_grade = " & Me.stiffener_grade.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_type = " & Me.weld_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("groove_depth = " & Me.groove_depth.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("groove_angle = " & Me.groove_angle.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("h_fillet_weld = " & Me.h_fillet_weld.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("v_fillet_weld = " & Me.v_fillet_weld.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_strength = " & Me.weld_strength.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As StiffenerDetail = TryCast(other, StiffenerDetail)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        Equals = If(Me.stiffener_id.CheckChange(otherToCompare.stiffener_id, changes, categoryName, "Stiffener Id"), Equals, False)
        Equals = If(Me.stiffener_location.CheckChange(otherToCompare.stiffener_location, changes, categoryName, "Stiffener Location"), Equals, False)
        Equals = If(Me.stiffener_width.CheckChange(otherToCompare.stiffener_width, changes, categoryName, "Stiffener Width"), Equals, False)
        Equals = If(Me.stiffener_height.CheckChange(otherToCompare.stiffener_height, changes, categoryName, "Stiffener Height"), Equals, False)
        Equals = If(Me.stiffener_thickness.CheckChange(otherToCompare.stiffener_thickness, changes, categoryName, "Stiffener Thickness"), Equals, False)
        Equals = If(Me.stiffener_h_notch.CheckChange(otherToCompare.stiffener_h_notch, changes, categoryName, "Stiffener H Notch"), Equals, False)
        Equals = If(Me.stiffener_v_notch.CheckChange(otherToCompare.stiffener_v_notch, changes, categoryName, "Stiffener V Notch"), Equals, False)
        Equals = If(Me.stiffener_grade.CheckChange(otherToCompare.stiffener_grade, changes, categoryName, "Stiffener Grade"), Equals, False)
        Equals = If(Me.weld_type.CheckChange(otherToCompare.weld_type, changes, categoryName, "Weld Type"), Equals, False)
        Equals = If(Me.groove_depth.CheckChange(otherToCompare.groove_depth, changes, categoryName, "Groove Depth"), Equals, False)
        Equals = If(Me.groove_angle.CheckChange(otherToCompare.groove_angle, changes, categoryName, "Groove Angle"), Equals, False)
        Equals = If(Me.h_fillet_weld.CheckChange(otherToCompare.h_fillet_weld, changes, categoryName, "H Fillet Weld"), Equals, False)
        Equals = If(Me.v_fillet_weld.CheckChange(otherToCompare.v_fillet_weld, changes, categoryName, "V Fillet Weld"), Equals, False)
        Equals = If(Me.weld_strength.CheckChange(otherToCompare.weld_strength, changes, categoryName, "Weld Strength"), Equals, False)

    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class StiffenerResults
    Inherits EDSObjectWithQueries
    'StiffenerResults are currently not being referenced. Stiffeners reported with plate details. 
#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Stiffener Results"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.stiffener_results"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Stiffener Result (INSERT).sql")
        'SQLInsert = CCI_Engineering_Templates.My.Resources.Stiffener_Result_INSERT
        SQLInsert = SQLInsert.Replace("[STIFFENER RESULT VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[STIFFENER RESULT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert

    End Function

    'Public Overrides Function SQLUpdate() As String

    '    SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Connection Material (UPDATE).sql")
    '    SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
    '    SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
    '    SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

    '    Return SQLUpdate

    'End Function

    'Public Overrides Function SQLDelete() As String

    '    SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Connection Material (DELETE).sql")
    '    SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
    '    SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

    '    Return SQLDelete

    'End Function

#End Region

#Region "Define"
    Private _stiffener_id As Integer?
    Private _local_id As Integer?
    Private _local_stiffener_group_id As Integer?
    'Private _work_order_seq_num As Double? 'not provided in Excel
    Private _rating As Decimal?
    Private _result_lkup As String
    'Private _modified_person_id As Integer? 'not provided in Excel
    'Private _process_stage As String 'not provided in Excel
    'Private _modified_date As DateTime? 'not provided in Excel

    <Category("Stiffener Results"), Description(""), DisplayName("Stiffener Id")>
    <DataMember()> Public Property stiffener_id() As Integer?
        Get
            Return Me._stiffener_id
        End Get
        Set
            Me._stiffener_id = Value
        End Set
    End Property
    <Category("Stiffener Results"), Description(""), DisplayName("Local Id")>
    <DataMember()> Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property
    <Category("Stiffener Results"), Description(""), DisplayName("Local Bolt Group Id")>
    <DataMember()> Public Property local_stiffener_group_id() As Integer?
        Get
            Return Me._local_stiffener_group_id
        End Get
        Set
            Me._local_stiffener_group_id = Value
        End Set
    End Property
    '<Category("Stiffener Results"), Description(""), DisplayName("Work Order Seq Num")>
    '<DataMember()> Public Property work_order_seq_num() As Double?
    '    Get
    '        Return Me._work_order_seq_num
    '    End Get
    '    Set
    '        Me._work_order_seq_num = Value
    '    End Set
    'End Property
    <Category("Stiffener Results"), Description(""), DisplayName("Rating")>
    <DataMember()> Public Property rating() As Decimal?
        Get
            Return Me._rating
        End Get
        Set
            Me._rating = Value
        End Set
    End Property
    <Category("Stiffener Results"), Description(""), DisplayName("Result Lkup")>
    <DataMember()> Public Property result_lkup() As String
        Get
            Return Me._result_lkup
        End Get
        Set
            Me._result_lkup = Value
        End Set
    End Property
    '<Category("Stiffener Results"), Description(""), DisplayName("Modified Person Id")>
    '<DataMember()> Public Property modified_person_id() As Integer?
    '    Get
    '        Return Me._modified_person_id
    '    End Get
    '    Set
    '        Me._modified_person_id = Value
    '    End Set
    'End Property
    '<Category("Stiffener Results"), Description(""), DisplayName("Process Stage")>
    '<DataMember()> Public Property process_stage() As String
    '    Get
    '        Return Me._process_stage
    '    End Get
    '    Set
    '        Me._process_stage = Value
    '    End Set
    'End Property
    '<Category("Stiffener Results"), Description(""), DisplayName("Modified Date")>
    '<DataMember()> Public Property modified_date() As DateTime?
    '    Get
    '        Return Me._modified_date
    '    End Get
    '    Set
    '        Me._modified_date = Value
    '    End Set
    'End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal brrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = brrow

        Me.stiffener_id = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_id = DBtoNullableInt(dr.Item("local_id"))
            Me.local_stiffener_group_id = DBtoNullableInt(dr.Item("local_bolt_group_id"))
        End If
        'Me.work_order_seq_num = DBtoNullableDbl(dr.Item("work_order_seq_num"))
        Me.rating = DBtoNullableDec(dr.Item("rating")) 'same in all 
        Me.result_lkup = DBtoStr(dr.Item("result_lkup")) 'same in all
        'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
        'Me.process_stage = DBtoStr(dr.Item("process_stage"))
        'Me.modified_date = DBtoStr(dr.Item("modified_date"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_details_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_date.ToString.FormatDBValue)


        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_date")


        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_id = " & Me.stiffener_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rating = " & Me.rating.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_lkup = " & Me.result_lkup.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_date = " & Me.modified_date.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As StiffenerResults = TryCast(other, StiffenerResults)
        If otherToCompare Is Nothing Then Return False


    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class BridgeStiffenerDetail
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Bridge Stiffener Details"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.bridge_stiffeners"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Bridge Stiffener Detail (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Bridge_Stiffener_Detail_INSERT
        SQLInsert = SQLInsert.Replace("[BRIDGE STIFFENER DETAIL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[BRIDGE STIFFENER DETAIL FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        'Plate Material
        For Each row As CCIplateMaterial In CCIplateMaterials
            SQLInsert = SQLInsert.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert)
        Next

        ''Results - Probably placing under connections
        'For Each row As BoltResults In BoltResults
        '    SQLInsert = SQLInsert.Replace("--BEGIN --[BOLT RESULTS INSERT BEGIN]", "BEGIN --[BOLT RESULTS INSERT BEGIN]")
        '    SQLInsert = SQLInsert.Replace("--END --[BOLT RESULTS INSERT END]", "END --[BOLT RESULTS INSERT END]")
        '    SQLInsert = SQLInsert.Replace("--[BOLT RESULTS INSERT]", row.SQLInsert)
        'Next

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        'SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Bridge Stiffener Detail (UPDATE).sql")
        SQLUpdate = CCI_Engineering_Templates.My.Resources.CCIplate_Bridge_Stiffener_Detail_UPDATE
        SQLUpdate = SQLUpdate.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.Replace("[UPDATE]", Me.SQLUpdateFieldsandValues)
        SQLUpdate = SQLUpdate.TrimEnd() 'Removes empty rows that generate within query for each record

        'Plate Material
        For Each row As CCIplateMaterial In CCIplateMaterials
            SQLUpdate = SQLUpdate.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert) 'Can only insert materials, no deleting or updating since database is referenced by all BUs. 
        Next

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        'SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Bridge Stiffener Detail (DELETE).sql")
        SQLDelete = CCI_Engineering_Templates.My.Resources.CCIplate_Bridge_Stiffener_Detail_DELETE
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete

    End Function

#End Region

#Region "Define"
    Private _local_id As Integer?
    Private _local_connection_id As Integer?
    Private _connection_id As Integer?
    Private _stiffener_type As String
    Private _analysis_type As String
    Private _quantity As Double?
    Private _bridge_stiffener_width As Double?
    Private _bridge_stiffener_thickness As Double?
    Private _bridge_stiffener_material As Integer?
    Private _unbraced_length As Double?
    Private _total_length As Double?
    Private _weld_size As Double?
    Private _exx As Double?
    Private _upper_weld_length As Double?
    Private _lower_weld_length As Double?
    Private _upper_plate_width As Double?
    Private _lower_plate_width As Double?
    Private _neglect_flange_connection As Boolean?
    Private _bolt_hole_diameter As Double?
    Private _bolt_qty_eccentric As Double?
    Private _bolt_qty_shear As Double?
    Private _intermediate_bolt_spacing As Double?
    Private _bolt_diameter As Double?
    Private _bolt_sleeve_diameter As Double?
    Private _washer_diameter As Double?
    Private _bolt_tensile_strength As Double?
    Private _bolt_allowable_shear As Double?
    Private _exx_shim_plate As Double?
    Private _filler_shim_thickness As Double?

    <DataMember()> Public Property CCIplateMaterials As New List(Of CCIplateMaterial)

    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Local Id")>
    <DataMember()> Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Local Connection Id")>
    <DataMember()> Public Property local_connection_id() As Integer?
        Get
            Return Me._local_connection_id
        End Get
        Set
            Me._local_connection_id = Value
        End Set
    End Property

    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Connection Id")>
    <DataMember()> Public Property connection_id() As Integer?
        Get
            Return Me._connection_id
        End Get
        Set
            Me._connection_id = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Stiffener Type")>
    <DataMember()> Public Property stiffener_type() As String
        Get
            Return Me._stiffener_type
        End Get
        Set
            Me._stiffener_type = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Analysis Type")>
    <DataMember()> Public Property analysis_type() As String
        Get
            Return Me._analysis_type
        End Get
        Set
            Me._analysis_type = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Quantity")>
    <DataMember()> Public Property quantity() As Double?
        Get
            Return Me._quantity
        End Get
        Set
            Me._quantity = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bridge Stiffener Width")>
    <DataMember()> Public Property bridge_stiffener_width() As Double?
        Get
            Return Me._bridge_stiffener_width
        End Get
        Set
            Me._bridge_stiffener_width = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bridge Stiffener Thickness")>
    <DataMember()> Public Property bridge_stiffener_thickness() As Double?
        Get
            Return Me._bridge_stiffener_thickness
        End Get
        Set
            Me._bridge_stiffener_thickness = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bridge Stiffener Material")>
    <DataMember()> Public Property bridge_stiffener_material() As Integer?
        Get
            Return Me._bridge_stiffener_material
        End Get
        Set
            Me._bridge_stiffener_material = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Unbraced Length")>
    <DataMember()> Public Property unbraced_length() As Double?
        Get
            Return Me._unbraced_length
        End Get
        Set
            Me._unbraced_length = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Total Length")>
    <DataMember()> Public Property total_length() As Double?
        Get
            Return Me._total_length
        End Get
        Set
            Me._total_length = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Weld Size")>
    <DataMember()> Public Property weld_size() As Double?
        Get
            Return Me._weld_size
        End Get
        Set
            Me._weld_size = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Exx")>
    <DataMember()> Public Property exx() As Double?
        Get
            Return Me._exx
        End Get
        Set
            Me._exx = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Upper Weld Length")>
    <DataMember()> Public Property upper_weld_length() As Double?
        Get
            Return Me._upper_weld_length
        End Get
        Set
            Me._upper_weld_length = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Lower Weld Length")>
    <DataMember()> Public Property lower_weld_length() As Double?
        Get
            Return Me._lower_weld_length
        End Get
        Set
            Me._lower_weld_length = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Upper Plate Width")>
    <DataMember()> Public Property upper_plate_width() As Double?
        Get
            Return Me._upper_plate_width
        End Get
        Set
            Me._upper_plate_width = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Lower Plate Width")>
    <DataMember()> Public Property lower_plate_width() As Double?
        Get
            Return Me._lower_plate_width
        End Get
        Set
            Me._lower_plate_width = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Neglect Flange Connection")>
    <DataMember()> Public Property neglect_flange_connection() As Boolean?
        Get
            Return Me._neglect_flange_connection
        End Get
        Set
            Me._neglect_flange_connection = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Hole Diameter")>
    <DataMember()> Public Property bolt_hole_diameter() As Double?
        Get
            Return Me._bolt_hole_diameter
        End Get
        Set
            Me._bolt_hole_diameter = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Qty Eccentric")>
    <DataMember()> Public Property bolt_qty_eccentric() As Double?
        Get
            Return Me._bolt_qty_eccentric
        End Get
        Set
            Me._bolt_qty_eccentric = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Qty Shear")>
    <DataMember()> Public Property bolt_qty_shear() As Double?
        Get
            Return Me._bolt_qty_shear
        End Get
        Set
            Me._bolt_qty_shear = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Intermediate Bolt Spacing")>
    <DataMember()> Public Property intermediate_bolt_spacing() As Double?
        Get
            Return Me._intermediate_bolt_spacing
        End Get
        Set
            Me._intermediate_bolt_spacing = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Diameter")>
    <DataMember()> Public Property bolt_diameter() As Double?
        Get
            Return Me._bolt_diameter
        End Get
        Set
            Me._bolt_diameter = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Sleeve Diameter")>
    <DataMember()> Public Property bolt_sleeve_diameter() As Double?
        Get
            Return Me._bolt_sleeve_diameter
        End Get
        Set
            Me._bolt_sleeve_diameter = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Washer Diameter")>
    <DataMember()> Public Property washer_diameter() As Double?
        Get
            Return Me._washer_diameter
        End Get
        Set
            Me._washer_diameter = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Tensile Strength")>
    <DataMember()> Public Property bolt_tensile_strength() As Double?
        Get
            Return Me._bolt_tensile_strength
        End Get
        Set
            Me._bolt_tensile_strength = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Bolt Allowable Shear")>
    <DataMember()> Public Property bolt_allowable_shear() As Double?
        Get
            Return Me._bolt_allowable_shear
        End Get
        Set
            Me._bolt_allowable_shear = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Exx Shim Plate")>
    <DataMember()> Public Property exx_shim_plate() As Double?
        Get
            Return Me._exx_shim_plate
        End Get
        Set
            Me._exx_shim_plate = Value
        End Set
    End Property
    <Category("Bridge Stiffener Details"), Description(""), DisplayName("Filler Shim Thickness")>
    <DataMember()> Public Property filler_shim_thickness() As Double?
        Get
            Return Me._filler_shim_thickness
        End Get
        Set
            Me._filler_shim_thickness = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal bsdrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = bsdrow
        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_plate_id")))
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_id = DBtoNullableInt(dr.Item("local_bridge_stiffener_id")) 'currently not being referenced for anything
        End If
        Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id")) 'need to store local_connection_id within EDS since user may adjust relationship to any elevation and need to identify this as a change to perform update function. 
        Me.connection_id = If(EDStruefalse, DBtoNullableInt(dr.Item("plate_id")), DBtoNullableInt(dr.Item("connection_id")))
        'If EDStruefalse = True Then 'Only pull in when referencing EDS
        '    Me.plate_id = DBtoNullableInt(dr.Item("plate_id"))
        'End If
        Me.stiffener_type = DBtoStr(dr.Item("stiffener_type"))
        Me.analysis_type = DBtoStr(dr.Item("analysis_type"))
        Me.quantity = DBtoNullableDbl(dr.Item("quantity"))
        Me.bridge_stiffener_width = DBtoNullableDbl(dr.Item("bridge_stiffener_width"))
        Me.bridge_stiffener_thickness = DBtoNullableDbl(dr.Item("bridge_stiffener_thickness"))
        Me.bridge_stiffener_material = DBtoNullableInt(dr.Item("bridge_stiffener_material"))
        Me.unbraced_length = DBtoNullableDbl(dr.Item("unbraced_length"))
        Me.total_length = DBtoNullableDbl(dr.Item("total_length"))
        Me.weld_size = DBtoNullableDbl(dr.Item("weld_size"))
        Me.exx = DBtoNullableDbl(dr.Item("exx"))
        Me.upper_weld_length = DBtoNullableDbl(dr.Item("upper_weld_length"))
        Me.lower_weld_length = DBtoNullableDbl(dr.Item("lower_weld_length"))
        Me.upper_plate_width = DBtoNullableDbl(dr.Item("upper_plate_width"))
        Me.lower_plate_width = DBtoNullableDbl(dr.Item("lower_plate_width"))
        Me.neglect_flange_connection = If(EDStruefalse, DBtoNullableBool(dr.Item("neglect_flange_connection")), If(DBtoStr(dr.Item("neglect_flange_connection")) = "Yes", True, If(DBtoStr(dr.Item("neglect_flange_connection")) = "No", False, DBtoNullableBool(dr.Item("neglect_flange_connection")))))
        Me.bolt_hole_diameter = DBtoNullableDbl(dr.Item("bolt_hole_diameter"))
        Me.bolt_qty_eccentric = DBtoNullableDbl(dr.Item("bolt_qty_eccentric"))
        Me.bolt_qty_shear = DBtoNullableDbl(dr.Item("bolt_qty_shear"))
        Me.intermediate_bolt_spacing = DBtoNullableDbl(dr.Item("intermediate_bolt_spacing"))
        Me.bolt_diameter = DBtoNullableDbl(dr.Item("bolt_diameter"))
        Me.bolt_sleeve_diameter = DBtoNullableDbl(dr.Item("bolt_sleeve_diameter"))
        Me.washer_diameter = DBtoNullableDbl(dr.Item("washer_diameter"))
        Me.bolt_tensile_strength = DBtoNullableDbl(dr.Item("bolt_tensile_strength"))
        Me.bolt_allowable_shear = DBtoNullableDbl(dr.Item("bolt_allowable_shear"))
        Me.exx_shim_plate = DBtoNullableDbl(dr.Item("exx_shim_plate"))
        Me.filler_shim_thickness = DBtoNullableDbl(dr.Item("filler_shim_thickness"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.local_connection_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.stiffener_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.analysis_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bridge_stiffener_width.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bridge_stiffener_thickness.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bridge_stiffener_material.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel3ID")
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.unbraced_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.total_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.weld_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.exx.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.upper_weld_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.lower_weld_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.upper_plate_width.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.lower_plate_width.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.neglect_flange_connection.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_hole_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_qty_eccentric.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_qty_shear.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.intermediate_bolt_spacing.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_sleeve_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.washer_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_tensile_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_allowable_shear.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.exx_shim_plate.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.filler_shim_thickness.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""
        SQLInsertFields = SQLInsertFields.AddtoDBString("local_connection_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("stiffener_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("analysis_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bridge_stiffener_width")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bridge_stiffener_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bridge_stiffener_material")
        SQLInsertFields = SQLInsertFields.AddtoDBString("unbraced_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("total_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("weld_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("exx")
        SQLInsertFields = SQLInsertFields.AddtoDBString("upper_weld_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("lower_weld_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("upper_plate_width")
        SQLInsertFields = SQLInsertFields.AddtoDBString("lower_plate_width")
        SQLInsertFields = SQLInsertFields.AddtoDBString("neglect_flange_connection")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_hole_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_qty_eccentric")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_qty_shear")
        SQLInsertFields = SQLInsertFields.AddtoDBString("intermediate_bolt_spacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_sleeve_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("washer_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_tensile_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_allowable_shear")
        SQLInsertFields = SQLInsertFields.AddtoDBString("exx_shim_plate")
        SQLInsertFields = SQLInsertFields.AddtoDBString("filler_shim_thickness")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("local_connection_id = " & Me.local_connection_id.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_id = " & "@SubLevel1ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("stiffener_type = " & Me.stiffener_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("analysis_type = " & Me.analysis_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("quantity = " & Me.quantity.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bridge_stiffener_width = " & Me.bridge_stiffener_width.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bridge_stiffener_thickness = " & Me.bridge_stiffener_thickness.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bridge_stiffener_material = " & Me.bridge_stiffener_material.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bridge_stiffener_material = " & "@SubLevel3ID")
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("unbraced_length = " & Me.unbraced_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("total_length = " & Me.total_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("weld_size = " & Me.weld_size.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("exx = " & Me.exx.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("upper_weld_length = " & Me.upper_weld_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("lower_weld_length = " & Me.lower_weld_length.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("upper_plate_width = " & Me.upper_plate_width.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("lower_plate_width = " & Me.lower_plate_width.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("neglect_flange_connection = " & Me.neglect_flange_connection.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_hole_diameter = " & Me.bolt_hole_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_qty_eccentric = " & Me.bolt_qty_eccentric.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_qty_shear = " & Me.bolt_qty_shear.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("intermediate_bolt_spacing = " & Me.intermediate_bolt_spacing.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_diameter = " & Me.bolt_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_sleeve_diameter = " & Me.bolt_sleeve_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("washer_diameter = " & Me.washer_diameter.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_tensile_strength = " & Me.bolt_tensile_strength.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_allowable_shear = " & Me.bolt_allowable_shear.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("exx_shim_plate = " & Me.exx_shim_plate.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("filler_shim_thickness = " & Me.filler_shim_thickness.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'plate_material references the local id when coming from Excel. Need to convert to EDS ID when performing Equals function
        Dim material As Integer?
        For Each row As CCIplateMaterial In CCIplateMaterials
            If Me.bridge_stiffener_material = row.local_id And row.ID > 0 Then
                material = row.ID
                Exit For
            End If
        Next

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As BridgeStiffenerDetail = TryCast(other, BridgeStiffenerDetail)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.local_connection_id.CheckChange(otherToCompare.local_connection_id, changes, categoryName, "Plate Id"), Equals, False)
        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        'Equals = If(Me.plate_id.CheckChange(otherToCompare.plate_id, changes, categoryName, "Plate Id"), Equals, False)
        Equals = If(Me.stiffener_type.CheckChange(otherToCompare.stiffener_type, changes, categoryName, "Stiffener Type"), Equals, False)
        Equals = If(Me.analysis_type.CheckChange(otherToCompare.analysis_type, changes, categoryName, "Analysis Type"), Equals, False)
        Equals = If(Me.quantity.CheckChange(otherToCompare.quantity, changes, categoryName, "Quantity"), Equals, False)
        Equals = If(Me.bridge_stiffener_width.CheckChange(otherToCompare.bridge_stiffener_width, changes, categoryName, "Bridge Stiffener Width"), Equals, False)
        Equals = If(Me.bridge_stiffener_thickness.CheckChange(otherToCompare.bridge_stiffener_thickness, changes, categoryName, "Bridge Stiffener Thickness"), Equals, False)
        'Equals = If(Me.bridge_stiffener_material.CheckChange(otherToCompare.bridge_stiffener_material, changes, categoryName, "Bridge Stiffener Material"), Equals, False)
        Equals = If(material.CheckChange(otherToCompare.bridge_stiffener_material, changes, categoryName, "Bridge Stiffener Material"), Equals, False)
        Equals = If(Me.unbraced_length.CheckChange(otherToCompare.unbraced_length, changes, categoryName, "Unbraced Length"), Equals, False)
        Equals = If(Me.total_length.CheckChange(otherToCompare.total_length, changes, categoryName, "Total Length"), Equals, False)
        Equals = If(Me.weld_size.CheckChange(otherToCompare.weld_size, changes, categoryName, "Weld Size"), Equals, False)
        Equals = If(Me.exx.CheckChange(otherToCompare.exx, changes, categoryName, "Exx"), Equals, False)
        Equals = If(Me.upper_weld_length.CheckChange(otherToCompare.upper_weld_length, changes, categoryName, "Upper Weld Length"), Equals, False)
        Equals = If(Me.lower_weld_length.CheckChange(otherToCompare.lower_weld_length, changes, categoryName, "Lower Weld Length"), Equals, False)
        Equals = If(Me.upper_plate_width.CheckChange(otherToCompare.upper_plate_width, changes, categoryName, "Upper Plate Width"), Equals, False)
        Equals = If(Me.lower_plate_width.CheckChange(otherToCompare.lower_plate_width, changes, categoryName, "Lower Plate Width"), Equals, False)
        Equals = If(Me.neglect_flange_connection.CheckChange(otherToCompare.neglect_flange_connection, changes, categoryName, "Neglect Flange Connection"), Equals, False)
        Equals = If(Me.bolt_hole_diameter.CheckChange(otherToCompare.bolt_hole_diameter, changes, categoryName, "Bolt Hole Diameter"), Equals, False)
        Equals = If(Me.bolt_qty_eccentric.CheckChange(otherToCompare.bolt_qty_eccentric, changes, categoryName, "Bolt Qty Eccentric"), Equals, False)
        Equals = If(Me.bolt_qty_shear.CheckChange(otherToCompare.bolt_qty_shear, changes, categoryName, "Bolt Qty Shear"), Equals, False)
        Equals = If(Me.intermediate_bolt_spacing.CheckChange(otherToCompare.intermediate_bolt_spacing, changes, categoryName, "Intermediate Bolt Spacing"), Equals, False)
        Equals = If(Me.bolt_diameter.CheckChange(otherToCompare.bolt_diameter, changes, categoryName, "Bolt Diameter"), Equals, False)
        Equals = If(Me.bolt_sleeve_diameter.CheckChange(otherToCompare.bolt_sleeve_diameter, changes, categoryName, "Bolt Sleeve Diameter"), Equals, False)
        Equals = If(Me.washer_diameter.CheckChange(otherToCompare.washer_diameter, changes, categoryName, "Washer Diameter"), Equals, False)
        Equals = If(Me.bolt_tensile_strength.CheckChange(otherToCompare.bolt_tensile_strength, changes, categoryName, "Bolt Tensile Strength"), Equals, False)
        Equals = If(Me.bolt_allowable_shear.CheckChange(otherToCompare.bolt_allowable_shear, changes, categoryName, "Bolt Allowable Shear"), Equals, False)
        Equals = If(Me.exx_shim_plate.CheckChange(otherToCompare.exx_shim_plate, changes, categoryName, "Exx Shim Plate"), Equals, False)
        Equals = If(Me.filler_shim_thickness.CheckChange(otherToCompare.filler_shim_thickness, changes, categoryName, "Filler Shim Thickness"), Equals, False)

        'Materials
        If Me.CCIplateMaterials.Count > 0 Then
            Equals = If(Me.CCIplateMaterials.CheckChange(otherToCompare.CCIplateMaterials, changes, categoryName, "CCIplate Materials"), Equals, False)
        End If

    End Function
#End Region

End Class

<DataContractAttribute()>
Partial Public Class ConnectionResults
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Return "Connection Results"
        End Get
    End Property
    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Return "conn.connection_results"
        End Get
    End Property

    Public Overrides Function SQLInsert() As String

        'SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Connection Result (INSERT).sql")
        SQLInsert = CCI_Engineering_Templates.My.Resources.CCIplate_Connection_Result_INSERT
        SQLInsert = SQLInsert.Replace("[CONNECTION RESULT VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[CONNECTION RESULT FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLInsert

    End Function


#End Region

#Region "Define"
    Private _plate_id As Integer?
    Private _local_connection_id As Integer?
    'Private _local_bolt_group_id As Integer?
    'Private _work_order_seq_num As Double? 'not provided in Excel
    Private _rating As Decimal?
    Private _result_lkup As String
    'Private _modified_person_id As Integer? 'not provided in Excel
    'Private _process_stage As String 'not provided in Excel
    'Private _modified_date As DateTime? 'not provided in Excel

    <Category("Connection Results"), Description(""), DisplayName("Plate Id")>
    <DataMember()> Public Property plate_id() As Integer?
        Get
            Return Me._plate_id
        End Get
        Set
            Me._plate_id = Value
        End Set
    End Property
    <Category("Connection Results"), Description(""), DisplayName("Local Connection Id")>
    <DataMember()> Public Property local_connection_id() As Integer?
        Get
            Return Me._local_connection_id
        End Get
        Set
            Me._local_connection_id = Value
        End Set
    End Property
    '<Category("Connection Results"), Description(""), DisplayName("Local Bolt Group Id")>
    '<DataMember()> Public Property local_bolt_group_id() As Integer?
    '    Get
    '        Return Me._local_bolt_group_id
    '    End Get
    '    Set
    '        Me._local_bolt_group_id = Value
    '    End Set
    'End Property
    '<Category("Bolt Results"), Description(""), DisplayName("Work Order Seq Num")>
    '<DataMember()> Public Property work_order_seq_num() As Double?
    '    Get
    '        Return Me._work_order_seq_num
    '    End Get
    '    Set
    '        Me._work_order_seq_num = Value
    '    End Set
    'End Property
    <Category("Connection Results"), Description(""), DisplayName("Rating")>
    <DataMember()> Public Property rating() As Decimal?
        Get
            Return Me._rating
        End Get
        Set
            Me._rating = Value
        End Set
    End Property
    <Category("Connection Results"), Description(""), DisplayName("Result Lkup")>
    <DataMember()> Public Property result_lkup() As String
        Get
            Return Me._result_lkup
        End Get
        Set
            Me._result_lkup = Value
        End Set
    End Property
    '<Category("Bolt Results"), Description(""), DisplayName("Modified Person Id")>
    '<DataMember()> Public Property modified_person_id() As Integer?
    '    Get
    '        Return Me._modified_person_id
    '    End Get
    '    Set
    '        Me._modified_person_id = Value
    '    End Set
    'End Property
    '<Category("Bolt Results"), Description(""), DisplayName("Process Stage")>
    '<DataMember()> Public Property process_stage() As String
    '    Get
    '        Return Me._process_stage
    '    End Get
    '    Set
    '        Me._process_stage = Value
    '    End Set
    'End Property
    '<Category("Bolt Results"), Description(""), DisplayName("Modified Date")>
    '<DataMember()> Public Property modified_date() As DateTime?
    '    Get
    '        Return Me._modified_date
    '    End Get
    '    Set
    '        Me._modified_date = Value
    '    End Set
    'End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal crrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = crrow

        Me.plate_id = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_connection_id = DBtoNullableInt(dr.Item("local_connection_id"))
        End If
        'Me.work_order_seq_num = DBtoNullableDbl(dr.Item("work_order_seq_num"))
        Me.rating = DBtoNullableDec(dr.Item("rating")) 'same in all 
        Me.result_lkup = DBtoStr(dr.Item("result_lkup")) 'same in all
        'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
        'Me.process_stage = DBtoStr(dr.Item("process_stage"))
        'Me.modified_date = DBtoStr(dr.Item("modified_date"))

    End Sub

#End Region

#Region "Save to EDS"
    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString("@SubLevel1ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.plate_details_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.work_order_seq_num.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rating.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.result_lkup.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_date.ToString.FormatDBValue)


        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        SQLInsertFields = SQLInsertFields.AddtoDBString("plate_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("work_order_seq_num")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rating")
        SQLInsertFields = SQLInsertFields.AddtoDBString("result_lkup")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_date")


        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_id = " & Me.plate_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("work_order_seq_num = " & Me.work_order_seq_num.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rating = " & Me.rating.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("result_lkup = " & Me.result_lkup.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_date = " & Me.modified_date.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As ConnectionResults = TryCast(other, ConnectionResults)
        If otherToCompare Is Nothing Then Return False


    End Function
#End Region

End Class