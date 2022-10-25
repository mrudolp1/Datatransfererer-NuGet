Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
'Imports Microsoft.Office.Interop

Partial Public Class CCIplate
    Inherits EDSExcelObject


#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String = "CCIplates"
    Public Overrides ReadOnly Property EDSTableName As String = "conn.connections"
    Public Overrides ReadOnly Property templatePath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "CCIplate.xlsm")
    Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
        'Add additional sub table references here. Table names should be consistent with EDS table names. 
        Get
            Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("CCIplates", "A1:K2", "Details (SAPI)"),
                                                        New EXCELDTParameter("Connections", "C2:G18", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("Plate Details", "I2:S33", "Sub Tables (SAPI)"),
                                                        New EXCELDTParameter("CCIplate Materials", "AQ2:BC55", "Sub Tables (SAPI)")}

            'note: Excel table names are consistent with EDS table names to limit work required within constructors

        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    Private _Delete As String

    Public Overrides Function SQLInsert() As String

        If _Insert = "" Then
            _Insert = QueryBuilderFromFile(queryPath & "CCIplate\CCIplate (INSERT).sql")
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
            _Update = QueryBuilderFromFile(queryPath & "CCIplate\CCIplate (UPDATE).sql")
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
            _Delete = QueryBuilderFromFile(queryPath & "CCIplate\CCIplate (DELETE).sql")
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

    Public Property Connections As New List(Of Connection)

    '<Category("Connection"), Description(""), DisplayName("Id")>
    'Public Property ID() As Integer?
    '    Get
    '        Return Me._ID
    '    End Get
    '    Set
    '        Me._ID = Value
    '    End Set
    'End Property
    '<Category("Connection"), Description(""), DisplayName("Bus Unit")>
    'Public Property bus_unit() As String
    '    Get
    '        Return Me._bus_unit
    '    End Get
    '    Set
    '        Me._bus_unit = Value
    '    End Set
    'End Property
    '<Category("Connection"), Description(""), DisplayName("Structure Id")>
    'Public Property structure_id() As String
    '    Get
    '        Return Me._structure_id
    '    End Get
    '    Set
    '        Me._structure_id = Value
    '    End Set
    'End Property
    <Category("CCIplate"), Description(""), DisplayName("Anchor Rod Spacing")>
    Public Property anchor_rod_spacing() As Double?
        Get
            Return Me._anchor_rod_spacing
        End Get
        Set
            Me._anchor_rod_spacing = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Clip Distance")>
    Public Property clip_distance() As Double?
        Get
            Return Me._clip_distance
        End Get
        Set
            Me._clip_distance = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Barb Cl Elevation")>
    Public Property barb_cl_elevation() As Double?
        Get
            Return Me._barb_cl_elevation
        End Get
        Set
            Me._barb_cl_elevation = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Include Pole Reactions")>
    Public Property include_pole_reactions() As Boolean?
        Get
            Return Me._include_pole_reactions
        End Get
        Set
            Me._include_pole_reactions = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Consider Ar Eccentricity")>
    Public Property consider_ar_eccentricity() As Boolean?
        Get
            Return Me._consider_ar_eccentricity
        End Get
        Set
            Me._consider_ar_eccentricity = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Leg Mod Eccentricity")>
    Public Property leg_mod_eccentricity() As Double?
        Get
            Return Me._leg_mod_eccentricity
        End Get
        Set
            Me._leg_mod_eccentricity = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Seismic")>
    Public Property seismic() As Boolean?
        Get
            Return Me._seismic
        End Get
        Set
            Me._seismic = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Seismic Flanges")>
    Public Property seismic_flanges() As Boolean?
        Get
            Return Me._seismic_flanges
        End Get
        Set
            Me._seismic_flanges = Value
        End Set
    End Property
    <Category("CCIplate"), Description(""), DisplayName("Structural 105")>
    Public Property Structural_105() As Boolean?
        Get
            Return Me._Structural_105
        End Get
        Set
            Me._Structural_105 = Value
        End Set
    End Property
    '<Category("Connection"), Description(""), DisplayName("Tool Version")>
    'Public Property tool_version() As String
    '    Get
    '        Return Me._tool_version
    '    End Get
    '    Set
    '        Me._tool_version = Value
    '    End Set
    'End Property
    '<Category("Connection"), Description(""), DisplayName("Modified Person Id")>
    'Public Property modified_person_id() As Integer?
    '    Get
    '        Return Me._modified_person_id
    '    End Get
    '    Set
    '        Me._modified_person_id = Value
    '    End Set
    'End Property
    '<Category("Connection"), Description(""), DisplayName("Process Stage")>
    'Public Property process_stage() As String
    '    Get
    '        Return Me._process_stage
    '    End Get
    '    Set
    '        Me._process_stage = Value
    '    End Set
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal dr As DataRow, ByRef strDS As DataSet, Optional ByRef Parent As EDSObject = Nothing) 'Added strDS in order to pull EDS data from subtables
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        'Following used to create dataset, regardless if source was EDS or Excel. Boolean used to identify source. EDS = True
        BuildFromDataset(dr, strDS, True, Me)

    End Sub 'Generate a CCIplate from EDS

    Public Sub New(ExcelFilePath As String, Optional ByRef Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet

        For Each item As EXCELDTParameter In excelDTParams
            'Get additional tables from excel file 
            Try
                excelDS.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
            Catch ex As Exception
                Debug.Print(String.Format("Failed to create datatable for: {0}, {1}, {2}", IO.Path.GetFileName(ExcelFilePath), item.xlsSheet, item.xlsRange))
            End Try
        Next

        If excelDS.Tables.Contains("CCIplates") Then
            Dim dr = excelDS.Tables("CCIplates").Rows(0)

            'Following used to create dataset, regardless if source was EDS or Excel. Boolean used to identify source. Excel = False
            BuildFromDataset(dr, excelDS, False, Me)

        End If


        'If excelDS.Tables.Contains("Pile Results EXCEL") Then

        '    For Each Row As DataRow In excelDS.Tables("Pile Results EXCEL").Rows

        '        'For Tools with multiple foundation or sub items, use Row.Item("ID") or add a local_ID column to filter which results should be associated with each foundation

        '        Me.Results.Add(New EDSResult(Row, Me))

        '    Next

        'End If

    End Sub 'Generate a CCIplate from Excel

    Private Sub BuildFromDataset(ByVal dr As DataRow, ByRef ds As DataSet, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing)
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
        Me.tool_version = DBtoStr(dr.Item("tool_version"))
        Me.modified_person_id = If(EDStruefalse, DBtoNullableInt(dr.Item("modified_person_id")), Me.modified_person_id) 'Not provided in Excel
        Me.process_stage = If(EDStruefalse, DBtoStr(dr.Item("process_stage")), Me.process_stage) 'Not provided in Excel

        Dim plConnection As New Connection 'Connection
        Dim plPlateDetail As New PlateDetail ' Plate
        Dim plCCIplateMaterial As New CCIplateMaterial

        For Each crow As DataRow In ds.Tables(plConnection.EDSObjectName).Rows
            'create a new connection based on the datarow from above
            plConnection = New Connection(crow, EDStruefalse, Me)
            'Check if the parent id, in the case cciplate id is equal to the original object id (Me)                    
            If If(EDStruefalse, plConnection.connection_id = Me.ID, True) Then 'If coming from Excel, all connections provided will be associated to CCIplate. 
                'If it is equal then add the newly created connection to the list of connections 
                Connections.Add(plConnection)
                'Loop through all plates pulled from EDS and check if they are associated with the newly created connection
                For Each pdrow As DataRow In ds.Tables(plPlateDetail.EDSObjectName).Rows
                    'Create a new plate from the plate datarow from EDS
                    plPlateDetail = New PlateDetail(pdrow, EDStruefalse, Me)
                    If If(EDStruefalse, plPlateDetail.plate_id = plConnection.ID, plPlateDetail.plate_id = plConnection.local_id) Then
                        plConnection.PlateDetails.Add(plPlateDetail)
                        For Each mrow As DataRow In ds.Tables(plCCIplateMaterial.EDSObjectName).Rows
                            plCCIplateMaterial = New CCIplateMaterial(mrow, EDStruefalse, Me)
                            If If(EDStruefalse, plCCIplateMaterial.ID = plPlateDetail.plate_material, plCCIplateMaterial.local_id = plPlateDetail.plate_material) Then
                                'plPlateDetail.plate_material = plConnectionMaterial
                                plPlateDetail.CCIplateMaterials.Add(plCCIplateMaterial)
                                Exit For 'Once matched, don't need to continue checking. 
                            End If
                        Next
                        'For Each drrr As DataRow In ds.Tables(plBolt.EDSTableName).Rows)
                        '            plBolt = New Bolt(drrr, plPlate)
                        '    If plBolt.plate_id = plPlate.ID Then
                        '        plPlate.bolts.Add(plBolt)
                        '    End If
                        'Next
                    End If
                Next
            End If
        Next

        'End Function
    End Sub

#End Region

#Region "Save to Excel"

    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        '''''Customize for each excel tool'''''

        'Site Code Criteria
        Dim tia_current, site_name, structure_type As String
        Dim rev_h_section_15_5 As Boolean?

        With wb
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
            If Not IsNothing(Me.barb_cl_elevation) Then
                .Worksheets("Custom Connection").Range("H8").Value = CType(Me.barb_cl_elevation, Double)
            Else
                .Worksheets("Custom Connection").Range("H8").ClearContents
            End If
            If Not IsNothing(Me.include_pole_reactions) Then
                .Worksheets("BARB").Range("X2").Value = CType(Me.include_pole_reactions, Boolean)
            End If
            If Not IsNothing(Me.consider_ar_eccentricity) Then
                .Worksheets("Engine").Range("D17").Value = CType(Me.consider_ar_eccentricity, Boolean)
            End If
            If Not IsNothing(Me.leg_mod_eccentricity) Then
                .Worksheets("Custom Connection").Range("J8").Value = CType(Me.leg_mod_eccentricity, Double)
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

            'Site Code Criteria
            'Site Name
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.site_name) Then
                site_name = Me.ParentStructure?.structureCodeCriteria?.site_name
                .Worksheets("Main").Range("C4").Value = CType(site_name, String)
            End If
            'Order Number
            'If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.order_number) Then
            '    site_name = Me.ParentStructure?.structureCodeCriteria?.order_number
            '    .Worksheets("Input").Range("D7").Value = CType(order_number, String)
            'End If
            'Tower Type - Defaulting to Monopole if not one of the main tower types
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.structure_type) Then
                If Me.ParentStructure?.structureCodeCriteria?.structure_type = "SELF SUPPORT" Then
                    structure_type = "Self Suppot"
                ElseIf Me.ParentStructure?.structureCodeCriteria?.structure_type = "MONOPOLE" Then
                    structure_type = "Monopole"
                Else
                    structure_type = "Monopole"
                End If
                .Worksheets("Main").Range("tower_type").Value = CType(structure_type, String)
            End If
            'TIA Revision- Defaulting to Rev. H if not available. 
            If Not IsNothing(Me.ParentStructure?.structureCodeCriteria?.tia_current) Then
                If Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-F" Then
                    tia_current = "F"
                ElseIf Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-G" Then
                    tia_current = "G"
                ElseIf Me.ParentStructure?.structureCodeCriteria?.tia_current = "TIA-222-H" Then
                    tia_current = "H"
                Else
                    tia_current = "H"
                End If
                .Worksheets("Main").Range("C9").Value = CType(tia_current, String)
            End If
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


            If Me.Connections.Count > 0 Then
                'identify first row to copy data into Excel Sheet
                Dim PlateRow As Integer = 3 'SAPI Tab
                Dim PlateRow2 As Integer = 46 'MP Connection Summary Tab
                Dim PlateRow3 As Integer = 19 'Main Tab (currently not used)
                Dim PlateDRow As Integer = 3 'SAPI Tab
                Dim PlateDRow2 As Integer = 46 'MP Connection Summary Tab
                Dim MatRow As Integer = 40 'Materials Tab
                Dim MatRow2 As Integer = 38 'SAPI Tab
                Dim i As Integer
                Dim tempMaterials As New List(Of CCIplateMaterial)
                Dim tempMaterial As New CCIplateMaterial
                Dim matflag As Boolean = False


                For Each row As Connection In Connections

                    If Not IsNothing(row.ID) Then
                        .Worksheets("Sub Tables (SAPI)").Range("D" & PlateRow).Value = CType(row.ID, Integer)
                    End If
                    If Not IsNothing(row.connection_elevation) Then
                        .Worksheets("MP Connection Summary").Range("C" & PlateRow2).Value = CType(row.connection_elevation, Double)
                    Else
                        .Worksheets("MP Connection Summary").Range("C" & PlateRow2).ClearContents
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

                    If Not IsNothing(row.bolt_configuration) Then
                        .Worksheets("MP Connection Summary").Range("K" & PlateRow2).Value = CType(row.bolt_configuration, String)
                    End If

                    PlateRow += 1
                    PlateRow2 -= 2
                    'PlateRow3 -= 1

                    For Each pdrow As PlateDetail In row.PlateDetails
                        If pdrow.plate_id = row.ID Then

                            If pdrow.plate_location = "Bottom" Then
                                i = 1
                            Else
                                i = 0
                            End If

                            If Not IsNothing(pdrow.ID) Then
                                .Worksheets("Sub Tables (SAPI)").Range("K" & PlateDRow).Value = CType(pdrow.ID, Integer)
                            End If
                            'If Not IsNothing(row.plate_location) Then 'do not need, tool will autopopulate
                            '    .Worksheets("MP Connection Summary").Range("D" & PlateRow2).Value = CType(row.plate_location, String)
                            'End If
                            If Not IsNothing(pdrow.plate_type) Then
                                .Worksheets("MP Connection Summary").Range("E" & PlateDRow2 + i).Value = CType(pdrow.plate_type, String)
                                'need to add in a if/then statement to adjust row dependent on whether plate location is top or bottom. 
                            End If
                            If Not IsNothing(pdrow.plate_diameter) Then
                                .Worksheets("MP Connection Summary").Range("F" & PlateDRow2 + i).Value = CType(pdrow.plate_diameter, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("F" & PlateDRow2 + i).ClearContents
                            End If
                            If Not IsNothing(pdrow.plate_thickness) Then
                                .Worksheets("MP Connection Summary").Range("G" & PlateDRow2 + i).Value = CType(pdrow.plate_thickness, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("G" & PlateDRow2 + i).ClearContents
                            End If
                            'If Not IsNothing(pdrow.plate_material) Then
                            '    .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).Value = CType(pdrow.plate_material, Integer)
                            'Else
                            '    .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).ClearContents
                            'End If
                            For Each mrow As CCIplateMaterial In pdrow.CCIplateMaterials
                                If mrow.default_material = True Then
                                    If Not IsNothing(mrow.name) Then
                                        .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).Value = CType(mrow.name, String)
                                    Else
                                        .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).ClearContents
                                    End If
                                Else 'After adding new materail, save material name in a list to reference for other plates to see if materail was already added. 
                                    For Each tmrow In tempMaterials
                                        If mrow.ID = tmrow.ID Then
                                            matflag = True 'don't add to excel
                                            Exit For
                                        End If
                                    Next
                                    If matflag = False Then
                                        tempMaterial = New CCIplateMaterial(mrow.ID)
                                        tempMaterials.Add(tempMaterial)

                                        ''.Worksheets("Sub Tables (SAPI)").Range("AR").Count
                                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Columns("AR").Count 'counts total columns in excel
                                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").Cells("AR:AR").Count
                                        'Dim testrow As Integer = .Worksheets("Sub Tables (SAPI)").GetDataRange("AR:AR")
                                        If Not IsNothing(mrow.ID) Then
                                            .Worksheets("Sub Tables (SAPI)").Range("AR" & MatRow2).Value = CType(mrow.ID, Integer)
                                        End If
                                        If Not IsNothing(mrow.name) Then
                                            .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).Value = CType(mrow.name, String)
                                        Else
                                            .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).ClearContents
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
                                        MatRow += 1
                                        MatRow2 += 1
                                    Else
                                        If Not IsNothing(mrow.name) Then
                                            .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).Value = CType(mrow.name, String)
                                        Else
                                            .Worksheets("MP Connection Summary").Range("H" & PlateDRow2 + i).ClearContents
                                        End If
                                    End If
                                End If
                                matflag = False 'reset flag
                            Next

                            If Not IsNothing(pdrow.stiffener_configuration) Then
                                .Worksheets("MP Connection Summary").Range("I" & PlateDRow2 + i).Value = CType(pdrow.stiffener_configuration, Integer)
                            Else
                                .Worksheets("MP Connection Summary").Range("I" & PlateDRow2 + i).ClearContents
                            End If
                            If Not IsNothing(pdrow.stiffener_clear_space) Then
                                .Worksheets("MP Connection Summary").Range("D" & PlateDRow2 + i + 39).Value = CType(pdrow.stiffener_clear_space, Double)
                            Else
                                .Worksheets("MP Connection Summary").Range("D" & PlateDRow2 + i + 39).ClearContents
                            End If
                            If Not IsNothing(pdrow.plate_check) Then
                                If pdrow.plate_check = True Then
                                    .Worksheets("MP Connection Summary").Range("J" & PlateDRow2 + i).Value = "Yes"
                                Else
                                    .Worksheets("MP Connection Summary").Range("J" & PlateDRow2 + i).Value = "No"
                                End If
                            End If

                            PlateDRow += 1

                        End If
                    Next


                    PlateDRow2 -= 2

                Next
            End If


            'Worksheet Change Events
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
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tool_version.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

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
        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

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
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tool_version = " & Me.tool_version.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        'SQLUpdate = SQLUpdate.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)

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
        Equals = If(Me.tool_version.CheckChange(otherToCompare.tool_version, changes, categoryName, "Tool Version"), Equals, False)
        'Equals = If(Me.modified_person_id.CheckChange(otherToCompare.modified_person_id, changes, categoryName, "Modified Person Id"), Equals, False)
        'Equals = If(Me.process_stage.CheckChange(otherToCompare.process_stage, changes, categoryName, "Process Stage"), Equals, False)

        'Connection
        If Me.Connections.Count > 0 Then
            Equals = If(Me.Connections.CheckChange(otherToCompare.Connections, changes, categoryName, "Plates"), Equals, False)
        End If

        Return Equals

    End Function
#End Region

End Class

Partial Public Class Connection
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "Connections"
    Public Overrides ReadOnly Property EDSTableName As String = "conn.plates"

    Public Overrides Function SQLInsert() As String

        SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Connection (INSERT).sql")
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

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Connection (UPDATE).sql")
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

        Return SQLUpdate

    End Function

    Public Overrides Function SQLDelete() As String

        SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Connection (DELETE).sql")
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

        Return SQLDelete

    End Function

#End Region

#Region "Define"
    Private _local_id As Integer?
    Private _ID As Integer?
    Private _connection_elevation As Double?
    Private _connection_type As String
    Private _bolt_configuration As String
    Private _connection_id As Integer?

    Public Property PlateDetails As New List(Of PlateDetail)

    <Category("Connection"), Description(""), DisplayName("Local Id")>
    Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property

    <Category("Connection"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Connection"), Description(""), DisplayName("Connection Elevation")>
    Public Property connection_elevation() As Double?
        Get
            Return Me._connection_elevation
        End Get
        Set
            Me._connection_elevation = Value
        End Set
    End Property
    <Category("Connection"), Description(""), DisplayName("Connection Type")>
    Public Property connection_type() As String
        Get
            Return Me._connection_type
        End Get
        Set
            Me._connection_type = Value
        End Set
    End Property
    <Category("Connection"), Description(""), DisplayName("Bolt Configuration")>
    Public Property bolt_configuration() As String
        Get
            Return Me._bolt_configuration
        End Get
        Set
            Me._bolt_configuration = Value
        End Set
    End Property
    <Category("Connection"), Description(""), DisplayName("Connection Id")>
    Public Property connection_id() As Integer?
        Get
            Return Me._connection_id
        End Get
        Set
            Me._connection_id = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal row As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal prow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = row
        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_id"))) '-Must only associate this to EDS since 0 vs. >0 triggers different functions (e.g. update vs. delete)
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_id = DBtoNullableInt(dr.Item("local_id"))
        End If
        Me.connection_elevation = DBtoNullableDbl(dr.Item("connection_elevation"))
        Me.connection_type = DBtoStr(dr.Item("connection_type"))
        Me.bolt_configuration = DBtoStr(dr.Item("bolt_configuration"))
        Me.connection_id = If(EDStruefalse, DBtoNullableInt(dr.Item("connection_id")), Me.connection_id) 'Not provided in Excel

    End Sub

#End Region

#Region "Save to Excel"

    'Public Sub New(ByVal row As DataRow)

    'End Sub
    'Public Sub SaveExcel(ByRef wb As Workbook)

    '    Dim PlateRow As Integer = 3 'identify first row to copy data into Excel Sheet
    '    Dim PlateRow2 As Integer = 46 'identify first row to copy data into Excel Sheet
    '    Dim PlateRow3 As Integer = 19 'identify first row to copy data into Excel Sheet


    '    With wb
    '        If Not IsNothing(Me.ID) Then
    '            .Worksheets("Sub Tables (SAPI)").Range("D" & PlateRow).Value = CType(Me.ID, Integer)
    '        End If
    '        If Not IsNothing(Me.connection_elevation) Then
    '            .Worksheets("MP Connection Summary").Range("C" & PlateRow2).Value = CType(Me.connection_elevation, Double)
    '        Else
    '            .Worksheets("MP Connection Summary").Range("C" & PlateRow2).ClearContents
    '        End If
    '    End With

    'End Sub

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
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_elevation = " & Me.connection_elevation.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_type = " & Me.connection_type.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_configuration = " & Me.bolt_configuration.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("connection_id = " & Me.connection_id.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

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
        'Equals = If(Me.connection_id.CheckChange(otherToCompare.connection_id, changes, categoryName, "Connection Id"), Equals, False)

        'Plate Details
        If Me.PlateDetails.Count > 0 Then
            Equals = If(Me.PlateDetails.CheckChange(otherToCompare.PlateDetails, changes, categoryName, "Plate Details"), Equals, False)
        End If

    End Function

End Class

Partial Public Class PlateDetail
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "Plate Details"
    Public Overrides ReadOnly Property EDSTableName As String = "conn.plate_details"

    Public Overrides Function SQLInsert() As String

        SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\Plate Detail (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[PLATE DETAIL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[PLATE DETAIL FIELDS]", Me.SQLInsertFields)
        SQLInsert = SQLInsert.TrimEnd() 'Removes empty rows that generate within query for each record

        'Plate Material
        For Each row As CCIplateMaterial In CCIplateMaterials
            SQLInsert = SQLInsert.Replace("--[CCIPLATE MATERIAL INSERT]", row.SQLInsert)
        Next

        ''Plate Material
        'For Each row As ConnectionMaterial In ConnectionMaterials
        '    If row.ID = Me.plate_material Then
        '        SQLInsert = SQLInsert.Replace("--[PLATE MATERIAL INSERT]", row.SQLInsert)
        '    End If
        '    If Me.plate_material > 0 Then

        '    End If
        'Next

        Return SQLInsert

    End Function

    Public Overrides Function SQLUpdate() As String

        SQLUpdate = QueryBuilderFromFile(queryPath & "CCIplate\Plate Detail (UPDATE).sql")
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

        SQLDelete = QueryBuilderFromFile(queryPath & "CCIplate\Plate Detail (DELETE).sql")
        SQLDelete = SQLDelete.Replace("[ID]", Me.ID.ToString.FormatDBValue)
        SQLDelete = SQLDelete.TrimEnd() 'Removes empty rows that generate within query for each record

        Return SQLDelete

    End Function

#End Region

#Region "Define"
    Private _local_id As Integer?
    Private _ID As Integer?
    Private _plate_id As Integer?
    Private _plate_location As String
    Private _plate_type As String
    Private _plate_diameter As Double?
    Private _plate_thickness As Double?
    Private _plate_material As Integer?
    Private _stiffener_configuration As Integer?
    Private _stiffener_clear_space As Double?
    Private _plate_check As Boolean?

    Public Property CCIplateMaterials As New List(Of CCIplateMaterial)

    <Category("Plate Details"), Description(""), DisplayName("Local Id")>
    Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property

    <Category("Plate Details"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Id")>
    Public Property plate_id() As Integer?
        Get
            Return Me._plate_id
        End Get
        Set
            Me._plate_id = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Location")>
    Public Property plate_location() As String
        Get
            Return Me._plate_location
        End Get
        Set
            Me._plate_location = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Type")>
    Public Property plate_type() As String
        Get
            Return Me._plate_type
        End Get
        Set
            Me._plate_type = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Diameter")>
    Public Property plate_diameter() As Double?
        Get
            Return Me._plate_diameter
        End Get
        Set
            Me._plate_diameter = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Thickness")>
    Public Property plate_thickness() As Double?
        Get
            Return Me._plate_thickness
        End Get
        Set
            Me._plate_thickness = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Material")>
    Public Property plate_material() As Integer?
        Get
            Return Me._plate_material
        End Get
        Set
            Me._plate_material = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Stiffener Configuration")>
    Public Property stiffener_configuration() As Integer?
        Get
            Return Me._stiffener_configuration
        End Get
        Set
            Me._stiffener_configuration = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Stiffener Clear Space")>
    Public Property stiffener_clear_space() As Double?
        Get
            Return Me._stiffener_clear_space
        End Get
        Set
            Me._stiffener_clear_space = Value
        End Set
    End Property
    <Category("Plate Details"), Description(""), DisplayName("Plate Check")>
    Public Property plate_check() As Boolean?
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

    Public Sub New(ByVal pdrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal pdrow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = pdrow
        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_plate_id")))
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_id = DBtoNullableInt(dr.Item("local_plate_id"))
        End If
        Me.plate_id = If(EDStruefalse, DBtoNullableInt(dr.Item("plate_id")), DBtoNullableInt(dr.Item("local_id")))
        Me.plate_location = DBtoStr(dr.Item("plate_location"))
        Me.plate_type = DBtoStr(dr.Item("plate_type"))
        Me.plate_diameter = DBtoNullableDbl(dr.Item("plate_diameter"))
        Me.plate_thickness = DBtoNullableDbl(dr.Item("plate_thickness"))
        Me.plate_material = DBtoNullableInt(dr.Item("plate_material"))
        Me.stiffener_configuration = DBtoNullableInt(dr.Item("stiffener_configuration"))
        Me.stiffener_clear_space = DBtoNullableDbl(dr.Item("stiffener_clear_space"))
        'Me.plate_check = DBtoNullableBool(dr.Item("plate_check"))
        Me.plate_check = If(DBtoStr(dr.Item("plate_check")) = "Yes" Or DBtoStr(dr.Item("plate_check")) = "True", True, False) 'Listed as a string and need to convert to Boolean

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
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("plate_id = " & Me.plate_id.ToString.FormatDBValue)
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

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

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
        Equals = If(Me.plate_material.CheckChange(otherToCompare.plate_material, changes, categoryName, "Plate Material"), Equals, False)
        Equals = If(Me.stiffener_configuration.CheckChange(otherToCompare.stiffener_configuration, changes, categoryName, "Stiffener Configuration"), Equals, False)
        Equals = If(Me.stiffener_clear_space.CheckChange(otherToCompare.stiffener_clear_space, changes, categoryName, "Stiffener Clear Space"), Equals, False)
        Equals = If(Me.plate_check.CheckChange(otherToCompare.plate_check, changes, categoryName, "Plate Check"), Equals, False)

        'Materials
        If Me.CCIplateMaterials.Count > 0 Then
            Equals = If(Me.CCIplateMaterials.CheckChange(otherToCompare.CCIplateMaterials, changes, categoryName, "CCIplate Materials"), Equals, False)
        End If

    End Function

End Class
Partial Public Class CCIplateMaterial
    Inherits EDSObjectWithQueries

#Region "Inheritted"
    Public Overrides ReadOnly Property EDSObjectName As String = "CCIplate Materials"
    Public Overrides ReadOnly Property EDSTableName As String = "gen.connection_material_properties"

    Public Overrides Function SQLInsert() As String

        SQLInsert = QueryBuilderFromFile(queryPath & "CCIplate\CCIplate Material (INSERT).sql")
        SQLInsert = SQLInsert.Replace("[MATERIAL PROPERTY ID]", Me.ID.ToString.FormatDBValue)
        SQLInsert = SQLInsert.Replace("[SELECT]", Me.SQLUpdateFieldsandValues)
        SQLInsert = SQLInsert.Replace("[CCIPLATE MATERIAL VALUES]", Me.SQLInsertValues)
        SQLInsert = SQLInsert.Replace("[CCIPLATE MATERIAL FIELDS]", Me.SQLInsertFields)
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
    Private _ID As Integer?
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
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer?
        Get
            Return Me._ID
        End Get
        Set
            Me._ID = Value
        End Set
    End Property
    <Category("Connection Material Properties"), Description(""), DisplayName("Local Id")>
    Public Property local_id() As Integer?
        Get
            Return Me._local_id
        End Get
        Set
            Me._local_id = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Name")>
    Public Property name() As String
        Get
            Return Me._name
        End Get
        Set
            Me._name = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 0")>
    Public Property fy_0() As Double?
        Get
            Return Me._fy_0
        End Get
        Set
            Me._fy_0 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 1 125")>
    Public Property fy_1_125() As Double?
        Get
            Return Me._fy_1_125
        End Get
        Set
            Me._fy_1_125 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 1 625")>
    Public Property fy_1_625() As Double?
        Get
            Return Me._fy_1_625
        End Get
        Set
            Me._fy_1_625 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 2 625")>
    Public Property fy_2_625() As Double?
        Get
            Return Me._fy_2_625
        End Get
        Set
            Me._fy_2_625 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fy 4 125")>
    Public Property fy_4_125() As Double?
        Get
            Return Me._fy_4_125
        End Get
        Set
            Me._fy_4_125 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 0")>
    Public Property fu_0() As Double?
        Get
            Return Me._fu_0
        End Get
        Set
            Me._fu_0 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 1 125")>
    Public Property fu_1_125() As Double?
        Get
            Return Me._fu_1_125
        End Get
        Set
            Me._fu_1_125 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 1 625")>
    Public Property fu_1_625() As Double?
        Get
            Return Me._fu_1_625
        End Get
        Set
            Me._fu_1_625 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 2 625")>
    Public Property fu_2_625() As Double?
        Get
            Return Me._fu_2_625
        End Get
        Set
            Me._fu_2_625 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Fu 4 125")>
    Public Property fu_4_125() As Double?
        Get
            Return Me._fu_4_125
        End Get
        Set
            Me._fu_4_125 = Value
        End Set
    End Property
    <Category("CCIplate Material Properties"), Description(""), DisplayName("Default Material")>
    Public Property default_material() As Boolean?
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

    Public Sub New(ByVal mrow As DataRow, ByVal EDStruefalse As Boolean, Optional ByRef Parent As EDSObject = Nothing) '(ByVal mrow As DataRow, ByRef strDS As DataSet)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        Dim dr = mrow
        'Me.ID = If(EDStruefalse, DBtoNullableInt(dr.Item("ID")), DBtoNullableInt(dr.Item("local_id")))
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        If EDStruefalse = False Then 'Only pull in local id when referencing Excel
            Me.local_id = DBtoNullableInt(dr.Item("local_id"))
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

    Public Sub New(ByVal ID As Integer?)
        'This is used to store a temp list of new materials to add to the Excel tool
        Me.ID = ID
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
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("fy_0 = " & Me.fy_0.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("fy_1_125 = " & Me.fy_1_125.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("fy_1_625 = " & Me.fy_1_625.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("fy_2_625 = " & Me.fy_2_625.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("fy_4_125 = " & Me.fy_4_125.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("fu_0 = " & Me.fu_0.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("fu_1_125 = " & Me.fu_1_125.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("fu_1_625 = " & Me.fu_1_625.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("fu_2_625 = " & Me.fu_2_625.ToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("fu_4_125 = " & Me.fu_4_125.ToString.FormatDBValue)
        'SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.SelectDBString("default_material = " & Me.default_material.ToString.FormatDBValue)

        Return SQLUpdateFieldsandValues
    End Function
#End Region

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

End Class