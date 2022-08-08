Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
Imports Microsoft.Office.Interop

Partial Public Class Pile
    Inherits EDSExcelObject
    'Inherits EDSObjectWithQueries

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String = "Pile"
    Public Overrides ReadOnly Property EDSTableName As String = "fnd.pile"
    Public Overrides ReadOnly Property templatePath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Pile Foundation.xlsm")
    Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
        Get
            Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("Pile General Details EXCEL", "A1:BG2", "Details (SAPI)"),
                                                        New EXCELDTParameter("Pile Soil Profile EXCEL", "BI1:BJ2", "Details (SAPI)"),
                                                        New EXCELDTParameter("Pile Soil Layer EXCEL", "A3:I17", "SAPI")}
            '***Add additional table references here****
        End Get
    End Property
    Private _Insert As String
    Private _Update As String
    Private _Delete As String
    Public Overrides ReadOnly Property Insert() As String
        Get
            If _Insert = "" Then
                _Insert = QueryBuilderFromFile(queryPath & "Pile\Pile (INSERT).sql")
            End If
            Insert = _Insert
            'InsertString = InsertString.Replace("[BU NUMBER]", Me.bus_unit.FormatDBValue)
            'InsertString = InsertString.Replace("[STRUCTURE ID]", Me.structure_id.FormatDBValue)
            Insert = Insert.Replace("[FOUNDATION VALUES]", Me.SQLInsertValues)
            Insert = Insert.Replace("[FOUNDATION FIELDS]", Me.SQLInsertFields)
            'currentSortedList.Insert(i, Nothing)
            'InsertString = InsertString.Replace("[SOIL PROFILE]", Me.SQLInsertValues)
            'InsertString = InsertString.Replace("[SOIL PROFILE VALUES]", Me.SQLInsertValues)
            'InsertString = InsertString.Replace("[SOIL PROFILE FIELDS]", Me.SQLInsertFields)
            'InsertString = InsertString.Replace("[SOIL LAYER VALUES]", Me.SQLInsertValues)
            'InsertString = InsertString.Replace("[SOIL LAYER FIELDS]", Me.SQLInsertFields)
            'InsertString = InsertString.Replace("[RESULTS]", Me.Results.EDSResultQuery(False))
            '***Add additional table references here****

            'If Me.pile_soil_capacity_given = False And Me.pile_shape <> "H-pile" Then

            'End If

            'Dim SoilProfileQuery As String = QueryBuilderFromFile(queryPath & "Soil Profile\Soil Profile (INSERT).sql")
            'SoilProfileQuery = SoilProfileQuery.Replace("[SOIL PROFILE VALUES]", Me.SQLInsertValues)
            'SoilProfileQuery = SoilProfileQuery.Replace("[SOIL PROFILE FIELDS]", Me.SQLInsertValues)

            Insert = Insert.Replace("--[SOIL PROFILE INSERT]", SoilProfile.Insert)

            'If Me.pile_soil_capacity_given = False And Me.pile_shape <> "H-Pile" Then
            '    For Each pfsl As DataRow In strDS.Tables("Pile").Rows
            '        'For Each pfsl As PileSoilLayer In pf.soil_layers
            '        'line added below to avoid adding blank rows to tables when rows are removed. 
            '        If Not IsNothing(pfsl.bottom_depth) Or Not IsNothing(pfsl.effective_soil_density) Or Not IsNothing(pfsl.cohesion) Or Not IsNothing(pfsl.friction_angle) Or Not IsNothing(pfsl.spt_blow_count) Or Not IsNothing(pfsl.ultimate_skin_friction_comp) Or Not IsNothing(pfsl.ultimate_skin_friction_uplift) Then
            '            Dim tempSoilLayer As String = InsertPileSoilLayer(pfsl)

            '            If Not firstOne Then
            '                mySoils += ",(" & tempSoilLayer & ")"
            '            Else
            '                mySoils += "(" & tempSoilLayer & ")"
            '            End If

            '            firstOne = False
            '        End If
            '    Next 'Add Soil Layer INSERT statments

            'Insert = Insert.Replace("--[SOIL LAYER INSERT]", SoilLayers.Insert)
            '    firstOne = True
            ''Else
            '    PileSaver = PileSaver.Replace("INSERT INTO fnd.pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])", "--INSERT INTO fnd.pile_soil_layer VALUES ([INSERT ALL SOIL LAYERS])")
            'End If

            Insert = Insert.Replace("--BEGIN --[SOIL LAYER INSERT BEGIN]", "BEGIN --[SOIL LAYER INSERT BEGIN]")
            Insert = Insert.Replace("--END --[SOIL LAYER INSERT END]", "END --[SOIL LAYER INSERT END]")
            For Each row As SoilLayers In SoilLayers
                'For Each row In SoilLayers
                'SoilLayers.Insert()
                'Piles(row)
                '.Insert.
                Insert = Insert.Replace("--[SOIL LAYER INSERT]", row.Insert)
                'Dim SoilLayer As String = QueryBuilderFromFile(queryPath & "CCIplate\CCIplate (SUB QUERY Bolt Group).sql")
            Next

            Return Insert
        End Get
    End Property

    Public Overrides ReadOnly Property Update() As String
        Get
            If _Update = "" Then
                _Update = QueryBuilderFromFile(queryPath & "Pile\Pile (UPDATE).sql")
            End If
            Dim UpdateString As String = _Update
            UpdateString = UpdateString.Replace("[ID]", Me.ID.ToString.FormatDBValue)
            UpdateString = UpdateString.Replace("[UPDATE]", Me.SQLUpdate)
            'UpdateString = UpdateString.Replace("[RESULTS]", Me.Results.EDSResultQuery)
            Return UpdateString
        End Get
    End Property

    Public Overrides ReadOnly Property Delete() As String
        Get
            If _Delete = "" Then
                _Delete = QueryBuilderFromFile(queryPath & "Pile\Pile (DELETE).sql")
            End If
            Dim DeleteString As String = _Delete
            DeleteString = DeleteString.Replace("[ID]", Me.ID.ToString.FormatDBValue)
            DeleteString = DeleteString.Replace("--[SOIL PROFILE INSERT]", SoilProfile.Delete)
            'DeleteString = DeleteString.Replace("--[SOIL LAYER INSERT]", SoilLayers.Delete)
            For Each row As SoilLayers In SoilLayers
                DeleteString = DeleteString.Replace("--[SOIL LAYER INSERT]", row.Delete)
            Next
            Return DeleteString
        End Get
    End Property

#End Region

#Region "Define"

    'Private _ID As Integer?
    'Private _bus_unit As String
    'Private _structure_id As String
    Private _load_eccentricity As Double?
    Private _bolt_circle_bearing_plate_width As Double?
    Private _pile_shape As String
    Private _pile_material As String
    Private _pile_length As Double?
    Private _pile_diameter_width As Double?
    Private _pile_pipe_thickness As Double?
    Private _pile_soil_capacity_given As Boolean?
    Private _steel_yield_strength As Double?
    Private _pile_type_option As String
    Private _rebar_quantity As Double?
    Private _pile_group_config As String
    Private _foundation_depth As Double?
    Private _pad_thickness As Double?
    Private _pad_width_dir1 As Double?
    Private _pad_width_dir2 As Double?
    Private _pad_rebar_size_bottom As Integer?
    Private _pad_rebar_size_top As Integer?
    Private _pad_rebar_quantity_bottom_dir1 As Double?
    Private _pad_rebar_quantity_top_dir1 As Double?
    Private _pad_rebar_quantity_bottom_dir2 As Double?
    Private _pad_rebar_quantity_top_dir2 As Double?
    Private _pier_shape As String
    Private _pier_diameter As Double?
    Private _extension_above_grade As Double?
    Private _pier_rebar_size As Integer?
    Private _pier_rebar_quantity As Double?
    Private _pier_tie_size As Integer?
    Private _rebar_grade As Double?
    Private _concrete_compressive_strength As Double?
    Private _groundwater_depth As Double?
    Private _total_soil_unit_weight As Double?
    Private _cohesion As Double?
    Private _friction_angle As Double?
    Private _neglect_depth As Double?
    Private _spt_blow_count As Double?
    Private _pile_negative_friction_force As Double?
    Private _pile_ultimate_compression As Double?
    Private _pile_ultimate_tension As Double?
    Private _top_and_bottom_rebar_different As Boolean?
    Private _ultimate_gross_end_bearing As Double?
    Private _skin_friction_given As Boolean?
    Private _pile_quantity_circular As Double?
    Private _group_diameter_circular As Double?
    Private _pile_column_quantity As Double?
    Private _pile_row_quantity As Double?
    Private _pile_columns_spacing As Double?
    Private _pile_row_spacing As Double?
    Private _group_efficiency_factor_given As Boolean?
    Private _group_efficiency_factor As Double?
    Private _cap_type As String
    Private _pile_quantity_asymmetric As Double?
    Private _pile_spacing_min_asymmetric As Double?
    Private _quantity_piles_surrounding As Double?
    Private _pile_cap_reference As String
    Private _Soil_110 As Boolean?
    Private _Structural_105 As Boolean?
    Private _soil_profile_id As Integer?
    'Private _tool_version As String
    'Private _modified_person_id As Integer?
    'Private _process_stage As String

    'Private _tia_current As String
    'Private _rev_h_section_15_5 As Boolean?
    'Private _load_z As Boolean?
    'Need to also capture following
    '-BU #
    '-Site Name
    '-Order
    '-Tower Type
    '-Reactions

    'Private _overall_tower_height As Double? 'TNX  *not in piles
    'Private _base_face_width As Double? 'TNX *not in piles
    'Private _bp_dist_above_fnd As Double? 'CCIplate *not in piles
    'Private _ar_bolt_circle As Double? 'CCIplate *not in piles
    'Private _seismic_design_category As String 'Seismic Tool? *not in piles

    'Public Property SoilProfiles As New List(Of SoilProfile)
    Public Property SoilProfile As SoilProfile
    'Public Property SoilLayer As SoilLayer
    Public Property SoilLayers As New List(Of SoilLayers)


    '<Category("Pile"), Description(""), DisplayName("Id")>
    'Public Property ID() As Integer?
    '    Get
    '        Return Me._ID
    '    End Get
    '    Set
    '        Me._ID = Value
    '    End Set
    'End Property
    '<Category("Pile"), Description(""), DisplayName("Bus Unit")>
    'Public Property bus_unit() As String
    '    Get
    '        Return Me._bus_unit
    '    End Get
    '    Set
    '        Me._bus_unit = Value
    '    End Set
    'End Property
    '<Category("Pile"), Description(""), DisplayName("Structure Id")>
    'Public Property structure_id() As String
    '    Get
    '        Return Me._structure_id
    '    End Get
    '    Set
    '        Me._structure_id = Value
    '    End Set
    'End Property
    <Category("Pile"), Description(""), DisplayName("Load Eccentricity")>
    Public Property load_eccentricity() As Double?
        Get
            Return Me._load_eccentricity
        End Get
        Set
            Me._load_eccentricity = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Bolt Circle Bearing Plate Width")>
    Public Property bolt_circle_bearing_plate_width() As Double?
        Get
            Return Me._bolt_circle_bearing_plate_width
        End Get
        Set
            Me._bolt_circle_bearing_plate_width = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Shape")>
    Public Property pile_shape() As String
        Get
            Return Me._pile_shape
        End Get
        Set
            Me._pile_shape = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Material")>
    Public Property pile_material() As String
        Get
            Return Me._pile_material
        End Get
        Set
            Me._pile_material = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Length")>
    Public Property pile_length() As Double?
        Get
            Return Me._pile_length
        End Get
        Set
            Me._pile_length = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Diameter Width")>
    Public Property pile_diameter_width() As Double?
        Get
            Return Me._pile_diameter_width
        End Get
        Set
            Me._pile_diameter_width = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Pipe Thickness")>
    Public Property pile_pipe_thickness() As Double?
        Get
            Return Me._pile_pipe_thickness
        End Get
        Set
            Me._pile_pipe_thickness = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Soil Capacity Given")>
    Public Property pile_soil_capacity_given() As Boolean?
        Get
            Return Me._pile_soil_capacity_given
        End Get
        Set
            Me._pile_soil_capacity_given = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Steel Yield Strength")>
    Public Property steel_yield_strength() As Double?
        Get
            Return Me._steel_yield_strength
        End Get
        Set
            Me._steel_yield_strength = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Type Option")>
    Public Property pile_type_option() As String
        Get
            Return Me._pile_type_option
        End Get
        Set
            Me._pile_type_option = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Rebar Quantity")>
    Public Property rebar_quantity() As Double?
        Get
            Return Me._rebar_quantity
        End Get
        Set
            Me._rebar_quantity = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Group Config")>
    Public Property pile_group_config() As String
        Get
            Return Me._pile_group_config
        End Get
        Set
            Me._pile_group_config = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Foundation Depth")>
    Public Property foundation_depth() As Double?
        Get
            Return Me._foundation_depth
        End Get
        Set
            Me._foundation_depth = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pad Thickness")>
    Public Property pad_thickness() As Double?
        Get
            Return Me._pad_thickness
        End Get
        Set
            Me._pad_thickness = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pad Width Dir1")>
    Public Property pad_width_dir1() As Double?
        Get
            Return Me._pad_width_dir1
        End Get
        Set
            Me._pad_width_dir1 = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pad Width Dir2")>
    Public Property pad_width_dir2() As Double?
        Get
            Return Me._pad_width_dir2
        End Get
        Set
            Me._pad_width_dir2 = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pad Rebar Size Bottom")>
    Public Property pad_rebar_size_bottom() As Integer?
        Get
            Return Me._pad_rebar_size_bottom
        End Get
        Set
            Me._pad_rebar_size_bottom = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pad Rebar Size Top")>
    Public Property pad_rebar_size_top() As Integer?
        Get
            Return Me._pad_rebar_size_top
        End Get
        Set
            Me._pad_rebar_size_top = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pad Rebar Quantity Bottom Dir1")>
    Public Property pad_rebar_quantity_bottom_dir1() As Double?
        Get
            Return Me._pad_rebar_quantity_bottom_dir1
        End Get
        Set
            Me._pad_rebar_quantity_bottom_dir1 = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pad Rebar Quantity Top Dir1")>
    Public Property pad_rebar_quantity_top_dir1() As Double?
        Get
            Return Me._pad_rebar_quantity_top_dir1
        End Get
        Set
            Me._pad_rebar_quantity_top_dir1 = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pad Rebar Quantity Bottom Dir2")>
    Public Property pad_rebar_quantity_bottom_dir2() As Double?
        Get
            Return Me._pad_rebar_quantity_bottom_dir2
        End Get
        Set
            Me._pad_rebar_quantity_bottom_dir2 = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pad Rebar Quantity Top Dir2")>
    Public Property pad_rebar_quantity_top_dir2() As Double?
        Get
            Return Me._pad_rebar_quantity_top_dir2
        End Get
        Set
            Me._pad_rebar_quantity_top_dir2 = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pier Shape")>
    Public Property pier_shape() As String
        Get
            Return Me._pier_shape
        End Get
        Set
            Me._pier_shape = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pier Diameter")>
    Public Property pier_diameter() As Double?
        Get
            Return Me._pier_diameter
        End Get
        Set
            Me._pier_diameter = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Extension Above Grade")>
    Public Property extension_above_grade() As Double?
        Get
            Return Me._extension_above_grade
        End Get
        Set
            Me._extension_above_grade = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pier Rebar Size")>
    Public Property pier_rebar_size() As Integer?
        Get
            Return Me._pier_rebar_size
        End Get
        Set
            Me._pier_rebar_size = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pier Rebar Quantity")>
    Public Property pier_rebar_quantity() As Double?
        Get
            Return Me._pier_rebar_quantity
        End Get
        Set
            Me._pier_rebar_quantity = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pier Tie Size")>
    Public Property pier_tie_size() As Integer?
        Get
            Return Me._pier_tie_size
        End Get
        Set
            Me._pier_tie_size = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Rebar Grade")>
    Public Property rebar_grade() As Double?
        Get
            Return Me._rebar_grade
        End Get
        Set
            Me._rebar_grade = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Concrete Compressive Strength")>
    Public Property concrete_compressive_strength() As Double?
        Get
            Return Me._concrete_compressive_strength
        End Get
        Set
            Me._concrete_compressive_strength = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Groundwater Depth")>
    Public Property groundwater_depth() As Double?
        Get
            Return Me._groundwater_depth
        End Get
        Set
            Me._groundwater_depth = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Total Soil Unit Weight")>
    Public Property total_soil_unit_weight() As Double?
        Get
            Return Me._total_soil_unit_weight
        End Get
        Set
            Me._total_soil_unit_weight = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Cohesion")>
    Public Property cohesion() As Double?
        Get
            Return Me._cohesion
        End Get
        Set
            Me._cohesion = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Friction Angle")>
    Public Property friction_angle() As Double?
        Get
            Return Me._friction_angle
        End Get
        Set
            Me._friction_angle = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Neglect Depth")>
    Public Property neglect_depth() As Double?
        Get
            Return Me._neglect_depth
        End Get
        Set
            Me._neglect_depth = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Spt Blow Count")>
    Public Property spt_blow_count() As Double?
        Get
            Return Me._spt_blow_count
        End Get
        Set
            Me._spt_blow_count = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Pile Negative Friction Force")>
    Public Property pile_negative_friction_force() As Double?
        Get
            Return Me._pile_negative_friction_force
        End Get
        Set
            Me._pile_negative_friction_force = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Ultimate Compression")>
    Public Property pile_ultimate_compression() As Double?
        Get
            Return Me._pile_ultimate_compression
        End Get
        Set
            Me._pile_ultimate_compression = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Ultimate Tension")>
    Public Property pile_ultimate_tension() As Double?
        Get
            Return Me._pile_ultimate_tension
        End Get
        Set
            Me._pile_ultimate_tension = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Top And Bottom Rebar Different")>
    Public Property top_and_bottom_rebar_different() As Boolean?
        Get
            Return Me._top_and_bottom_rebar_different
        End Get
        Set
            Me._top_and_bottom_rebar_different = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Ultimate Gross End Bearing")>
    Public Property ultimate_gross_end_bearing() As Double?
        Get
            Return Me._ultimate_gross_end_bearing
        End Get
        Set
            Me._ultimate_gross_end_bearing = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Skin Friction Given")>
    Public Property skin_friction_given() As Boolean?
        Get
            Return Me._skin_friction_given
        End Get
        Set
            Me._skin_friction_given = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Quantity Circular")>
    Public Property pile_quantity_circular() As Double?
        Get
            Return Me._pile_quantity_circular
        End Get
        Set
            Me._pile_quantity_circular = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Group Diameter Circular")>
    Public Property group_diameter_circular() As Double?
        Get
            Return Me._group_diameter_circular
        End Get
        Set
            Me._group_diameter_circular = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Column Quantity")>
    Public Property pile_column_quantity() As Double?
        Get
            Return Me._pile_column_quantity
        End Get
        Set
            Me._pile_column_quantity = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Row Quantity")>
    Public Property pile_row_quantity() As Double?
        Get
            Return Me._pile_row_quantity
        End Get
        Set
            Me._pile_row_quantity = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Columns Spacing")>
    Public Property pile_columns_spacing() As Double?
        Get
            Return Me._pile_columns_spacing
        End Get
        Set
            Me._pile_columns_spacing = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Row Spacing")>
    Public Property pile_row_spacing() As Double?
        Get
            Return Me._pile_row_spacing
        End Get
        Set
            Me._pile_row_spacing = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Group Efficiency Factor Given")>
    Public Property group_efficiency_factor_given() As Boolean?
        Get
            Return Me._group_efficiency_factor_given
        End Get
        Set
            Me._group_efficiency_factor_given = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Group Efficiency Factor")>
    Public Property group_efficiency_factor() As Double?
        Get
            Return Me._group_efficiency_factor
        End Get
        Set
            Me._group_efficiency_factor = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Cap Type")>
    Public Property cap_type() As String
        Get
            Return Me._cap_type
        End Get
        Set
            Me._cap_type = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Quantity Asymmetric")>
    Public Property pile_quantity_asymmetric() As Double?
        Get
            Return Me._pile_quantity_asymmetric
        End Get
        Set
            Me._pile_quantity_asymmetric = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Spacing Min Asymmetric")>
    Public Property pile_spacing_min_asymmetric() As Double?
        Get
            Return Me._pile_spacing_min_asymmetric
        End Get
        Set
            Me._pile_spacing_min_asymmetric = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Quantity Piles Surrounding")>
    Public Property quantity_piles_surrounding() As Double?
        Get
            Return Me._quantity_piles_surrounding
        End Get
        Set
            Me._quantity_piles_surrounding = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Pile Cap Reference")>
    Public Property pile_cap_reference() As String
        Get
            Return Me._pile_cap_reference
        End Get
        Set
            Me._pile_cap_reference = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Soil 110")>
    Public Property Soil_110() As Boolean?
        Get
            Return Me._Soil_110
        End Get
        Set
            Me._Soil_110 = Value
        End Set
    End Property
    <Category("Pile"), Description(""), DisplayName("Structural 105")>
    Public Property Structural_105() As Boolean?
        Get
            Return Me._Structural_105
        End Get
        Set
            Me._Structural_105 = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Soil Profile ID")>
    Public Property Soil_Profile_id() As Integer?
        Get
            Return Me._soil_profile_id
        End Get
        Set
            Me._soil_profile_id = Value
        End Set
    End Property
    '<Category("Pile"), Description(""), DisplayName("Tool Version")>
    'Public Property tool_version() As String
    '    Get
    '        Return Me._tool_version
    '    End Get
    '    Set
    '        Me._tool_version = Value
    '    End Set
    'End Property
    '<Category("Pile"), Description(""), DisplayName("Modified Person Id")>
    'Public Property modified_person_id() As Integer?
    '    Get
    '        Return Me._modified_person_id
    '    End Get
    '    Set
    '        Me._modified_person_id = Value
    '    End Set
    'End Property
    '<Category("Pile"), Description(""), DisplayName("Process Stage")>
    'Public Property process_stage() As String
    '    Get
    '        Return Me._process_stage
    '    End Get
    '    Set
    '        Me._process_stage = Value
    '    End Set
    'End Property
    '<Category("Pile"), Description(""), DisplayName("TIA")>
    'Public Property tia_current() As String
    '    Get
    '        Return If(Me.ParentStructure.structureCodeCriteria.tia_current, Me._tia_current)
    '    End Get
    '    Set
    '        Me._tia_current = Value
    '    End Set
    'End Property
    '<Category("Pile"), Description(""), DisplayName("Rev H Section 15.5")>
    'Public Property rev_h_section_15_5() As Boolean?
    '    Get
    '        Return If(Me.ParentStructure.structureCodeCriteria.rev_h_section_15_5, Me._rev_h_section_15_5)
    '    End Get
    '    Set
    '        Me._rev_h_section_15_5 = Value
    '    End Set
    'End Property
    'Load Z
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByRef Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        'Get values from structure code criteria
        'Not sure this is necessary, could just read the values from the structure code criteria when creating the Excel sheet
        'Me.tia_current = Me.ParentStructure?.structureCodeCriteria?.tia_current
        'Me.rev_h_section_15_5 = Me.ParentStructure?.structureCodeCriteria?.rev_h_section_15_5
        'Me.seismic_design_category = Me.ParentStructure?.structureCodeCriteria?.seismic_design_category

        ''''''Customize for each foundation type'''''
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.bus_unit = DBtoStr(dr.Item("bus_unit"))
        Me.structure_id = DBtoStr(dr.Item("structure_id"))
        Me.load_eccentricity = DBtoNullableDbl(dr.Item("load_eccentricity"))
        Me.bolt_circle_bearing_plate_width = DBtoNullableDbl(dr.Item("bolt_circle_bearing_plate_width"))
        Me.pile_shape = DBtoStr(dr.Item("pile_shape"))
        Me.pile_material = DBtoStr(dr.Item("pile_material"))
        Me.pile_length = DBtoNullableDbl(dr.Item("pile_length"))
        Me.pile_diameter_width = DBtoNullableDbl(dr.Item("pile_diameter_width"))
        Me.pile_pipe_thickness = DBtoNullableDbl(dr.Item("pile_pipe_thickness"))
        Me.pile_soil_capacity_given = DBtoNullableBool(dr.Item("pile_soil_capacity_given"))
        Me.steel_yield_strength = DBtoNullableDbl(dr.Item("steel_yield_strength"))
        Me.pile_type_option = DBtoStr(dr.Item("pile_type_option"))
        Me.rebar_quantity = DBtoNullableDbl(dr.Item("rebar_quantity"))
        Me.pile_group_config = DBtoStr(dr.Item("pile_group_config"))
        Me.foundation_depth = DBtoNullableDbl(dr.Item("foundation_depth"))
        Me.pad_thickness = DBtoNullableDbl(dr.Item("pad_thickness"))
        Me.pad_width_dir1 = DBtoNullableDbl(dr.Item("pad_width_dir1"))
        Me.pad_width_dir2 = DBtoNullableDbl(dr.Item("pad_width_dir2"))
        Me.pad_rebar_size_bottom = DBtoNullableInt(dr.Item("pad_rebar_size_bottom"))
        Me.pad_rebar_size_top = DBtoNullableInt(dr.Item("pad_rebar_size_top"))
        Me.pad_rebar_quantity_bottom_dir1 = DBtoNullableDbl(dr.Item("pad_rebar_quantity_bottom_dir1"))
        Me.pad_rebar_quantity_top_dir1 = DBtoNullableDbl(dr.Item("pad_rebar_quantity_top_dir1"))
        Me.pad_rebar_quantity_bottom_dir2 = DBtoNullableDbl(dr.Item("pad_rebar_quantity_bottom_dir2"))
        Me.pad_rebar_quantity_top_dir2 = DBtoNullableDbl(dr.Item("pad_rebar_quantity_top_dir2"))
        Me.pier_shape = DBtoStr(dr.Item("pier_shape"))
        Me.pier_diameter = DBtoNullableDbl(dr.Item("pier_diameter"))
        Me.extension_above_grade = DBtoNullableDbl(dr.Item("extension_above_grade"))
        Me.pier_rebar_size = DBtoNullableInt(dr.Item("pier_rebar_size"))
        Me.pier_rebar_quantity = DBtoNullableDbl(dr.Item("pier_rebar_quantity"))
        Me.pier_tie_size = DBtoNullableInt(dr.Item("pier_tie_size"))
        Me.rebar_grade = DBtoNullableDbl(dr.Item("rebar_grade"))
        Me.concrete_compressive_strength = DBtoNullableDbl(dr.Item("concrete_compressive_strength"))
        Me.groundwater_depth = DBtoNullableDbl(dr.Item("groundwater_depth"))
        Me.total_soil_unit_weight = DBtoNullableDbl(dr.Item("total_soil_unit_weight"))
        Me.cohesion = DBtoNullableDbl(dr.Item("cohesion"))
        Me.friction_angle = DBtoNullableDbl(dr.Item("friction_angle"))
        Me.neglect_depth = DBtoNullableDbl(dr.Item("neglect_depth"))
        Me.spt_blow_count = DBtoNullableDbl(dr.Item("spt_blow_count"))
        Me.pile_negative_friction_force = DBtoNullableDbl(dr.Item("pile_negative_friction_force"))
        Me.pile_ultimate_compression = DBtoNullableDbl(dr.Item("pile_ultimate_compression"))
        Me.pile_ultimate_tension = DBtoNullableDbl(dr.Item("pile_ultimate_tension"))
        Me.top_and_bottom_rebar_different = DBtoNullableBool(dr.Item("top_and_bottom_rebar_different"))
        Me.ultimate_gross_end_bearing = DBtoNullableDbl(dr.Item("ultimate_gross_end_bearing"))
        Me.skin_friction_given = DBtoNullableBool(dr.Item("skin_friction_given"))
        Me.pile_quantity_circular = DBtoNullableDbl(dr.Item("pile_quantity_circular"))
        Me.group_diameter_circular = DBtoNullableDbl(dr.Item("group_diameter_circular"))
        Me.pile_column_quantity = DBtoNullableDbl(dr.Item("pile_column_quantity"))
        Me.pile_row_quantity = DBtoNullableDbl(dr.Item("pile_row_quantity"))
        Me.pile_columns_spacing = DBtoNullableDbl(dr.Item("pile_columns_spacing"))
        Me.pile_row_spacing = DBtoNullableDbl(dr.Item("pile_row_spacing"))
        Me.group_efficiency_factor_given = DBtoNullableBool(dr.Item("group_efficiency_factor_given"))
        Me.group_efficiency_factor = DBtoNullableDbl(dr.Item("group_efficiency_factor"))
        Me.cap_type = DBtoStr(dr.Item("cap_type"))
        Me.pile_quantity_asymmetric = DBtoNullableDbl(dr.Item("pile_quantity_asymmetric"))
        Me.pile_spacing_min_asymmetric = DBtoNullableDbl(dr.Item("pile_spacing_min_asymmetric"))
        Me.quantity_piles_surrounding = DBtoNullableDbl(dr.Item("quantity_piles_surrounding"))
        Me.pile_cap_reference = DBtoStr(dr.Item("pile_cap_reference"))
        Me.Soil_110 = DBtoNullableBool(dr.Item("Soil_110"))
        Me.Structural_105 = DBtoNullableBool(dr.Item("Structural_105"))
        Me.tool_version = DBtoStr(dr.Item("tool_version"))
        Me.Soil_Profile_id = DBtoNullableInt(dr.Item("soil_profile_id"))
        Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
        Me.process_stage = DBtoStr(dr.Item("process_stage"))

        If ds.Tables.Contains("PILE SOIL PROFILE EXCEL") Then

            Me.SoilProfile = New SoilProfile(excelDS.Tables("PILE SOIL PROFILE EXCEL").Rows(0))

        End If

        For Each Row As DataRow In ds.Tables.Contains

            'Need to add in information for soil profile and soil layers here
            If Me.pile_soil_capacity_given = False And Me.pile_shape <> "H-Pile" Then
            For Each SoilLayerDataRow As DataRow In ds.Tables("Pile Soil SQL").Rows
                Dim soilRefID As Integer = CType(SoilLayerDataRow.Item("pile_fnd_id"), Integer)
                If soilRefID = refID Then
                    Me.soil_layers.Add(New PileSoilLayer(SoilLayerDataRow))
                End If
            Next 'Add Soild Layers to to Pile Soil Layer Object
        End If

        'If Me.pile_group_config = "Asymmetric" Then
        '    For Each LocationDataRow As DataRow In ds.Tables("Pile Location SQL").Rows
        '        Dim locRefID As Integer = CType(LocationDataRow.Item("pile_fnd_id"), Integer)
        '        If locRefID = refID Then
        '            Me.pile_locations.Add(New PileLocation(LocationDataRow))
        '        End If
        '    Next 'Add Soild Layers to to Pile Location Object
        'End If

    End Sub 'Generate a pf from EDS

    'Public Sub New(ExcelFilePath As String, Optional BU As String = Nothing, Optional structureID As String = Nothing)
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

        If excelDS.Tables.Contains("Pile General Details EXCEL") Then
            Dim dr = excelDS.Tables("Pile General Details EXCEL").Rows(0)
            'Need to dimension DataRow from GenStructure/TNX and anywhere else inputs may come from as well - MRR

            Me.ID = DBtoNullableInt(dr.Item("pile_id"))
            'Me.bus_unit = DBtoStr(dr.Item("bus_unit"))
            'Me.structure_id = DBtoStr(dr.Item("structure_id"))
            Me.load_eccentricity = DBtoNullableDbl(dr.Item("load_eccentricity"))
            Me.bolt_circle_bearing_plate_width = DBtoNullableDbl(dr.Item("bolt_circle_bearing_plate_width"))
            Me.pile_shape = DBtoStr(dr.Item("pile_shape"))
            Me.pile_material = DBtoStr(dr.Item("pile_material"))
            Me.pile_length = DBtoNullableDbl(dr.Item("pile_length"))
            Me.pile_diameter_width = DBtoNullableDbl(dr.Item("pile_diameter_width"))
            Me.pile_pipe_thickness = DBtoNullableDbl(dr.Item("pile_pipe_thickness"))
            Me.pile_soil_capacity_given = DBtoNullableBool(dr.Item("pile_soil_capacity_given"))
            'Me.pile_soil_capacity_given = If(DBtoStr(dr.Item("pile_soil_capacity_given")) = "yes", True)
            'Me.pile_soil_capacity_given = trueFalseYesNo(dr.Item("pile_soil_capacity_given"))
            Me.steel_yield_strength = DBtoNullableDbl(dr.Item("steel_yield_strength"))
            Me.pile_type_option = DBtoStr(dr.Item("pile_type_option"))
            Me.rebar_quantity = DBtoNullableDbl(dr.Item("rebar_quantity"))
            Me.pile_group_config = DBtoStr(dr.Item("pile_group_config"))
            Me.foundation_depth = DBtoNullableDbl(dr.Item("foundation_depth"))
            Me.pad_thickness = DBtoNullableDbl(dr.Item("pad_thickness"))
            Me.pad_width_dir1 = DBtoNullableDbl(dr.Item("pad_width_dir1"))
            Me.pad_width_dir2 = DBtoNullableDbl(dr.Item("pad_width_dir2"))
            Me.pad_rebar_size_bottom = DBtoNullableInt(dr.Item("pad_rebar_size_bottom"))
            Me.pad_rebar_size_top = DBtoNullableInt(dr.Item("pad_rebar_size_top"))
            Me.pad_rebar_quantity_bottom_dir1 = DBtoNullableDbl(dr.Item("pad_rebar_quantity_bottom_dir1"))
            Me.pad_rebar_quantity_top_dir1 = DBtoNullableDbl(dr.Item("pad_rebar_quantity_top_dir1"))
            Me.pad_rebar_quantity_bottom_dir2 = DBtoNullableDbl(dr.Item("pad_rebar_quantity_bottom_dir2"))
            Me.pad_rebar_quantity_top_dir2 = DBtoNullableDbl(dr.Item("pad_rebar_quantity_top_dir2"))
            Me.pier_shape = DBtoStr(dr.Item("pier_shape"))
            Me.pier_diameter = DBtoNullableDbl(dr.Item("pier_diameter"))
            Me.extension_above_grade = DBtoNullableDbl(dr.Item("extension_above_grade"))
            Me.pier_rebar_size = DBtoNullableInt(dr.Item("pier_rebar_size"))
            Me.pier_rebar_quantity = DBtoNullableDbl(dr.Item("pier_rebar_quantity"))
            Me.pier_tie_size = DBtoNullableInt(dr.Item("pier_tie_size"))
            Me.rebar_grade = DBtoNullableDbl(dr.Item("rebar_grade"))
            Me.concrete_compressive_strength = DBtoNullableDbl(dr.Item("concrete_compressive_strength"))
            Me.groundwater_depth = DBtoNullableDbl(dr.Item("groundwater_depth"))
            Me.total_soil_unit_weight = DBtoNullableDbl(dr.Item("total_soil_unit_weight"))
            Me.cohesion = DBtoNullableDbl(dr.Item("cohesion"))
            Me.friction_angle = DBtoNullableDbl(dr.Item("friction_angle"))
            Me.neglect_depth = DBtoNullableDbl(dr.Item("neglect_depth"))
            Me.spt_blow_count = DBtoNullableDbl(dr.Item("spt_blow_count"))
            Me.pile_negative_friction_force = DBtoNullableDbl(dr.Item("pile_negative_friction_force"))
            Me.pile_ultimate_compression = DBtoNullableDbl(dr.Item("pile_ultimate_compression"))
            Me.pile_ultimate_tension = DBtoNullableDbl(dr.Item("pile_ultimate_tension"))
            Me.top_and_bottom_rebar_different = DBtoNullableBool(dr.Item("top_and_bottom_rebar_different"))
            Me.ultimate_gross_end_bearing = DBtoNullableDbl(dr.Item("ultimate_gross_end_bearing"))
            Me.skin_friction_given = DBtoNullableBool(dr.Item("skin_friction_given"))
            Me.pile_quantity_circular = DBtoNullableDbl(dr.Item("pile_quantity_circular"))
            Me.group_diameter_circular = DBtoNullableDbl(dr.Item("group_diameter_circular"))
            Me.pile_column_quantity = DBtoNullableDbl(dr.Item("pile_column_quantity"))
            Me.pile_row_quantity = DBtoNullableDbl(dr.Item("pile_row_quantity"))
            Me.pile_columns_spacing = DBtoNullableDbl(dr.Item("pile_columns_spacing"))
            Me.pile_row_spacing = DBtoNullableDbl(dr.Item("pile_row_spacing"))
            Me.group_efficiency_factor_given = DBtoNullableBool(dr.Item("group_efficiency_factor_given"))
            Me.group_efficiency_factor = DBtoNullableDbl(dr.Item("group_efficiency_factor"))
            Me.cap_type = DBtoStr(dr.Item("cap_type"))
            Me.pile_quantity_asymmetric = DBtoNullableDbl(dr.Item("pile_quantity_asymmetric"))
            Me.pile_spacing_min_asymmetric = DBtoNullableDbl(dr.Item("pile_spacing_min_asymmetric"))
            Me.quantity_piles_surrounding = DBtoNullableDbl(dr.Item("quantity_piles_surrounding"))
            Me.pile_cap_reference = DBtoStr(dr.Item("pile_cap_reference"))
            Me.Soil_110 = DBtoNullableBool(dr.Item("Soil_110"))
            Me.Structural_105 = DBtoNullableBool(dr.Item("Structural_105"))
            Me.tool_version = DBtoStr(dr.Item("tool_version"))
            'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
            'Me.process_stage = DBtoStr(dr.Item("process_stage"))

        End If

        If excelDS.Tables.Contains("PILE SOIL PROFILE EXCEL") Then

            Me.SoilProfile = New SoilProfile(excelDS.Tables("PILE SOIL PROFILE EXCEL").Rows(0))
            'SoilProfile(SoilProfile(Row, 0))

            'For Each Row As DataRow In excelDS.Tables("PILE SOIL PROFILE EXCEL").Rows

            '    'For Tools with multiple foundation or sub items, use Row.Item("ID") or add a local_ID column to filter which results should be associated with each foundation

            '    Me.SoilProfile.Add(New SoilProfile(Row))

            'Next

        End If

        If excelDS.Tables.Contains("PILE SOIL LAYER EXCEL") Then

            For Each Row As DataRow In excelDS.Tables("Pile Soil Layer EXCEL").Rows
                'For Each dr As DataRow In ds.Tables("Pile Soil EXCEL").Rows
                Me.SoilLayers.Add(New SoilLayers(Row))
            Next 'Add Soil Layers to to Pile Soil Layer Object


            'Me.SoilLayer = New SoilLayer(excelDS.Tables("PILE SOIL LAYER EXCEL").Rows(0))


        End If




        'If excelDS.Tables.Contains("Pier and Pad General Results EXCEL") Then

        '    For Each Row As DataRow In excelDS.Tables("Pier and Pad General Results EXCEL").Rows

        '        'For Tools with multiple foundation or sub items, use Row.Item("ID") or add a local_ID column to filter which results should be associated with each foundation

        '        Me.Results.Add(New EDSResult(Row, Me))

        '    Next

        'End If

    End Sub 'Generate a ub from Excel

#End Region

#Region "Save to Excel"

    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        '''''Customize for each excel tool'''''

        With wb
            If Not IsNothing(Me.ID) Then
                .Worksheets("Input").Range("ID").Value = CType(Me.ID, Integer)
            Else
                .Worksheets("Input").Range("ID").ClearContents
            End If
            If Not IsNothing(Me.bus_unit) Then
                .Worksheets("").Range("").Value = CType(Me.bus_unit, String)
            End If
            If Not IsNothing(Me.structure_id) Then
                .Worksheets("").Range("").Value = CType(Me.structure_id, String)
            End If
            If Not IsNothing(Me.load_eccentricity) Then
                .Worksheets("Input").Range("").Value = CType(Me.load_eccentricity, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.bolt_circle_bearing_plate_width) Then
                .Worksheets("Input").Range("").Value = CType(Me.bolt_circle_bearing_plate_width, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pile_shape) Then
                .Worksheets("Input").Range("D23").Value = CType(Me.pile_shape, String)
            End If
            If Not IsNothing(Me.pile_material) Then
                .Worksheets("Input").Range("D24").Value = CType(Me.pile_material, String)
            End If
            If Not IsNothing(Me.pile_length) Then
                .Worksheets("Input").Range("").Value = CType(Me.pile_length, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pile_diameter_width) Then
                .Worksheets("Input").Range("D26").Value = CType(Me.pile_diameter_width, Double)
            Else
                .Worksheets("Input").Range("D26").ClearContents
            End If
            If Not IsNothing(Me.pile_pipe_thickness) Then
                .Worksheets("Input").Range("D27").Value = CType(Me.pile_pipe_thickness, Double)
            Else
                .Worksheets("Input").Range("D27").ClearContents
            End If
            If Me.pile_soil_capacity_given = False Then
                .Worksheets("Input").Range("D29").Value = "No"
            Else
                .Worksheets("Input").Range("D29").Value = "Yes"
            End If
            'If Not IsNothing(Me.pile_soil_capacity_given) Then
            '    .Worksheets("Input").Range("D29").Value = CType(Me.pile_soil_capacity_given, Boolean)
            'End If
            If Not IsNothing(Me.steel_yield_strength) Then
                .Worksheets("Input").Range("D30").Value = CType(Me.steel_yield_strength, Double)
            Else
                .Worksheets("Input").Range("D30").ClearContents
            End If
            If Not IsNothing(Me.pile_type_option) Then
                .Worksheets("Input").Range("").Value = CType(Me.pile_type_option, String)
            End If
            If Not IsNothing(Me.rebar_quantity) Then
                .Worksheets("Input").Range("").Value = CType(Me.rebar_quantity, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pile_group_config) Then
                .Worksheets("Input").Range("").Value = CType(Me.pile_group_config, String)
            End If
            If Not IsNothing(Me.foundation_depth) Then
                .Worksheets("Input").Range("").Value = CType(Me.foundation_depth, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pad_thickness) Then
                .Worksheets("Input").Range("").Value = CType(Me.pad_thickness, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pad_width_dir1) Then
                .Worksheets("Input").Range("").Value = CType(Me.pad_width_dir1, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pad_width_dir2) Then
                .Worksheets("Input").Range("").Value = CType(Me.pad_width_dir2, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pad_rebar_size_bottom) Then
                .Worksheets("Input").Range("").Value = CType(Me.pad_rebar_size_bottom, Integer)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pad_rebar_size_top) Then
                .Worksheets("Input").Range("").Value = CType(Me.pad_rebar_size_top, Integer)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pad_rebar_quantity_bottom_dir1) Then
                .Worksheets("Input").Range("").Value = CType(Me.pad_rebar_quantity_bottom_dir1, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pad_rebar_quantity_top_dir1) Then
                .Worksheets("Input").Range("").Value = CType(Me.pad_rebar_quantity_top_dir1, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pad_rebar_quantity_bottom_dir2) Then
                .Worksheets("Input").Range("").Value = CType(Me.pad_rebar_quantity_bottom_dir2, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pad_rebar_quantity_top_dir2) Then
                .Worksheets("Input").Range("").Value = CType(Me.pad_rebar_quantity_top_dir2, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pier_shape) Then
                .Worksheets("Input").Range("D57").Value = CType(Me.pier_shape, String)
            End If
            If Not IsNothing(Me.pier_diameter) Then
                .Worksheets("Input").Range("").Value = CType(Me.pier_diameter, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.extension_above_grade) Then
                .Worksheets("Input").Range("").Value = CType(Me.extension_above_grade, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pier_rebar_size) Then
                .Worksheets("Input").Range("").Value = CType(Me.pier_rebar_size, Integer)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pier_rebar_quantity) Then
                .Worksheets("Input").Range("").Value = CType(Me.pier_rebar_quantity, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pier_tie_size) Then
                .Worksheets("Input").Range("").Value = CType(Me.pier_tie_size, Integer)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.rebar_grade) Then
                .Worksheets("Input").Range("").Value = CType(Me.rebar_grade, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.concrete_compressive_strength) Then
                .Worksheets("Input").Range("").Value = CType(Me.concrete_compressive_strength, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Me.groundwater_depth = -1 Then
                .Worksheets("Input").Range("D69").Value = "N/A"
            Else
                .Worksheets("Input").Range("D69").Value = CType(Me.groundwater_depth, Double)
            End If
            'If Not IsNothing(Me.groundwater_depth) Then
            '    .Worksheets("Input").Range("D69").Value = CType(Me.groundwater_depth, Double)
            'Else
            '    .Worksheets("Input").Range("D69").ClearContents
            'End If
            If Not IsNothing(Me.total_soil_unit_weight) Then
                .Worksheets("Input").Range("").Value = CType(Me.total_soil_unit_weight, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.cohesion) Then
                .Worksheets("Input").Range("").Value = CType(Me.cohesion, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.friction_angle) Then
                .Worksheets("Input").Range("").Value = CType(Me.friction_angle, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.neglect_depth) Then
                .Worksheets("Input").Range("").Value = CType(Me.neglect_depth, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.spt_blow_count) Then
                .Worksheets("Input").Range("").Value = CType(Me.spt_blow_count, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pile_negative_friction_force) Then
                .Worksheets("Input").Range("").Value = CType(Me.pile_negative_friction_force, Double)
            Else
                .Worksheets("Input").Range("").ClearContents
            End If
            If Not IsNothing(Me.pile_ultimate_compression) Then
                .Worksheets("Input").Range("K45").Value = CType(Me.pile_ultimate_compression, Double)
            Else
                .Worksheets("Input").Range("K45").ClearContents
            End If
            If Not IsNothing(Me.pile_ultimate_tension) Then
                .Worksheets("Input").Range("K46").Value = CType(Me.pile_ultimate_tension, Double)
            Else
                .Worksheets("Input").Range("K46").ClearContents
            End If
            If Not IsNothing(Me.top_and_bottom_rebar_different) Then
                .Worksheets("Input").Range("Z10").Value = CType(Me.top_and_bottom_rebar_different, Boolean)
            End If
            If Not IsNothing(Me.ultimate_gross_end_bearing) Then
                .Worksheets("Input").Range("M71").Value = CType(Me.ultimate_gross_end_bearing, Double)
            Else
                .Worksheets("Input").Range("M71").ClearContents
            End If
            If Me.skin_friction_given = False Then
                .Worksheets("Input").Range("N54").Value = "No"
            Else
                .Worksheets("Input").Range("N54").Value = "Yes"
            End If
            'If Not IsNothing(Me.skin_friction_given) Then
            '    .Worksheets("Input").Range("N54").Value = CType(Me.skin_friction_given, Boolean)
            'End If
            If Not IsNothing(Me.pile_quantity_circular) Then
                .Worksheets("Input").Range("D36").Value = CType(Me.pile_quantity_circular, Double)
            Else
                .Worksheets("Input").Range("D36").ClearContents
            End If
            If Not IsNothing(Me.group_diameter_circular) Then
                .Worksheets("Input").Range("D37").Value = CType(Me.group_diameter_circular, Double)
            Else
                .Worksheets("Input").Range("D37").ClearContents
            End If
            If Not IsNothing(Me.pile_column_quantity) Then
                .Worksheets("Input").Range("D36").Value = CType(Me.pile_column_quantity, Double)
            Else
                .Worksheets("Input").Range("D36").ClearContents
            End If
            If Not IsNothing(Me.pile_row_quantity) Then
                .Worksheets("Input").Range("D37").Value = CType(Me.pile_row_quantity, Double)
            Else
                .Worksheets("Input").Range("D37").ClearContents
            End If
            If Not IsNothing(Me.pile_columns_spacing) Then
                .Worksheets("Input").Range("D38").Value = CType(Me.pile_columns_spacing, Double)
            Else
                .Worksheets("Input").Range("D38").ClearContents
            End If
            If Not IsNothing(Me.pile_row_spacing) Then
                .Worksheets("Input").Range("D39").Value = CType(Me.pile_row_spacing, Double)
            Else
                .Worksheets("Input").Range("D39").ClearContents
            End If
            If Me.group_efficiency_factor_given = False Then
                .Worksheets("Input").Range("D41").Value = "No"
            Else
                .Worksheets("Input").Range("D41").Value = "Yes"
            End If
            'If Not IsNothing(Me.group_efficiency_factor_given) Then
            '    .Worksheets("Input").Range("D41").Value = CType(Me.group_efficiency_factor_given, Boolean)
            'End If
            If Not IsNothing(Me.group_efficiency_factor) Then
                .Worksheets("Input").Range("D42").Value = CType(Me.group_efficiency_factor, Double)
            Else
                .Worksheets("Input").Range("D42").ClearContents
            End If
            If Not IsNothing(Me.cap_type) Then
                .Worksheets("Input").Range("D45").Value = CType(Me.cap_type, String)
            End If
            If Not IsNothing(Me.pile_quantity_asymmetric) Then
                .Worksheets("Moment of Inertia").Range("D10").Value = CType(Me.pile_quantity_asymmetric, Double)
            Else
                .Worksheets("Moment of Inertia").Range("D10").ClearContents
            End If
            If Not IsNothing(Me.pile_spacing_min_asymmetric) Then
                .Worksheets("Moment of Inertia").Range("D11").Value = CType(Me.pile_spacing_min_asymmetric, Double)
            Else
                .Worksheets("Moment of Inertia").Range("D11").ClearContents
            End If
            If Not IsNothing(Me.quantity_piles_surrounding) Then
                .Worksheets("Moment of Inertia").Range("D12").Value = CType(Me.quantity_piles_surrounding, Double)
            Else
                .Worksheets("Moment of Inertia").Range("D12").ClearContents
            End If
            If Not IsNothing(Me.pile_cap_reference) Then
                .Worksheets("Input").Range("G47").Value = CType(Me.pile_cap_reference, String)
            End If
            If Not IsNothing(Me.Soil_110) Then
                .Worksheets("Input").Range("Z13").Value = CType(Me.Soil_110, Boolean)
            End If
            If Not IsNothing(Me.Structural_105) Then
                .Worksheets("Input").Range("Z14").Value = CType(Me.Structural_105, Boolean)
            End If
            If Not IsNothing(Me.tool_version) Then
                .Worksheets("Revision History").Range("").Value = CType(Me.tool_version, String)
            End If
            If Not IsNothing(Me.modified_person_id) Then
                .Worksheets("").Range("").Value = CType(Me.modified_person_id, Integer)
            Else
                .Worksheets("").Range("").ClearContents
            End If
            If Not IsNothing(Me.process_stage) Then
                .Worksheets("").Range("").Value = CType(Me.process_stage, String)
            End If

        End With

    End Sub


    'Sub UnitBaseRunPrint()
    '    Dim xlApp As New excel.Application
    '    Dim xlWb As excel.Workbook
    '    'Dim xlSheet As excel.Worksheet
    '    Dim piernpadtemplateloc As String = ""
    '    Try
    '        With xlApp
    '            xlApp.Visible = False
    '            xlApp.DisplayAlerts = False
    '            xlWb = .Workbooks.Add(piernpadtemplateloc)
    '            'xlSheet = xlWb.Sheets("Data")
    '            'xlSheet.Range("C4").Value = "X"
    '            xlApp.Run("Module1.RunAnalysis")
    '        End With
    '    Catch ex As Exception
    '        Try
    '            xlWb.Close(False)
    '            xlApp.Quit()
    '        Catch
    '            MsgBox("Something didn't close. Figure it out yourself")
    '        End Try
    '    End Try
    '    releaseObject(xlSheet)
    '    releaseObject(xlWb)
    '    releaseObject(xlApp)
    'End Sub
    'Sub UnitBaseRunnPrint(templatepath As String)
    '    Dim xlApp As New Excel.Application
    '    Dim xlWb As Excel.Workbook
    '    Dim xlSheet As Excel.Worksheet
    '    Dim unitbasetemplateloc As String = templatepath
    '    Try
    '        With xlApp
    '            xlApp.Visible = False
    '            xlApp.DisplayAlerts = False
    '            xlWb = xlApp.Workbooks.Add(unitbasetemplateloc)
    '            xlSheet = xlWb.Worksheets("Data")
    '            xlSheet.Range("C4").Value = "X"
    '            xlApp.Run("Module1.RunAnalysis")
    '        End With
    '    Catch ex As Exception
    '        Try
    '            xlWb.Close(False)
    '            xlApp.Quit()
    '        Catch
    '            MsgBox("Something didn't close. Figure it out yourself")
    '        End Try
    '    End Try
    '    releaseObject(xlSheet)
    '    releaseObject(xlWb)
    '    releaseObject(xlApp)
    'End Sub
    'Private Sub releaseObject(ByVal obj As Object)
    '    Try
    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
    '        obj = Nothing
    '    Catch ex As Exception
    '        obj = Nothing
    '    Finally
    '        GC.Collect()
    '    End Try
    'End Sub

#End Region

#Region "Save to EDS"

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.load_eccentricity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_circle_bearing_plate_width.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_shape.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_material.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_length.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_diameter_width.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_pipe_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_soil_capacity_given.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.steel_yield_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_type_option.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_group_config.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.foundation_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_width_dir1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_width_dir2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_size_bottom.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_size_top.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_bottom_dir1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_top_dir1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_bottom_dir2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_top_dir2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_shape.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.extension_above_grade.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_rebar_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_rebar_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_tie_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_grade.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.concrete_compressive_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.groundwater_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.total_soil_unit_weight.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cohesion.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.friction_angle.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.neglect_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.spt_blow_count.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_negative_friction_force.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_ultimate_compression.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_ultimate_tension.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.top_and_bottom_rebar_different.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ultimate_gross_end_bearing.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.skin_friction_given.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_quantity_circular.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.group_diameter_circular.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_column_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_row_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_columns_spacing.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_row_spacing.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.group_efficiency_factor_given.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.group_efficiency_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cap_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_quantity_asymmetric.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_spacing_min_asymmetric.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.quantity_piles_surrounding.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pile_cap_reference.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Soil_110.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Structural_105.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tool_version.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString("@Sub1ID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString("@PrevID")
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("load_eccentricity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_circle_bearing_plate_width")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_shape")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_material")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_length")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_diameter_width")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_pipe_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_soil_capacity_given")
        SQLInsertFields = SQLInsertFields.AddtoDBString("steel_yield_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_type_option")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rebar_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_group_config")
        SQLInsertFields = SQLInsertFields.AddtoDBString("foundation_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_width_dir1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_width_dir2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_size_bottom")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_size_top")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_quantity_bottom_dir1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_quantity_top_dir1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_quantity_bottom_dir2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_quantity_top_dir2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_shape")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("extension_above_grade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_rebar_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_rebar_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_tie_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rebar_grade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("concrete_compressive_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("groundwater_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("total_soil_unit_weight")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cohesion")
        SQLInsertFields = SQLInsertFields.AddtoDBString("friction_angle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("neglect_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("spt_blow_count")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_negative_friction_force")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_ultimate_compression")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_ultimate_tension")
        SQLInsertFields = SQLInsertFields.AddtoDBString("top_and_bottom_rebar_different")
        SQLInsertFields = SQLInsertFields.AddtoDBString("ultimate_gross_end_bearing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("skin_friction_given")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_quantity_circular")
        SQLInsertFields = SQLInsertFields.AddtoDBString("group_diameter_circular")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_column_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_row_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_columns_spacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_row_spacing")
        SQLInsertFields = SQLInsertFields.AddtoDBString("group_efficiency_factor_given")
        SQLInsertFields = SQLInsertFields.AddtoDBString("group_efficiency_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cap_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_quantity_asymmetric")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_spacing_min_asymmetric")
        SQLInsertFields = SQLInsertFields.AddtoDBString("quantity_piles_surrounding")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pile_cap_reference")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Soil_110")
        SQLInsertFields = SQLInsertFields.AddtoDBString("Structural_105")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tool_version")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("soil_profile_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""

        'SQLUpdate = SQLUpdate.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("bus_unit = " & Me.bus_unit.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("structure_id = " & Me.structure_id.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("load_eccentricity = " & Me.load_eccentricity.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("bolt_circle_bearing_plate_width = " & Me.bolt_circle_bearing_plate_width.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_shape = " & Me.pile_shape.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_material = " & Me.pile_material.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_length = " & Me.pile_length.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_diameter_width = " & Me.pile_diameter_width.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_pipe_thickness = " & Me.pile_pipe_thickness.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_soil_capacity_given = " & Me.pile_soil_capacity_given.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("steel_yield_strength = " & Me.steel_yield_strength.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_type_option = " & Me.pile_type_option.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("rebar_quantity = " & Me.rebar_quantity.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_group_config = " & Me.pile_group_config.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("foundation_depth = " & Me.foundation_depth.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_thickness = " & Me.pad_thickness.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_width_dir1 = " & Me.pad_width_dir1.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_width_dir2 = " & Me.pad_width_dir2.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_size_bottom = " & Me.pad_rebar_size_bottom.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_size_top = " & Me.pad_rebar_size_top.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_quantity_bottom_dir1 = " & Me.pad_rebar_quantity_bottom_dir1.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_quantity_top_dir1 = " & Me.pad_rebar_quantity_top_dir1.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_quantity_bottom_dir2 = " & Me.pad_rebar_quantity_bottom_dir2.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_quantity_top_dir2 = " & Me.pad_rebar_quantity_top_dir2.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_shape = " & Me.pier_shape.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_diameter = " & Me.pier_diameter.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("extension_above_grade = " & Me.extension_above_grade.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_rebar_size = " & Me.pier_rebar_size.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_rebar_quantity = " & Me.pier_rebar_quantity.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_tie_size = " & Me.pier_tie_size.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("rebar_grade = " & Me.rebar_grade.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("concrete_compressive_strength = " & Me.concrete_compressive_strength.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("groundwater_depth = " & Me.groundwater_depth.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("total_soil_unit_weight = " & Me.total_soil_unit_weight.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("cohesion = " & Me.cohesion.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("friction_angle = " & Me.friction_angle.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("neglect_depth = " & Me.neglect_depth.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("spt_blow_count = " & Me.spt_blow_count.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_negative_friction_force = " & Me.pile_negative_friction_force.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_ultimate_compression = " & Me.pile_ultimate_compression.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_ultimate_tension = " & Me.pile_ultimate_tension.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("top_and_bottom_rebar_different = " & Me.top_and_bottom_rebar_different.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("ultimate_gross_end_bearing = " & Me.ultimate_gross_end_bearing.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("skin_friction_given = " & Me.skin_friction_given.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_quantity_circular = " & Me.pile_quantity_circular.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("group_diameter_circular = " & Me.group_diameter_circular.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_column_quantity = " & Me.pile_column_quantity.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_row_quantity = " & Me.pile_row_quantity.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_columns_spacing = " & Me.pile_columns_spacing.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_row_spacing = " & Me.pile_row_spacing.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("group_efficiency_factor_given = " & Me.group_efficiency_factor_given.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("group_efficiency_factor = " & Me.group_efficiency_factor.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("cap_type = " & Me.cap_type.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_quantity_asymmetric = " & Me.pile_quantity_asymmetric.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_spacing_min_asymmetric = " & Me.pile_spacing_min_asymmetric.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("quantity_piles_surrounding = " & Me.quantity_piles_surrounding.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pile_cap_reference = " & Me.pile_cap_reference.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("Soil_110 = " & Me.Soil_110.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("Structural_105 = " & Me.Structural_105.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("tool_version = " & Me.tool_version.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        'SQLUpdate = SQLUpdate.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)

        Return SQLUpdate
    End Function

#End Region

#Region "Equals"
    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)
        Dim categoryName As String = Me.EDSObjectFullName

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As Pile = TryCast(other, Pile)
        If otherToCompare Is Nothing Then Return False

        'Equals = If(Me.ID.CheckChange(otherToCompare.ID, changes, categoryName, "Id"), Equals, False)
        'Equals = If(Me.bus_unit.CheckChange(otherToCompare.bus_unit, changes, categoryName, "Bus Unit"), Equals, False)
        'Equals = If(Me.structure_id.CheckChange(otherToCompare.structure_id, changes, categoryName, "Structure Id"), Equals, False)
        Equals = If(Me.load_eccentricity.CheckChange(otherToCompare.load_eccentricity, changes, categoryName, "Load Eccentricity"), Equals, False)
        Equals = If(Me.bolt_circle_bearing_plate_width.CheckChange(otherToCompare.bolt_circle_bearing_plate_width, changes, categoryName, "Bolt Circle Bearing Plate Width"), Equals, False)
        Equals = If(Me.pile_shape.CheckChange(otherToCompare.pile_shape, changes, categoryName, "Pile Shape"), Equals, False)
        Equals = If(Me.pile_material.CheckChange(otherToCompare.pile_material, changes, categoryName, "Pile Material"), Equals, False)
        Equals = If(Me.pile_length.CheckChange(otherToCompare.pile_length, changes, categoryName, "Pile Length"), Equals, False)
        Equals = If(Me.pile_diameter_width.CheckChange(otherToCompare.pile_diameter_width, changes, categoryName, "Pile Diameter Width"), Equals, False)
        Equals = If(Me.pile_pipe_thickness.CheckChange(otherToCompare.pile_pipe_thickness, changes, categoryName, "Pile Pipe Thickness"), Equals, False)
        Equals = If(Me.pile_soil_capacity_given.CheckChange(otherToCompare.pile_soil_capacity_given, changes, categoryName, "Pile Soil Capacity Given"), Equals, False)
        Equals = If(Me.steel_yield_strength.CheckChange(otherToCompare.steel_yield_strength, changes, categoryName, "Steel Yield Strength"), Equals, False)
        Equals = If(Me.pile_type_option.CheckChange(otherToCompare.pile_type_option, changes, categoryName, "Pile Type Option"), Equals, False)
        Equals = If(Me.rebar_quantity.CheckChange(otherToCompare.rebar_quantity, changes, categoryName, "Rebar Quantity"), Equals, False)
        Equals = If(Me.pile_group_config.CheckChange(otherToCompare.pile_group_config, changes, categoryName, "Pile Group Config"), Equals, False)
        Equals = If(Me.foundation_depth.CheckChange(otherToCompare.foundation_depth, changes, categoryName, "Foundation Depth"), Equals, False)
        Equals = If(Me.pad_thickness.CheckChange(otherToCompare.pad_thickness, changes, categoryName, "Pad Thickness"), Equals, False)
        Equals = If(Me.pad_width_dir1.CheckChange(otherToCompare.pad_width_dir1, changes, categoryName, "Pad Width Dir1"), Equals, False)
        Equals = If(Me.pad_width_dir2.CheckChange(otherToCompare.pad_width_dir2, changes, categoryName, "Pad Width Dir2"), Equals, False)
        Equals = If(Me.pad_rebar_size_bottom.CheckChange(otherToCompare.pad_rebar_size_bottom, changes, categoryName, "Pad Rebar Size Bottom"), Equals, False)
        Equals = If(Me.pad_rebar_size_top.CheckChange(otherToCompare.pad_rebar_size_top, changes, categoryName, "Pad Rebar Size Top"), Equals, False)
        Equals = If(Me.pad_rebar_quantity_bottom_dir1.CheckChange(otherToCompare.pad_rebar_quantity_bottom_dir1, changes, categoryName, "Pad Rebar Quantity Bottom Dir1"), Equals, False)
        Equals = If(Me.pad_rebar_quantity_top_dir1.CheckChange(otherToCompare.pad_rebar_quantity_top_dir1, changes, categoryName, "Pad Rebar Quantity Top Dir1"), Equals, False)
        Equals = If(Me.pad_rebar_quantity_bottom_dir2.CheckChange(otherToCompare.pad_rebar_quantity_bottom_dir2, changes, categoryName, "Pad Rebar Quantity Bottom Dir2"), Equals, False)
        Equals = If(Me.pad_rebar_quantity_top_dir2.CheckChange(otherToCompare.pad_rebar_quantity_top_dir2, changes, categoryName, "Pad Rebar Quantity Top Dir2"), Equals, False)
        Equals = If(Me.pier_shape.CheckChange(otherToCompare.pier_shape, changes, categoryName, "Pier Shape"), Equals, False)
        Equals = If(Me.pier_diameter.CheckChange(otherToCompare.pier_diameter, changes, categoryName, "Pier Diameter"), Equals, False)
        Equals = If(Me.extension_above_grade.CheckChange(otherToCompare.extension_above_grade, changes, categoryName, "Extension Above Grade"), Equals, False)
        Equals = If(Me.pier_rebar_size.CheckChange(otherToCompare.pier_rebar_size, changes, categoryName, "Pier Rebar Size"), Equals, False)
        Equals = If(Me.pier_rebar_quantity.CheckChange(otherToCompare.pier_rebar_quantity, changes, categoryName, "Pier Rebar Quantity"), Equals, False)
        Equals = If(Me.pier_tie_size.CheckChange(otherToCompare.pier_tie_size, changes, categoryName, "Pier Tie Size"), Equals, False)
        Equals = If(Me.rebar_grade.CheckChange(otherToCompare.rebar_grade, changes, categoryName, "Rebar Grade"), Equals, False)
        Equals = If(Me.concrete_compressive_strength.CheckChange(otherToCompare.concrete_compressive_strength, changes, categoryName, "Concrete Compressive Strength"), Equals, False)
        Equals = If(Me.groundwater_depth.CheckChange(otherToCompare.groundwater_depth, changes, categoryName, "Groundwater Depth"), Equals, False)
        Equals = If(Me.total_soil_unit_weight.CheckChange(otherToCompare.total_soil_unit_weight, changes, categoryName, "Total Soil Unit Weight"), Equals, False)
        Equals = If(Me.cohesion.CheckChange(otherToCompare.cohesion, changes, categoryName, "Cohesion"), Equals, False)
        Equals = If(Me.friction_angle.CheckChange(otherToCompare.friction_angle, changes, categoryName, "Friction Angle"), Equals, False)
        Equals = If(Me.neglect_depth.CheckChange(otherToCompare.neglect_depth, changes, categoryName, "Neglect Depth"), Equals, False)
        Equals = If(Me.spt_blow_count.CheckChange(otherToCompare.spt_blow_count, changes, categoryName, "Spt Blow Count"), Equals, False)
        Equals = If(Me.pile_negative_friction_force.CheckChange(otherToCompare.pile_negative_friction_force, changes, categoryName, "Pile Negative Friction Force"), Equals, False)
        Equals = If(Me.pile_ultimate_compression.CheckChange(otherToCompare.pile_ultimate_compression, changes, categoryName, "Pile Ultimate Compression"), Equals, False)
        Equals = If(Me.pile_ultimate_tension.CheckChange(otherToCompare.pile_ultimate_tension, changes, categoryName, "Pile Ultimate Tension"), Equals, False)
        Equals = If(Me.top_and_bottom_rebar_different.CheckChange(otherToCompare.top_and_bottom_rebar_different, changes, categoryName, "Top And Bottom Rebar Different"), Equals, False)
        Equals = If(Me.ultimate_gross_end_bearing.CheckChange(otherToCompare.ultimate_gross_end_bearing, changes, categoryName, "Ultimate Gross End Bearing"), Equals, False)
        Equals = If(Me.skin_friction_given.CheckChange(otherToCompare.skin_friction_given, changes, categoryName, "Skin Friction Given"), Equals, False)
        Equals = If(Me.pile_quantity_circular.CheckChange(otherToCompare.pile_quantity_circular, changes, categoryName, "Pile Quantity Circular"), Equals, False)
        Equals = If(Me.group_diameter_circular.CheckChange(otherToCompare.group_diameter_circular, changes, categoryName, "Group Diameter Circular"), Equals, False)
        Equals = If(Me.pile_column_quantity.CheckChange(otherToCompare.pile_column_quantity, changes, categoryName, "Pile Column Quantity"), Equals, False)
        Equals = If(Me.pile_row_quantity.CheckChange(otherToCompare.pile_row_quantity, changes, categoryName, "Pile Row Quantity"), Equals, False)
        Equals = If(Me.pile_columns_spacing.CheckChange(otherToCompare.pile_columns_spacing, changes, categoryName, "Pile Columns Spacing"), Equals, False)
        Equals = If(Me.pile_row_spacing.CheckChange(otherToCompare.pile_row_spacing, changes, categoryName, "Pile Row Spacing"), Equals, False)
        Equals = If(Me.group_efficiency_factor_given.CheckChange(otherToCompare.group_efficiency_factor_given, changes, categoryName, "Group Efficiency Factor Given"), Equals, False)
        Equals = If(Me.group_efficiency_factor.CheckChange(otherToCompare.group_efficiency_factor, changes, categoryName, "Group Efficiency Factor"), Equals, False)
        Equals = If(Me.cap_type.CheckChange(otherToCompare.cap_type, changes, categoryName, "Cap Type"), Equals, False)
        Equals = If(Me.pile_quantity_asymmetric.CheckChange(otherToCompare.pile_quantity_asymmetric, changes, categoryName, "Pile Quantity Asymmetric"), Equals, False)
        Equals = If(Me.pile_spacing_min_asymmetric.CheckChange(otherToCompare.pile_spacing_min_asymmetric, changes, categoryName, "Pile Spacing Min Asymmetric"), Equals, False)
        Equals = If(Me.quantity_piles_surrounding.CheckChange(otherToCompare.quantity_piles_surrounding, changes, categoryName, "Quantity Piles Surrounding"), Equals, False)
        Equals = If(Me.pile_cap_reference.CheckChange(otherToCompare.pile_cap_reference, changes, categoryName, "Pile Cap Reference"), Equals, False)
        Equals = If(Me.Soil_110.CheckChange(otherToCompare.Soil_110, changes, categoryName, "Soil 110"), Equals, False)
        Equals = If(Me.Structural_105.CheckChange(otherToCompare.Structural_105, changes, categoryName, "Structural 105"), Equals, False)
        Equals = If(Me.tool_version.CheckChange(otherToCompare.tool_version, changes, categoryName, "Tool Version"), Equals, False)
        'Equals = If(Me.modified_person_id.CheckChange(otherToCompare.modified_person_id, changes, categoryName, "Modified Person Id"), Equals, False)
        'Equals = If(Me.process_stage.CheckChange(otherToCompare.process_stage, changes, categoryName, "Process Stage"), Equals, False)


        Return Equals

    End Function
#End Region

End Class