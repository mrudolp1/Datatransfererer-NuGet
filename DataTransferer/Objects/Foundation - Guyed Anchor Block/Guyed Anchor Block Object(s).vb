Option Strict On
Imports System.ComponentModel

Partial Public Class GuyedAnchorBlock

#Region "Define"

    Private prop_anchor_id As Integer
    Private prop_local_anchor_profile As Integer?
    'Private prop_anchor_location As String
    'Private prop_guy_anchor_radius As Double
    Private prop_anchor_depth As Double?
    Private prop_anchor_width As Double?
    Private prop_anchor_thickness As Double?
    Private prop_anchor_length As Double?
    Private prop_anchor_toe_width As Double?
    Private prop_anchor_top_rebar_size As Integer?
    Private prop_anchor_top_rebar_quantity As Integer?
    Private prop_anchor_front_rebar_size As Integer?
    Private prop_anchor_front_rebar_quantity As Integer?
    Private prop_anchor_stirrup_size As Integer?
    Private prop_anchor_shaft_diameter As Double?
    Private prop_anchor_shaft_quantity As Integer?
    Private prop_anchor_shaft_area_override As Double?
    Private prop_anchor_shaft_shear_lag_factor As Double?
    Private prop_anchor_shaft_section_type As String
    Private prop_anchor_rebar_grade As Double?
    Private prop_concrete_compressive_strength As Double?
    Private prop_clear_cover As Double?
    Private prop_anchor_shaft_yield_strength As Double?
    Private prop_anchor_shaft_ultimate_strength As Double?
    Private prop_neglect_depth As Double?
    Private prop_groundwater_depth As Double?
    Private prop_soil_layer_quantity As Integer?
    Private prop_rebar_known As Boolean
    Private prop_anchor_shaft_known As Boolean
    Private prop_basic_soil_check As Boolean
    Private prop_structural_check As Boolean
    Private prop_tool_version As String
    Private prop_local_anchor_id As Integer?
    Private prop_foundation_id As Integer?
    Private prop_ID As Integer
    Public Property soil_layers As New List(Of GuyedAnchorBlockSoilLayer)
    Public Property anchor_profiles As New List(Of GuyedAnchorBlockProfile)

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Global ID")>
    Public Property anchor_id() As Integer
        Get
            Return Me.prop_anchor_id
        End Get
        Set
            Me.prop_anchor_id = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Local Anchor ID")>
    Public Property local_anchor_id() As Integer?
        Get
            Return Me.prop_local_anchor_id
        End Get
        Set
            Me.prop_local_anchor_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Local Profile ID")>
    Public Property local_anchor_profile() As Integer?
        Get
            Return Me.prop_local_anchor_profile
        End Get
        Set
            Me.prop_local_anchor_profile = Value
        End Set
    End Property

    '<Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Location")>
    'Public Property anchor_location() As String
    '    Get
    '        Return Me.prop_anchor_location
    '    End Get
    '    Set
    '        Me.prop_anchor_location = Value
    '    End Set
    'End Property

    '<Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Radius")>
    'Public Property guy_anchor_radius() As Double
    '    Get
    '        Return Me.prop_guy_anchor_radius
    '    End Get
    '    Set
    '        Me.prop_guy_anchor_radius = Value
    '    End Set
    'End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Depth")>
    Public Property anchor_depth() As Double?
        Get
            Return Me.prop_anchor_depth
        End Get
        Set
            Me.prop_anchor_depth = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Width")>
    Public Property anchor_width() As Double?
        Get
            Return Me.prop_anchor_width
        End Get
        Set
            Me.prop_anchor_width = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Thickness")>
    Public Property anchor_thickness() As Double?
        Get
            Return Me.prop_anchor_thickness
        End Get
        Set
            Me.prop_anchor_thickness = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Length")>
    Public Property anchor_length() As Double?
        Get
            Return Me.prop_anchor_length
        End Get
        Set
            Me.prop_anchor_length = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Toe Width")>
    Public Property anchor_toe_width() As Double?
        Get
            Return Me.prop_anchor_toe_width
        End Get
        Set
            Me.prop_anchor_toe_width = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Top Rebar Size")>
    Public Property anchor_top_rebar_size() As Integer?
        Get
            Return Me.prop_anchor_top_rebar_size
        End Get
        Set
            Me.prop_anchor_top_rebar_size = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Top Rebar Quantity")>
    Public Property anchor_top_rebar_quantity() As Integer?
        Get
            Return Me.prop_anchor_top_rebar_quantity
        End Get
        Set
            Me.prop_anchor_top_rebar_quantity = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Front Rebar Size")>
    Public Property anchor_front_rebar_size() As Integer?
        Get
            Return Me.prop_anchor_front_rebar_size
        End Get
        Set
            Me.prop_anchor_front_rebar_size = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Front Rebar Quantity")>
    Public Property anchor_front_rebar_quantity() As Integer?
        Get
            Return Me.prop_anchor_front_rebar_quantity
        End Get
        Set
            Me.prop_anchor_front_rebar_quantity = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Stirrup Size")>
    Public Property anchor_stirrup_size() As Integer?
        Get
            Return Me.prop_anchor_stirrup_size
        End Get
        Set
            Me.prop_anchor_stirrup_size = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Diameter")>
    Public Property anchor_shaft_diameter() As Double?
        Get
            Return Me.prop_anchor_shaft_diameter
        End Get
        Set
            Me.prop_anchor_shaft_diameter = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Quantity")>
    Public Property anchor_shaft_quantity() As Integer?
        Get
            Return Me.prop_anchor_shaft_quantity
        End Get
        Set
            Me.prop_anchor_shaft_quantity = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Area Override")>
    Public Property anchor_shaft_area_override() As Double?
        Get
            Return Me.prop_anchor_shaft_area_override
        End Get
        Set
            Me.prop_anchor_shaft_area_override = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Shear Lag Factor")>
    Public Property anchor_shaft_shear_lag_factor() As Double?
        Get
            Return Me.prop_anchor_shaft_shear_lag_factor
        End Get
        Set
            Me.prop_anchor_shaft_shear_lag_factor = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Section Type")>
    Public Property anchor_shaft_section_type() As String
        Get
            Return Me.prop_anchor_shaft_section_type
        End Get
        Set
            Me.prop_anchor_shaft_section_type = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Rebar Grade")>
    Public Property anchor_rebar_grade() As Double?
        Get
            Return Me.prop_anchor_rebar_grade
        End Get
        Set
            Me.prop_anchor_rebar_grade = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Concrete Compressive Strength")>
    Public Property concrete_compressive_strength() As Double?
        Get
            Return Me.prop_concrete_compressive_strength
        End Get
        Set
            Me.prop_concrete_compressive_strength = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Clear Cover")>
    Public Property clear_cover() As Double?
        Get
            Return Me.prop_clear_cover
        End Get
        Set
            Me.prop_clear_cover = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Yield Strength")>
    Public Property anchor_shaft_yield_strength() As Double?
        Get
            Return Me.prop_anchor_shaft_yield_strength
        End Get
        Set
            Me.prop_anchor_shaft_yield_strength = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Ultimate Strength")>
    Public Property anchor_shaft_ultimate_strength() As Double?
        Get
            Return Me.prop_anchor_shaft_ultimate_strength
        End Get
        Set
            Me.prop_anchor_shaft_ultimate_strength = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Neglect Depth")>
    Public Property neglect_depth() As Double?
        Get
            Return Me.prop_neglect_depth
        End Get
        Set
            Me.prop_neglect_depth = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Groundwater Depth")>
    Public Property groundwater_depth() As Double?
        Get
            Return Me.prop_groundwater_depth
        End Get
        Set
            Me.prop_groundwater_depth = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Soil Layer Quantity")>
    Public Property soil_layer_quantity() As Integer?
        Get
            Return Me.prop_soil_layer_quantity
        End Get
        Set
            Me.prop_soil_layer_quantity = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Rebar Known")>
    Public Property rebar_known() As Boolean
        Get
            Return Me.prop_rebar_known
        End Get
        Set
            Me.prop_rebar_known = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Known")>
    Public Property anchor_shaft_known() As Boolean
        Get
            Return Me.prop_anchor_shaft_known
        End Get
        Set
            Me.prop_anchor_shaft_known = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Basic Soil Interactions up to 110%?")>
    Public Property basic_soil_check() As Boolean
        Get
            Return Me.prop_basic_soil_check
        End Get
        Set
            Me.prop_basic_soil_check = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Structural Checks up to 105%?")>
    Public Property structural_check() As Boolean
        Get
            Return Me.prop_structural_check
        End Get
        Set
            Me.prop_structural_check = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Tool Version")>
    Public Property tool_version() As String
        Get
            Return Me.prop_tool_version
        End Get
        Set
            Me.prop_tool_version = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Foundation ID")>
    Public Property foundation_id() As Integer?
        Get
            Return Me.prop_foundation_id
        End Get
        Set
            Me.prop_foundation_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("ID")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property

#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal GuyedAnchorBlockDataRow As DataRow, refID As Integer)
        'General Guyed Anchor Block Details
        Try
            Me.anchor_id = CType(GuyedAnchorBlockDataRow.Item("anchor_id"), Integer)
        Catch
            Me.anchor_id = 0
        End Try 'Guyed Anchor Block ID
        'Try
        '    Me.anchor_location = CType(GuyedAnchorBlockDataRow.Item("anchor_location"), String)
        'Catch
        '    Me.anchor_location = ""
        'End Try 'Guyed Anchor Block Location
        'Try
        '    Me.guy_anchor_radius = CType(GuyedAnchorBlockDataRow.Item("guy_anchor_radius"), Double)
        'Catch
        '    Me.guy_anchor_radius = -1 'Set stored value of -1 to "" (Empty) in tool
        'End Try 'Guyed Anchor Block Radius
        Try
            If Not IsDBNull(Me.local_anchor_profile = CType(GuyedAnchorBlockDataRow.Item("local_anchor_profile"), Integer)) Then
                Me.local_anchor_profile = CType(GuyedAnchorBlockDataRow.Item("local_anchor_profile"), Integer)
            Else
                Me.local_anchor_profile = Nothing
            End If
        Catch
            Me.local_anchor_profile = Nothing
        End Try 'Local Anchor Profile
        Try
            If Not IsDBNull(Me.local_anchor_profile = CType(GuyedAnchorBlockDataRow.Item("local_anchor_id"), Integer)) Then
                Me.local_anchor_id = CType(GuyedAnchorBlockDataRow.Item("local_anchor_id"), Integer)
            Else
                Me.local_anchor_id = Nothing
            End If
        Catch
            Me.local_anchor_id = Nothing
        End Try 'Local Anchor ID
        Try
            If Not IsDBNull(Me.anchor_depth = CType(GuyedAnchorBlockDataRow.Item("anchor_depth"), Double)) Then
                Me.anchor_depth = CType(GuyedAnchorBlockDataRow.Item("anchor_depth"), Double)
            Else
                Me.anchor_depth = Nothing
            End If
        Catch
            Me.anchor_depth = Nothing
        End Try 'Guyed Anchor Block Depth
        Try
            If Not IsDBNull(Me.anchor_width = CType(GuyedAnchorBlockDataRow.Item("anchor_width"), Double)) Then
                Me.anchor_width = CType(GuyedAnchorBlockDataRow.Item("anchor_width"), Double)
            Else
                Me.anchor_width = Nothing
            End If
        Catch
            Me.anchor_width = Nothing
        End Try 'Guyed Anchor Block Width
        Try
            If Not IsDBNull(Me.anchor_thickness = CType(GuyedAnchorBlockDataRow.Item("anchor_thickness"), Double)) Then
                Me.anchor_thickness = CType(GuyedAnchorBlockDataRow.Item("anchor_thickness"), Double)
            Else
                Me.anchor_thickness = Nothing
            End If
        Catch
            Me.anchor_thickness = Nothing
        End Try 'Guyed Anchor Block Thickness
        Try
            If Not IsDBNull(Me.anchor_length = CType(GuyedAnchorBlockDataRow.Item("anchor_length"), Double)) Then
                Me.anchor_length = CType(GuyedAnchorBlockDataRow.Item("anchor_length"), Double)
            Else
                Me.anchor_length = Nothing
            End If
        Catch
            Me.anchor_length = Nothing
        End Try 'Guyed Anchor Block Length
        Try
            If Not IsDBNull(Me.anchor_toe_width = CType(GuyedAnchorBlockDataRow.Item("anchor_toe_width"), Double)) Then
                Me.anchor_toe_width = CType(GuyedAnchorBlockDataRow.Item("anchor_toe_width"), Double)
            Else
                Me.anchor_toe_width = Nothing
            End If
        Catch
            Me.anchor_toe_width = Nothing
        End Try 'Guyed Anchor Block Toe Width
        Try
            If Not IsDBNull(Me.anchor_top_rebar_size = CType(GuyedAnchorBlockDataRow.Item("anchor_top_rebar_size"), Integer)) Then
                Me.anchor_top_rebar_size = CType(GuyedAnchorBlockDataRow.Item("anchor_top_rebar_size"), Integer)
            Else
                Me.anchor_top_rebar_size = Nothing
            End If
        Catch
            Me.anchor_top_rebar_size = Nothing
        End Try 'Guyed Anchor Block Top Rebar Size
        Try
            If Not IsDBNull(Me.anchor_top_rebar_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_top_rebar_quantity"), Integer)) Then
                Me.anchor_top_rebar_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_top_rebar_quantity"), Integer)
            Else
                Me.anchor_top_rebar_quantity = Nothing
            End If
        Catch
            Me.anchor_top_rebar_quantity = Nothing
        End Try 'Guyed Anchor Block Top Rebar Quantity
        Try
            If Not IsDBNull(Me.anchor_front_rebar_size = CType(GuyedAnchorBlockDataRow.Item("anchor_front_rebar_size"), Integer)) Then
                Me.anchor_front_rebar_size = CType(GuyedAnchorBlockDataRow.Item("anchor_front_rebar_size"), Integer)
            Else
                Me.anchor_front_rebar_size = Nothing
            End If
        Catch
            Me.anchor_front_rebar_size = Nothing
        End Try 'Guyed Anchor Block Front Rebar Size
        Try
            If Not IsDBNull(Me.anchor_front_rebar_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_front_rebar_quantity"), Integer)) Then
                Me.anchor_front_rebar_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_front_rebar_quantity"), Integer)
            Else
                Me.anchor_front_rebar_quantity = Nothing
            End If
        Catch
            Me.anchor_front_rebar_quantity = Nothing
        End Try 'Guyed Anchor Block Front Rebar Quantity
        Try
            If Not IsDBNull(Me.anchor_stirrup_size = CType(GuyedAnchorBlockDataRow.Item("anchor_stirrup_size"), Integer)) Then
                Me.anchor_stirrup_size = CType(GuyedAnchorBlockDataRow.Item("anchor_stirrup_size"), Integer)
            Else
                Me.anchor_stirrup_size = Nothing
            End If
        Catch
            Me.anchor_stirrup_size = Nothing
        End Try 'Guyed Anchor Block Stirrup Size
        Try
            If Not IsDBNull(Me.anchor_shaft_diameter = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_diameter"), Double)) Then
                Me.anchor_shaft_diameter = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_diameter"), Double)
            Else
                Me.anchor_shaft_diameter = Nothing
            End If
        Catch
            Me.anchor_shaft_diameter = Nothing
        End Try 'Guyed Anchor Block Anchor Shaft Diameter
        Try
            If Not IsDBNull(Me.anchor_shaft_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_quantity"), Integer)) Then
                Me.anchor_shaft_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_quantity"), Integer)
            Else
                Me.anchor_shaft_quantity = Nothing
            End If
        Catch
            Me.anchor_shaft_quantity = Nothing
        End Try 'Guyed Anchor Block Anchor Shaft Quantity
        Try
            If Not IsDBNull(Me.anchor_shaft_area_override = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_area_override"), Double)) Then
                Me.anchor_shaft_area_override = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_area_override"), Double)
            Else
                Me.anchor_shaft_area_override = Nothing
            End If
        Catch
            Me.anchor_shaft_area_override = Nothing
        End Try 'Guyed Anchor Block Anchor Area Override
        Try
            If Not IsDBNull(Me.anchor_shaft_shear_lag_factor = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_shear_lag_factor"), Double)) Then
                Me.anchor_shaft_shear_lag_factor = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_shear_lag_factor"), Double)
            Else
                Me.anchor_shaft_shear_lag_factor = Nothing
            End If
        Catch
            Me.anchor_shaft_shear_lag_factor = Nothing
        End Try 'Guyed Anchor Block Anchor Shear Lag Factor
        Try
            If Not IsDBNull(Me.anchor_shaft_section_type = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_section_type"), String)) Then
                Me.anchor_shaft_section_type = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_section_type"), String)
            Else
                Me.anchor_shaft_section_type = Nothing
            End If
        Catch
            Me.anchor_shaft_section_type = Nothing
        End Try 'Guyed Anchor Block Anchor Section Type
        Try
            If Not IsDBNull(Me.anchor_rebar_grade = CType(GuyedAnchorBlockDataRow.Item("anchor_rebar_grade"), Double)) Then
                Me.anchor_rebar_grade = CType(GuyedAnchorBlockDataRow.Item("anchor_rebar_grade"), Double)
            Else
                Me.anchor_rebar_grade = Nothing
            End If
        Catch
            Me.anchor_rebar_grade = Nothing
        End Try 'Guyed Anchor Block Rebar Grade
        Try
            If Not IsDBNull(Me.concrete_compressive_strength = CType(GuyedAnchorBlockDataRow.Item("concrete_compressive_strength"), Double)) Then
                Me.concrete_compressive_strength = CType(GuyedAnchorBlockDataRow.Item("concrete_compressive_strength"), Double)
            Else
                Me.concrete_compressive_strength = Nothing
            End If
        Catch
            Me.concrete_compressive_strength = Nothing
        End Try 'Guyed Anchor Block Concrete Compressive Strength
        Try
            If Not IsDBNull(Me.clear_cover = CType(GuyedAnchorBlockDataRow.Item("clear_cover"), Double)) Then
                Me.clear_cover = CType(GuyedAnchorBlockDataRow.Item("clear_cover"), Double)
            Else
                Me.clear_cover = Nothing
            End If
        Catch
            Me.clear_cover = Nothing
        End Try 'Guyed Anchor Block Clear Cover
        Try
            If Not IsDBNull(Me.anchor_shaft_yield_strength = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_yield_strength"), Double)) Then
                Me.anchor_shaft_yield_strength = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_yield_strength"), Double)
            Else
                Me.anchor_shaft_yield_strength = Nothing
            End If
        Catch
            Me.anchor_shaft_yield_strength = Nothing
        End Try 'Guyed Anchor Block Anchor Shaft Yield Strength
        Try
            If Not IsDBNull(Me.anchor_shaft_ultimate_strength = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_ultimate_strength"), Double)) Then
                Me.anchor_shaft_ultimate_strength = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_ultimate_strength"), Double)
            Else
                Me.anchor_shaft_ultimate_strength = Nothing
            End If
        Catch
            Me.anchor_shaft_ultimate_strength = Nothing
        End Try 'Guyed Anchor Block Anchor Shaft Ultimate Strength
        Try
            If Not IsDBNull(Me.neglect_depth = CType(GuyedAnchorBlockDataRow.Item("neglect_depth"), Double)) Then
                Me.neglect_depth = CType(GuyedAnchorBlockDataRow.Item("neglect_depth"), Double)
            Else
                Me.neglect_depth = Nothing
            End If
        Catch
            Me.neglect_depth = Nothing
        End Try 'Guyed Anchor Block Anchor Neglect Depth
        Try
            If Not IsDBNull(Me.groundwater_depth = CType(GuyedAnchorBlockDataRow.Item("groundwater_depth"), Double)) Then
                Me.groundwater_depth = CType(GuyedAnchorBlockDataRow.Item("groundwater_depth"), Double)
            Else
                Me.groundwater_depth = Nothing
            End If
        Catch
            Me.groundwater_depth = -1
        End Try 'Guyed Anchor Block Anchor Groundwater Depth
        Try
            If Not IsDBNull(Me.soil_layer_quantity = CType(GuyedAnchorBlockDataRow.Item("soil_layer_quantity"), Integer)) Then
                Me.soil_layer_quantity = CType(GuyedAnchorBlockDataRow.Item("soil_layer_quantity"), Integer)
            Else
                Me.soil_layer_quantity = Nothing
            End If
        Catch
            Me.soil_layer_quantity = Nothing
        End Try 'Guyed Anchor Block Anchor Soil Layer Quantity
        Try
            If Not IsDBNull(Me.rebar_known = CType(GuyedAnchorBlockDataRow.Item("rebar_known"), Boolean)) Then
                Me.rebar_known = CType(GuyedAnchorBlockDataRow.Item("rebar_known"), Boolean)
            Else
                Me.rebar_known = Nothing
            End If
        Catch
            Me.rebar_known = Nothing
        End Try 'Rebar Known
        Try
            If Not IsDBNull(Me.anchor_shaft_known = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_known"), Boolean)) Then
                Me.anchor_shaft_known = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_known"), Boolean)
            Else
                Me.anchor_shaft_known = Nothing
            End If
        Catch
            Me.anchor_shaft_known = Nothing
        End Try 'Anchor Shaft Known
        Try
            If Not IsDBNull(Me.basic_soil_check = CType(GuyedAnchorBlockDataRow.Item("basic_soil_check"), Boolean)) Then
                Me.basic_soil_check = CType(GuyedAnchorBlockDataRow.Item("basic_soil_check"), Boolean)
            Else
                Me.basic_soil_check = Nothing
            End If
        Catch
            Me.basic_soil_check = Nothing
        End Try 'Basic Soil Interaction up to 110%
        Try
            If Not IsDBNull(Me.structural_check = CType(GuyedAnchorBlockDataRow.Item("structural_check"), Boolean)) Then
                Me.structural_check = CType(GuyedAnchorBlockDataRow.Item("structural_check"), Boolean)
            Else
                Me.structural_check = Nothing
            End If
        Catch
            Me.structural_check = Nothing
        End Try 'Basic Structural Check up to 105%
        Try
            If Not IsDBNull(Me.tool_version = CType(GuyedAnchorBlockDataRow.Item("tool_version"), String)) Then
                Me.tool_version = CType(GuyedAnchorBlockDataRow.Item("tool_version"), String)
            Else
                Me.tool_version = Nothing
            End If
        Catch
            Me.tool_version = Nothing
        End Try 'Tool Version
        Try
            If Not IsDBNull(Me.foundation_id = CType(GuyedAnchorBlockDataRow.Item("foundation_id"), Integer)) Then
                Me.foundation_id = CType(GuyedAnchorBlockDataRow.Item("foundation_id"), Integer)
            Else
                Me.foundation_id = Nothing
            End If
        Catch
            Me.foundation_id = Nothing
        End Try 'Foundation ID
        Try
            If Not IsDBNull(Me.ID = CType(GuyedAnchorBlockDataRow.Item("ID"), Integer)) Then
                Me.ID = CType(GuyedAnchorBlockDataRow.Item("ID"), Integer)
            Else
                Me.ID = Nothing
            End If
        Catch
            Me.ID = Nothing
        End Try 'ID

        For Each SoilLayerDataRow As DataRow In ds.Tables("Guyed Anchor Block Soil SQL").Rows
            Dim soilRefID As Integer = CType(SoilLayerDataRow.Item("anchor_id"), Integer)

            If soilRefID = refID Then
                Me.soil_layers.Add(New GuyedAnchorBlockSoilLayer(SoilLayerDataRow))
            End If
        Next 'Add Soil Layers to to Guyed Anchor Block Object

        For Each GuyedAnchorBlockProfileDataRow As DataRow In ds.Tables("Guyed Anchor Block Profiles SQL").Rows
            Dim profileRefID As Integer = CType(GuyedAnchorBlockProfileDataRow.Item("anchor_id"), Integer)

            If profileRefID = refID Then
                Me.anchor_profiles.Add(New GuyedAnchorBlockProfile(GuyedAnchorBlockProfileDataRow))
            End If
        Next 'Add Associated Profiles to Guyed Anchor Block Object

    End Sub 'Generate a guyed anchor block from EDS

    'Public Sub New(ByVal GuyedAnchorBlockDataRow As DataRow, ByVal refID As Integer, ByVal refcol As String, ByVal constants As List(Of EXCELRngParameter))
    Public Sub New(ByVal GuyedAnchorBlockDataRow As DataRow, ByVal refID As Integer, ByVal refcol As String)
        'General Guyed Anchor Block Details
        Try
            Me.anchor_id = CType(GuyedAnchorBlockDataRow.Item("anchor_id"), Integer)
        Catch
            Me.anchor_id = 0
        End Try 'Guyed Anchor Block ID
        'Try
        '    Me.anchor_location = CType(GuyedAnchorBlockDataRow.Item("anchor_location"), String)
        'Catch
        '    Me.anchor_location = ""
        'End Try 'Guyed Anchor Block Location
        'Try
        '    Me.guy_anchor_radius = CType(GuyedAnchorBlockDataRow.Item("guy_anchor_radius"), Double)
        'Catch
        '    Me.guy_anchor_radius = -1 'Set stored value of -1 to "" (Empty) in tool
        'End Try 'Guyed Anchor Block Radius
        Try
            Me.local_anchor_profile = CType(GuyedAnchorBlockDataRow.Item("local_anchor_profile"), Integer)
        Catch
            Me.local_anchor_profile = Nothing
        End Try 'Local Anchor Profile
        Try
            Me.local_anchor_id = CType(GuyedAnchorBlockDataRow.Item("local_anchor_id"), Integer)
        Catch
            Me.local_anchor_id = Nothing
        End Try 'Local Anchor ID
        Try
            Me.anchor_depth = CType(GuyedAnchorBlockDataRow.Item("anchor_depth"), Double)
        Catch
            Me.anchor_depth = Nothing
        End Try 'Guyed Anchor Block Depth
        Try
            Me.anchor_width = CType(GuyedAnchorBlockDataRow.Item("anchor_width"), Double)
        Catch
            Me.anchor_width = Nothing
        End Try 'Guyed Anchor Block Width
        Try
            Me.anchor_thickness = CType(GuyedAnchorBlockDataRow.Item("anchor_thickness"), Double)
        Catch
            Me.anchor_thickness = Nothing
        End Try 'Guyed Anchor Block Thickness
        Try
            Me.anchor_length = CType(GuyedAnchorBlockDataRow.Item("anchor_length"), Double)
        Catch
            Me.anchor_length = Nothing
        End Try 'Guyed Anchor Block Length
        Try
            Me.anchor_toe_width = CType(GuyedAnchorBlockDataRow.Item("anchor_toe_width"), Double)
        Catch
            Me.anchor_toe_width = Nothing
        End Try 'Guyed Anchor Block Toe Width
        Try
            Me.anchor_top_rebar_size = CType(GuyedAnchorBlockDataRow.Item("anchor_top_rebar_size"), Integer)
        Catch
            Me.anchor_top_rebar_size = Nothing
        End Try 'Guyed Anchor Block Top Rebar Size
        Try
            Me.anchor_top_rebar_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_top_rebar_quantity"), Integer)
        Catch
            Me.anchor_top_rebar_quantity = Nothing
        End Try 'Guyed Anchor Block Top Rebar Quantity
        Try
            Me.anchor_front_rebar_size = CType(GuyedAnchorBlockDataRow.Item("anchor_front_rebar_size"), Integer)
        Catch
            Me.anchor_front_rebar_size = Nothing
        End Try 'Guyed Anchor Block Front Rebar Size
        Try
            Me.anchor_front_rebar_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_front_rebar_quantity"), Integer)
        Catch
            Me.anchor_front_rebar_quantity = Nothing
        End Try 'Guyed Anchor Block Front Rebar Quantity
        Try
            Me.anchor_stirrup_size = CType(GuyedAnchorBlockDataRow.Item("anchor_stirrup_size"), Integer)
        Catch
            Me.anchor_stirrup_size = Nothing
        End Try 'Guyed Anchor Block Stirrup Size
        Try
            Me.anchor_shaft_diameter = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_diameter"), Double)
        Catch
            Me.anchor_shaft_diameter = Nothing
        End Try 'Guyed Anchor Block Anchor Shaft Diameter
        Try
            Me.anchor_shaft_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_quantity"), Integer)
        Catch
            Me.anchor_shaft_quantity = Nothing
        End Try 'Guyed Anchor Block Anchor Shaft Quantity
        Try
            Me.anchor_shaft_area_override = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_area_override"), Double)
        Catch
            Me.anchor_shaft_area_override = Nothing
        End Try 'Guyed Anchor Block Anchor Area Override
        Try
            Me.anchor_shaft_shear_lag_factor = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_shear_lag_factor"), Double)
        Catch
            Me.anchor_shaft_shear_lag_factor = Nothing
        End Try 'Guyed Anchor Block Anchor Shear Lag Factor
        Try
            Me.anchor_shaft_section_type = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_section_type"), String)
        Catch
            Me.anchor_shaft_section_type = Nothing
        End Try 'Guyed Anchor Block Anchor Section Type
        Try
            Me.anchor_rebar_grade = CType(GuyedAnchorBlockDataRow.Item("anchor_rebar_grade"), Double)
        Catch
            Me.anchor_rebar_grade = Nothing
        End Try 'Guyed Anchor Block Rebar Grade
        Try
            Me.concrete_compressive_strength = CType(GuyedAnchorBlockDataRow.Item("concrete_compressive_strength"), Double)
        Catch
            Me.concrete_compressive_strength = Nothing
        End Try 'Guyed Anchor Block Concrete Compressive Strength
        Try
            Me.clear_cover = CType(GuyedAnchorBlockDataRow.Item("clear_cover"), Double)
        Catch
            Me.clear_cover = Nothing
        End Try 'Guyed Anchor Block Clear Cover
        Try
            Me.anchor_shaft_yield_strength = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_yield_strength"), Double)
        Catch
            Me.anchor_shaft_yield_strength = Nothing
        End Try 'Guyed Anchor Block Anchor Shaft Yield Strength
        Try
            Me.anchor_shaft_ultimate_strength = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_ultimate_strength"), Double)
        Catch
            Me.anchor_shaft_ultimate_strength = Nothing
        End Try 'Guyed Anchor Block Anchor Shaft Ultimate Strength
        Try
            Me.neglect_depth = CType(GuyedAnchorBlockDataRow.Item("neglect_depth"), Double)
        Catch
            Me.neglect_depth = Nothing
        End Try 'Guyed Anchor Block Anchor Neglect Depth
        Try
            Me.groundwater_depth = CType(GuyedAnchorBlockDataRow.Item("groundwater_depth"), Double)
        Catch
            Me.groundwater_depth = -1 'Set stored value of -1 to "N/A" in tool
        End Try 'Guyed Anchor Block Anchor Groundwater Depth
        Try
            Me.soil_layer_quantity = CType(GuyedAnchorBlockDataRow.Item("soil_layer_quantity"), Integer)
        Catch
            Me.soil_layer_quantity = Nothing
        End Try 'Guyed Anchor Block Anchor Soil Layer Quantity
        Try
            Me.rebar_known = CType(GuyedAnchorBlockDataRow.Item("rebar_known"), Boolean)
        Catch
            Me.rebar_known = Nothing
        End Try 'Rebar Known
        Try
            Me.anchor_shaft_known = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_known"), Boolean)
        Catch
            Me.anchor_shaft_known = Nothing
        End Try 'Anchor Shaft Known
        Try
            Me.basic_soil_check = CType(GuyedAnchorBlockDataRow.Item("basic_soil_check"), Boolean)
        Catch
            Me.basic_soil_check = Nothing
        End Try 'Basic Soil Interaction up to 110%?
        Try
            Me.structural_check = CType(GuyedAnchorBlockDataRow.Item("structural_check"), Boolean)
        Catch
            Me.structural_check = Nothing
        End Try 'Structural Checks up to 105%?
        Try
            Me.tool_version = CType(GuyedAnchorBlockDataRow.Item("tool_version"), String)
        Catch
            Me.tool_version = Nothing
        End Try 'Tool Version
        Try
            Me.foundation_id = CType(GuyedAnchorBlockDataRow.Item("foundation_id"), Integer)
        Catch
            Me.foundation_id = Nothing
        End Try 'Foundation ID

        For Each SoilLayerDataRow As DataRow In ds.Tables("Guyed Anchor Block Soil EXCEL").Rows
            'Dim soilRefID As Integer = CType(SoilLayerDataRow.Item(refcol), Integer)
            'Dim soilRefID As Integer = CType(SoilLayerDataRow.Item("local_anchor_profile"), Integer)

            'If soilRefID = refID Then
            If CType(SoilLayerDataRow.Item("local_soil_profile"), Integer) = CType(GuyedAnchorBlockDataRow.Item("local_anchor_id"), Integer) Then
                Me.soil_layers.Add(New GuyedAnchorBlockSoilLayer(SoilLayerDataRow))
            End If
        Next 'Add Soil Layers to to Guyed Anchor Block Object

        For Each GuyedAnchorBlockProfileDataRow As DataRow In ds.Tables("Guyed Anchor Block Profiles EXCEL").Rows 'WIP
            Dim profileID As Integer?

            'Try
            '    If Not IsNothing(CType(GuyedAnchorBlockProfileDataRow.Item(refcol), Integer)) Then
            '        profileID = CType(GuyedAnchorBlockProfileDataRow.Item(refcol), Integer)
            '    Else
            '        profileID = Nothing
            '    End If
            'Catch
            '    profileID = Nothing
            'End Try 'Profile Reference ID

            'If profileID = refID Then
            If CType(GuyedAnchorBlockProfileDataRow.Item("local_anchor_id"), Integer) = CType(GuyedAnchorBlockDataRow.Item("local_anchor_id"), Integer) Then
                Me.anchor_profiles.Add(New GuyedAnchorBlockProfile(GuyedAnchorBlockProfileDataRow))
            End If
        Next 'Add Profiles to to Drilled Pier Object
    End Sub 'Generate a guyed anchor block from Excel

#End Region

End Class


#Region "Guyed Anchor Block Extras"

Partial Public Class GuyedAnchorBlockSoilLayer

    Private prop_anchor_id As Integer?
    Private prop_soil_layer_id As Integer
    Private prop_bottom_depth As Double?
    Private prop_effective_soil_density As Double?
    Private prop_cohesion As Double?
    Private prop_friction_angle As Double?
    Private prop_skin_friction_override_uplift As Double?
    Private prop_spt_blow_count As Integer?
    Private prop_local_soil_layer_id As Integer?
    Private prop_ID As Integer

    <Category("Guyed Anchor Block Soil Layers"), Description(""), DisplayName("Anchor ID")>
    Public Property anchor_id() As Integer?
        Get
            Return Me.prop_anchor_id
        End Get
        Set
            Me.prop_anchor_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil Layers"), Description(""), DisplayName("Soil Layer ID")>
    Public Property soil_layer_id() As Integer
        Get
            Return Me.prop_soil_layer_id
        End Get
        Set
            Me.prop_soil_layer_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil Depth"), Description(""), DisplayName("Soil Depth")>
    Public Property bottom_depth() As Double?
        Get
            Return Me.prop_bottom_depth
        End Get
        Set
            Me.prop_bottom_depth = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil Effective Soil Density"), Description(""), DisplayName("Effective Soil Density")>
    Public Property effective_soil_density() As Double?
        Get
            Return Me.prop_effective_soil_density
        End Get
        Set
            Me.prop_effective_soil_density = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil Cohesion"), Description(""), DisplayName("Soil Cohesion")>
    Public Property cohesion() As Double?
        Get
            Return Me.prop_cohesion
        End Get
        Set
            Me.prop_cohesion = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil Friction Angle"), Description(""), DisplayName("Soil Friction Angle")>
    Public Property friction_angle() As Double?
        Get
            Return Me.prop_friction_angle
        End Get
        Set
            Me.prop_friction_angle = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil Skin Friction Override"), Description(""), DisplayName("Soil Skin Friction Override")>
    Public Property skin_friction_override_uplift() As Double?
        Get
            Return Me.prop_skin_friction_override_uplift
        End Get
        Set
            Me.prop_skin_friction_override_uplift = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil SPT Blow Count"), Description(""), DisplayName("Soil SPT Blow Count")>
    Public Property spt_blow_count() As Integer?
        Get
            Return Me.prop_spt_blow_count
        End Get
        Set
            Me.prop_spt_blow_count = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil SPT Blow Count"), Description(""), DisplayName("Local  Soil Layer ID")>
    Public Property local_soil_layer_id() As Integer?
        Get
            Return Me.prop_local_soil_layer_id
        End Get
        Set
            Me.prop_local_soil_layer_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil SPT Blow Count"), Description(""), DisplayName("ID")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal SoilLayerDataRow As DataRow)
        Try
            Me.anchor_id = CType(SoilLayerDataRow.Item("anchor_id"), Integer)
        Catch
            Me.anchor_id = Nothing
        End Try 'Anchor ID
        Try
            Me.soil_layer_id = CType(SoilLayerDataRow.Item("soil_layer_id"), Integer)
        Catch
            Me.soil_layer_id = 0
        End Try 'Soil Layer ID
        Try
            Me.bottom_depth = CType(SoilLayerDataRow.Item("bottom_depth"), Double)
        Catch
            Me.bottom_depth = Nothing
        End Try 'Soil Layer Depth
        Try
            Me.effective_soil_density = CType(SoilLayerDataRow.Item("effective_soil_density"), Double)
        Catch
            Me.effective_soil_density = Nothing
        End Try 'Soil Layer Effective Density
        Try
            Me.cohesion = CType(SoilLayerDataRow.Item("cohesion"), Double)
        Catch
            Me.cohesion = Nothing
        End Try 'Soil Layer Cohesion
        Try
            Me.friction_angle = CType(SoilLayerDataRow.Item("friction_angle"), Double)
        Catch
            Me.friction_angle = Nothing
        End Try 'Soil Layer Friction Angle
        Try
            Me.skin_friction_override_uplift = CType(SoilLayerDataRow.Item("skin_friction_override_uplift"), Double)
        Catch
            Me.skin_friction_override_uplift = Nothing
        End Try 'Soil Layer Skin Friction Override
        Try
            Me.spt_blow_count = CType(SoilLayerDataRow.Item("spt_blow_count"), Integer)
        Catch
            Me.spt_blow_count = Nothing
        End Try 'Soil Layer Blow Count
        Try
            Me.local_soil_layer_id = CType(SoilLayerDataRow.Item("local_soil_layer_id"), Integer)
        Catch
            Me.local_soil_layer_id = Nothing
        End Try 'Soil Layer Blow Count
        Try
            Me.ID = CType(SoilLayerDataRow.Item("ID"), Integer)
        Catch
            Me.ID = Nothing
        End Try 'ID
    End Sub

End Class

Partial Public Class GuyedAnchorBlockProfile

    Private prop_local_anchor_id As Integer?
    'Private prop_reaction_position As Integer?
    Private prop_anchor_id As Integer?
    Private prop_profile_id As Integer
    Private prop_reaction_location As String
    Private prop_anchor_profile As Integer?
    Private prop_soil_profile As Integer?
    Private prop_ID As Integer

    <Category("Guyed Anchor Block Profiles"), Description(""), DisplayName("Local Anchor ID")>
    Public Property local_anchor_id() As Integer?
        Get
            Return Me.prop_local_anchor_id
        End Get
        Set
            Me.prop_local_anchor_id = Value
        End Set
    End Property
    '<Category("Guyed Anchor Block Profiles"), Description(""), DisplayName("Reaction Position in Tool")>
    'Public Property reaction_position() As Integer?
    '    Get
    '        Return Me.prop_reaction_position
    '    End Get
    '    Set
    '        Me.prop_reaction_position = Value
    '    End Set
    'End Property
    <Category("Guyed Anchor Block Profiles"), Description(""), DisplayName("Anchor ID")>
    Public Property anchor_id() As Integer?
        Get
            Return Me.prop_anchor_id
        End Get
        Set
            Me.prop_anchor_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profiles"), Description(""), DisplayName("Profile ID")>
    Public Property profile_id() As Integer
        Get
            Return Me.prop_profile_id
        End Get
        Set
            Me.prop_profile_id = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profiles"), Description(""), DisplayName("Reaction Location")>
    Public Property reaction_location() As String
        Get
            Return Me.prop_reaction_location
        End Get
        Set
            Me.prop_reaction_location = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profiles"), Description(""), DisplayName("Anchor Profile")>
    Public Property anchor_profile() As Integer?
        Get
            Return Me.prop_anchor_profile
        End Get
        Set
            Me.prop_anchor_profile = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profiles"), Description(""), DisplayName("Soil Profile")>
    Public Property soil_profile() As Integer?
        Get
            Return Me.prop_soil_profile
        End Get
        Set
            Me.prop_soil_profile = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Profiles"), Description(""), DisplayName("ID")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal GuyedAnchorBlockProfileDataRow As DataRow)
        Try
            Me.local_anchor_id = CType(GuyedAnchorBlockProfileDataRow.Item("local_anchor_id"), Integer)
        Catch
            Me.local_anchor_id = 0
        End Try 'Local Anchor ID
        'Try
        '    If Not IsDBNull(Me.reaction_position = CType(GuyedAnchorBlockProfileDataRow.Item("reaction_position"), Integer)) Then
        '        Me.reaction_position = CType(GuyedAnchorBlockProfileDataRow.Item("reaction_position"), Integer)
        '    Else
        '        Me.reaction_position = Nothing
        '    End If
        'Catch
        '    Me.reaction_position = Nothing
        'End Try 'Reaction Position
        Try
            If Not IsDBNull(Me.anchor_id = CType(GuyedAnchorBlockProfileDataRow.Item("anchor_id"), Integer)) Then
                Me.anchor_id = CType(GuyedAnchorBlockProfileDataRow.Item("anchor_id"), Integer)
            Else
                Me.anchor_id = Nothing
            End If
        Catch
            Me.anchor_id = Nothing
        End Try 'Anchor ID
        Try
            If Not IsDBNull(Me.profile_id = CType(GuyedAnchorBlockProfileDataRow.Item("profile_id"), Integer)) Then
                Me.profile_id = CType(GuyedAnchorBlockProfileDataRow.Item("profile_id"), Integer)
            Else
                Me.profile_id = Nothing
            End If
        Catch
            Me.profile_id = Nothing
        End Try 'Profile ID
        Try
            If Not IsDBNull(Me.reaction_location = CType(GuyedAnchorBlockProfileDataRow.Item("reaction_location"), String)) Then
                Me.reaction_location = CType(GuyedAnchorBlockProfileDataRow.Item("reaction_location"), String)
            Else
                Me.reaction_location = Nothing
            End If
        Catch
            Me.reaction_location = Nothing
        End Try 'Reaction Location
        Try
            If Not IsDBNull(Me.anchor_profile = CType(GuyedAnchorBlockProfileDataRow.Item("anchor_profile"), Integer)) Then
                Me.anchor_profile = CType(GuyedAnchorBlockProfileDataRow.Item("anchor_profile"), Integer)
            Else
                Me.anchor_profile = Nothing
            End If
        Catch
            Me.anchor_profile = Nothing
        End Try 'Anchor Profile
        Try
            If Not IsDBNull(Me.soil_profile = CType(GuyedAnchorBlockProfileDataRow.Item("soil_profile"), Integer)) Then
                Me.soil_profile = CType(GuyedAnchorBlockProfileDataRow.Item("soil_profile"), Integer)
            Else
                Me.soil_profile = Nothing
            End If
        Catch
            Me.soil_profile = Nothing
        End Try 'Soil Profile
        Try
            If Not IsDBNull(Me.ID = CType(GuyedAnchorBlockProfileDataRow.Item("ID"), Integer)) Then
                Me.ID = CType(GuyedAnchorBlockProfileDataRow.Item("ID"), Integer)
            Else
                Me.ID = Nothing
            End If
        Catch
            Me.ID = Nothing
        End Try 'ID
    End Sub

End Class

#End Region

