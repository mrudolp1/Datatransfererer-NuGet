Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet

Public Class SST_Unit_Base

#Region "Define"
    Private prop_unit_base_id As Integer?
    Private prop_extension_above_grade As Double? '
    Private prop_foundation_depth As Double? '
    Private prop_concrete_compressive_strength As Double? '
    Private prop_dry_concrete_density As Double? '
    Private prop_rebar_grade As Double? '
    Private prop_top_and_bottom_rebar_different As Boolean '
    Private prop_block_foundation As Boolean '
    Private prop_rectangular_foundation As Boolean '
    Private prop_base_plate_distance_above_foundation As Double? '
    Private prop_bolt_circle_bearing_plate_width As Double? '
    Private prop_tower_centroid_offset As Boolean '
    Private prop_basic_soil_check As Boolean '
    Private prop_structural_check As Boolean '

    Private prop_pier_shape As String '
    Private prop_pier_diameter As Double? '
    Private prop_pier_rebar_quantity As Integer?
    Private prop_pier_rebar_size As Integer? '
    Private prop_pier_tie_quantity As Integer?
    Private prop_pier_tie_size As Integer? '
    Private prop_pier_reinforcement_type As String '
    Private prop_pier_clear_cover As Double? '

    Private prop_pad_width_1 As Double? '
    Private prop_pad_width_2 As Double? '
    Private prop_pad_thickness As Double? '
    Private prop_pad_rebar_size_top_dir1 As Integer? '
    Private prop_pad_rebar_size_bottom_dir1 As Integer? '
    Private prop_pad_rebar_size_top_dir2 As Integer? '
    Private prop_pad_rebar_size_bottom_dir2 As Integer? '
    Private prop_pad_rebar_quantity_top_dir1 As Integer? '
    Private prop_pad_rebar_quantity_bottom_dir1 As Integer? '
    Private prop_pad_rebar_quantity_top_dir2 As Integer? '
    Private prop_pad_rebar_quantity_bottom_dir2 As Integer? '
    Private prop_pad_clear_cover As Double? '

    Private prop_total_soil_unit_weight As Double? '
    Private prop_bearing_type As String '
    Private prop_nominal_bearing_capacity As Double? '
    Private prop_cohesion As Double? '
    Private prop_friction_angle As Double? '
    Private prop_spt_blow_count As Integer? '
    Private prop_base_friction_factor As Double? '
    Private prop_neglect_depth As Double? '
    Private prop_bearing_distribution_type As Boolean '
    Private prop_groundwater_depth As Double? '

    Private prop_tool_version As String '
    'Private prop_modified As Boolean

    'Public Property ModifiedRanges As New List(Of ModifiedRange)

    <Category("Unit Base Details"), Description(""), DisplayName("Unit Base ID")>
    Public Property unit_base_id() As Integer?
        Get
            Return Me.prop_unit_base_id
        End Get
        Set
            Me.prop_unit_base_id = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Extension Above Grade")>
    Public Property extension_above_grade() As Double?
        Get
            Return Me.prop_extension_above_grade
        End Get
        Set
            Me.prop_extension_above_grade = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Foundation Depth")>
    Public Property foundation_depth() As Double?
        Get
            Return Me.prop_foundation_depth
        End Get
        Set
            Me.prop_foundation_depth = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Concrete Compressive Strength")>
    Public Property concrete_compressive_strength() As Double?
        Get
            Return Me.prop_concrete_compressive_strength
        End Get
        Set
            Me.prop_concrete_compressive_strength = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Dry Concrete Density")>
    Public Property dry_concrete_density() As Double?
        Get
            Return Me.prop_dry_concrete_density
        End Get
        Set
            Me.prop_dry_concrete_density = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Rebar Grade")>
    Public Property rebar_grade() As Double?
        Get
            Return Me.prop_rebar_grade
        End Get
        Set
            Me.prop_rebar_grade = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Top and Bottom Rebar Different")>
    Public Property top_and_bottom_rebar_different() As Boolean
        Get
            Return Me.prop_top_and_bottom_rebar_different
        End Get
        Set
            Me.prop_top_and_bottom_rebar_different = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Block Foundation")>
    Public Property block_foundation() As Boolean
        Get
            Return Me.prop_block_foundation
        End Get
        Set
            Me.prop_block_foundation = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Rectangular Foundation")>
    Public Property rectangular_foundation() As Boolean
        Get
            Return Me.prop_rectangular_foundation
        End Get
        Set
            Me.prop_rectangular_foundation = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Base Plate Distance Above Foundation")>
    Public Property base_plate_distance_above_foundation() As Double?
        Get
            Return Me.prop_base_plate_distance_above_foundation
        End Get
        Set
            Me.prop_base_plate_distance_above_foundation = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Bolt Circle Bearing Plate Width")>
    Public Property bolt_circle_bearing_plate_width() As Double?
        Get
            Return Me.prop_bolt_circle_bearing_plate_width
        End Get
        Set
            Me.prop_bolt_circle_bearing_plate_width = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Tower Centroid Offset")>
    Public Property tower_centroid_offset() As Boolean
        Get
            Return Me.prop_tower_centroid_offset
        End Get
        Set
            Me.prop_tower_centroid_offset = Value
        End Set
    End Property

    <Category("Unit Base Details"), Description(""), DisplayName("Pier Shape")>
    Public Property pier_shape() As String
        Get
            Return Me.prop_pier_shape
        End Get
        Set
            Me.prop_pier_shape = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pier Diameter")>
    Public Property pier_diameter() As Double?
        Get
            Return Me.prop_pier_diameter
        End Get
        Set
            Me.prop_pier_diameter = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pier Rebar Quantity")>
    Public Property pier_rebar_quantity() As Integer?
        Get
            Return Me.prop_pier_rebar_quantity
        End Get
        Set
            Me.prop_pier_rebar_quantity = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pier Rebar Size")>
    Public Property pier_rebar_size() As Integer?
        Get
            Return Me.prop_pier_rebar_size
        End Get
        Set
            Me.prop_pier_rebar_size = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pier Tie Quantity")>
    Public Property pier_tie_quantity() As Integer?
        Get
            Return Me.prop_pier_tie_quantity
        End Get
        Set
            Me.prop_pier_tie_quantity = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pier Tie Size")>
    Public Property pier_tie_size() As Integer?
        Get
            Return Me.prop_pier_tie_size
        End Get
        Set
            Me.prop_pier_tie_size = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pier Reinforcement Type")>
    Public Property pier_reinforcement_type() As String
        Get
            Return Me.prop_pier_reinforcement_type
        End Get
        Set
            Me.prop_pier_reinforcement_type = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pier Clear Cover")>
    Public Property pier_clear_cover() As Double?
        Get
            Return Me.prop_pier_clear_cover
        End Get
        Set
            Me.prop_pier_clear_cover = Value
        End Set
    End Property

    <Category("Unit Base Details"), Description(""), DisplayName("Pad Width 1")>
    Public Property pad_width_1() As Double?
        Get
            Return Me.prop_pad_width_1
        End Get
        Set
            Me.prop_pad_width_1 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Width 2")>
    Public Property pad_width_2() As Double?
        Get
            Return Me.prop_pad_width_2
        End Get
        Set
            Me.prop_pad_width_2 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Thickness")>
    Public Property pad_thickness() As Double?
        Get
            Return Me.prop_pad_thickness
        End Get
        Set
            Me.prop_pad_thickness = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Size Top Direction 1")>
    Public Property pad_rebar_size_top_dir1() As Integer?
        Get
            Return Me.prop_pad_rebar_size_top_dir1
        End Get
        Set
            Me.prop_pad_rebar_size_top_dir1 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Size Bottom Direction 1")>
    Public Property pad_rebar_size_bottom_dir1() As Integer?
        Get
            Return Me.prop_pad_rebar_size_bottom_dir1
        End Get
        Set
            Me.prop_pad_rebar_size_bottom_dir1 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Size Top Direction 2")>
    Public Property pad_rebar_size_top_dir2() As Integer?
        Get
            Return Me.prop_pad_rebar_size_top_dir2
        End Get
        Set
            Me.prop_pad_rebar_size_top_dir2 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Size Bottom Direction 2")>
    Public Property pad_rebar_size_bottom_dir2() As Integer?
        Get
            Return Me.prop_pad_rebar_size_bottom_dir2
        End Get
        Set
            Me.prop_pad_rebar_size_bottom_dir2 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Quantity Top Direction 1")>
    Public Property pad_rebar_quantity_top_dir1() As Integer?
        Get
            Return Me.prop_pad_rebar_quantity_top_dir1
        End Get
        Set
            Me.prop_pad_rebar_quantity_top_dir1 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Quantity Bottom Direction 1")>
    Public Property pad_rebar_quantity_bottom_dir1() As Integer?
        Get
            Return Me.prop_pad_rebar_quantity_bottom_dir1
        End Get
        Set
            Me.prop_pad_rebar_quantity_bottom_dir1 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Quantity Top Direction 2")>
    Public Property pad_rebar_quantity_top_dir2() As Integer?
        Get
            Return Me.prop_pad_rebar_quantity_top_dir2
        End Get
        Set
            Me.prop_pad_rebar_quantity_top_dir2 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Quantity Bottom Direction 2")>
    Public Property pad_rebar_quantity_bottom_dir2() As Integer?
        Get
            Return Me.prop_pad_rebar_quantity_bottom_dir2
        End Get
        Set
            Me.prop_pad_rebar_quantity_bottom_dir2 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Clear Cover")>
    Public Property pad_clear_cover() As Double?
        Get
            Return Me.prop_pad_clear_cover
        End Get
        Set
            Me.prop_pad_clear_cover = Value
        End Set
    End Property

    <Category("Unit Base Details"), Description(""), DisplayName("Total Soil Unit Weight")>
    Public Property total_soil_unit_weight() As Double?
        Get
            Return Me.prop_total_soil_unit_weight
        End Get
        Set
            Me.prop_total_soil_unit_weight = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Bearing Type")>
    Public Property bearing_type() As String
        Get
            Return Me.prop_bearing_type
        End Get
        Set
            Me.prop_bearing_type = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Nominal Bearing Capacity")>
    Public Property nominal_bearing_capacity() As Double?
        Get
            Return Me.prop_nominal_bearing_capacity
        End Get
        Set
            Me.prop_nominal_bearing_capacity = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Cohesion")>
    Public Property cohesion() As Double?
        Get
            Return Me.prop_cohesion
        End Get
        Set
            Me.prop_cohesion = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Friction Angle")>
    Public Property friction_angle() As Double?
        Get
            Return Me.prop_friction_angle
        End Get
        Set
            Me.prop_friction_angle = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("SPT Blow Count")>
    Public Property spt_blow_count() As Integer?
        Get
            Return Me.prop_spt_blow_count
        End Get
        Set
            Me.prop_spt_blow_count = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Base Friction Factor")>
    Public Property base_friction_factor() As Double?
        Get
            Return Me.prop_base_friction_factor
        End Get
        Set
            Me.prop_base_friction_factor = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Neglect Depth")>
    Public Property neglect_depth() As Double?
        Get
            Return Me.prop_neglect_depth
        End Get
        Set
            Me.prop_neglect_depth = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Bearing Distribution Type")>
    Public Property bearing_distribution_type() As Boolean
        Get
            Return Me.prop_bearing_distribution_type
        End Get
        Set
            Me.prop_bearing_distribution_type = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Groundwater Depth")>
    Public Property groundwater_depth() As Double?
        Get
            Return Me.prop_groundwater_depth
        End Get
        Set
            Me.prop_groundwater_depth = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Basic Soil Interaction up to 110% Acceptable1?")>
    Public Property basic_soil_check() As Boolean
        Get
            Return Me.prop_basic_soil_check
        End Get
        Set
            Me.prop_basic_soil_check = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Structural Checks up to 105% Acceptable?")>
    Public Property structural_check() As Boolean
        Get
            Return Me.prop_structural_check
        End Get
        Set
            Me.prop_structural_check = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Tool Version")>
    Public Property tool_version() As String
        Get
            Return Me.prop_tool_version
        End Get
        Set
            Me.prop_tool_version = Value
        End Set
    End Property
    '<Category("Unit Base Details"), Description(""), DisplayName("Modified")>
    'Public Property modified() As Boolean
    '    Get
    '        Return Me.prop_modified
    '    End Get
    '    Set
    '        Me.prop_modified = Value
    '    End Set
    'End Property
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal UnitBaseDataRow As DataRow, refID As Integer)
        Try
            Me.unit_base_id = refID
        Catch
            Me.unit_base_id = 0
        End Try 'Unit Base ID
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("extension_above_grade"), Double)) Then
                Me.extension_above_grade = CType(UnitBaseDataRow.Item("extension_above_grade"), Double)
            Else
                Me.extension_above_grade = Nothing
            End If
        Catch
            Me.extension_above_grade = Nothing
        End Try 'Extension Above Grade
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("foundation_depth"), Double)) Then
                Me.foundation_depth = CType(UnitBaseDataRow.Item("foundation_depth"), Double)
            Else
                Me.foundation_depth = Nothing
            End If
        Catch
            Me.foundation_depth = Nothing
        End Try 'Foundation Depth
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("concrete_compressive_strength"), Double)) Then
                Me.concrete_compressive_strength = CType(UnitBaseDataRow.Item("concrete_compressive_strength"), Double)
            Else
                Me.concrete_compressive_strength = Nothing
            End If
        Catch
            Me.concrete_compressive_strength = Nothing
        End Try 'Concrete Compressive Strength
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("dry_concrete_density"), Double)) Then
                Me.dry_concrete_density = CType(UnitBaseDataRow.Item("dry_concrete_density"), Double)
            Else
                Me.dry_concrete_density = Nothing
            End If
        Catch
            Me.dry_concrete_density = Nothing
        End Try 'Dry Concrete Density
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("rebar_grade"), Double)) Then
                Me.rebar_grade = CType(UnitBaseDataRow.Item("rebar_grade"), Double)
            Else
                Me.rebar_grade = Nothing
            End If
        Catch
            Me.rebar_grade = Nothing
        End Try 'Rebar Grade
        Try
            Me.top_and_bottom_rebar_different = CType(UnitBaseDataRow.Item("top_and_bottom_rebar_different"), Boolean)
        Catch
            Me.top_and_bottom_rebar_different = False
        End Try 'Top and Bottom Rebar Different
        Try
            Me.block_foundation = CType(UnitBaseDataRow.Item("block_foundation"), Boolean)
        Catch
            Me.block_foundation = False
        End Try 'Block Foundation 
        Try
            Me.rectangular_foundation = CType(UnitBaseDataRow.Item("rectangular_foundation"), Boolean)
        Catch
            Me.rectangular_foundation = False
        End Try 'Rectangular Foundation
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("base_plate_distance_above_foundation"), Double)) Then
                Me.base_plate_distance_above_foundation = CType(UnitBaseDataRow.Item("base_plate_distance_above_foundation"), Double)
            Else
                Me.base_plate_distance_above_foundation = Nothing
            End If
        Catch
            Me.base_plate_distance_above_foundation = Nothing
        End Try 'Base Plate Distance Above Foundation
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("bolt_circle_bearing_plate_width"), Double)) Then
                Me.bolt_circle_bearing_plate_width = CType(UnitBaseDataRow.Item("bolt_circle_bearing_plate_width"), Double)
            Else
                Me.bolt_circle_bearing_plate_width = Nothing
            End If
        Catch
            Me.bolt_circle_bearing_plate_width = Nothing
        End Try 'Bolt Circle Bearing Plate Width
        Try
            Me.tower_centroid_offset = CType(UnitBaseDataRow.Item("tower_centroid_offset"), Boolean)
        Catch
            Me.tower_centroid_offset = False
        End Try 'Tower Centroid Offset 

        Try
            Me.pier_shape = CType(UnitBaseDataRow.Item("pier_shape"), String)
        Catch
            Me.pier_shape = ""
        End Try 'Pier Shape
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pier_diameter"), Double)) Then
                Me.pier_diameter = CType(UnitBaseDataRow.Item("pier_diameter"), Double)
            Else
                Me.pier_diameter = Nothing
            End If
        Catch
            Me.pier_diameter = Nothing
        End Try 'Pier Diameter
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pier_rebar_quantity"), Integer)) Then
                Me.pier_rebar_quantity = CType(UnitBaseDataRow.Item("pier_rebar_quantity"), Integer)
            Else
                Me.pier_rebar_quantity = Nothing
            End If
        Catch
            Me.pier_rebar_quantity = Nothing
        End Try 'Pier Rebar Quantity
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pier_rebar_size"), Integer)) Then
                Me.pier_rebar_size = CType(UnitBaseDataRow.Item("pier_rebar_size"), Integer)
            Else
                Me.pier_rebar_size = Nothing
            End If
        Catch
            Me.pier_rebar_size = Nothing
        End Try 'Pier Rebar Size
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pier_tie_quantity"), Integer)) Then
                Me.pier_tie_quantity = CType(UnitBaseDataRow.Item("pier_tie_quantity"), Integer)
            Else
                Me.pier_tie_quantity = Nothing
            End If
        Catch
            Me.pier_tie_quantity = Nothing
        End Try 'Pier Tie Quantity
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pier_tie_size"), Integer)) Then
                Me.pier_tie_size = CType(UnitBaseDataRow.Item("pier_tie_size"), Integer)
            Else
                Me.pier_tie_size = Nothing
            End If
        Catch
            Me.pier_tie_size = Nothing
        End Try 'Pier Tie Size
        Try
            Me.pier_reinforcement_type = CType(UnitBaseDataRow.Item("pier_reinforcement_type"), String)
        Catch
            Me.pier_reinforcement_type = "Tie"
        End Try 'Pier Reinforcement Type
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pier_clear_cover"), Double)) Then
                Me.pier_clear_cover = CType(UnitBaseDataRow.Item("pier_clear_cover"), Double)
            Else
                Me.pier_clear_cover = Nothing
            End If
        Catch
            Me.pier_clear_cover = Nothing
        End Try 'Pier Clear Cover

        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_width_1"), Double)) Then
                Me.pad_width_1 = CType(UnitBaseDataRow.Item("pad_width_1"), Double)
            Else
                Me.pad_width_1 = Nothing
            End If
        Catch
            Me.pad_width_1 = Nothing
        End Try 'Pad Width 1
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_width_2"), Double)) Then
                Me.pad_width_2 = CType(UnitBaseDataRow.Item("pad_width_2"), Double)
            Else
                Me.pad_width_2 = Nothing
            End If
        Catch
            Me.pad_width_2 = Nothing
        End Try 'Pad Width 2
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_thickness"), Double)) Then
                Me.pad_thickness = CType(UnitBaseDataRow.Item("pad_thickness"), Double)
            Else
                Me.pad_thickness = Nothing
            End If
        Catch
            Me.pad_thickness = Nothing
        End Try 'Pad Thickness
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_rebar_size_top_dir1"), Integer)) Then
                Me.pad_rebar_size_top_dir1 = CType(UnitBaseDataRow.Item("pad_rebar_size_top_dir1"), Integer)
            Else
                Me.pad_rebar_size_top_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_size_top_dir1 = Nothing
        End Try 'Pad Rebar Size (Top Direction 1)
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_rebar_size_bottom_dir1"), Integer)) Then
                Me.pad_rebar_size_bottom_dir1 = CType(UnitBaseDataRow.Item("pad_rebar_size_bottom_dir1"), Integer)
            Else
                Me.pad_rebar_size_bottom_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_size_bottom_dir1 = Nothing
        End Try 'Pad Rebar Size (Bottom Direction 1)
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_rebar_size_top_dir2"), Integer)) Then
                Me.pad_rebar_size_top_dir2 = CType(UnitBaseDataRow.Item("pad_rebar_size_top_dir2"), Integer)
            Else
                Me.pad_rebar_size_top_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_size_top_dir2 = Nothing
        End Try 'Pad Rebar Size (Top Direction 2)
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_rebar_size_bottom_dir2"), Integer)) Then
                Me.pad_rebar_size_bottom_dir2 = CType(UnitBaseDataRow.Item("pad_rebar_size_bottom_dir2"), Integer)
            Else
                Me.pad_rebar_size_bottom_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_size_bottom_dir2 = Nothing
        End Try 'Pad Rebar Size (Bottom Direction 2)
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_rebar_quantity_top_dir1"), Integer)) Then
                Me.pad_rebar_quantity_top_dir1 = CType(UnitBaseDataRow.Item("pad_rebar_quantity_top_dir1"), Integer)
            Else
                Me.pad_rebar_quantity_top_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_top_dir1 = Nothing
        End Try 'Pad Rebar Quantity (Top Direction 1)
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_rebar_quantity_bottom_dir1"), Integer)) Then
                Me.pad_rebar_quantity_bottom_dir1 = CType(UnitBaseDataRow.Item("pad_rebar_quantity_bottom_dir1"), Integer)
            Else
                Me.pad_rebar_quantity_bottom_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_bottom_dir1 = Nothing
        End Try 'Pad Rebar Quantity (Bottom Direction 1)
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_rebar_quantity_top_dir2"), Integer)) Then
                Me.pad_rebar_quantity_top_dir2 = CType(UnitBaseDataRow.Item("pad_rebar_quantity_top_dir2"), Integer)
            Else
                Me.pad_rebar_quantity_top_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_top_dir2 = Nothing
        End Try 'Pad Rebar Quantity (Top Direction 2)
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_rebar_quantity_bottom_dir2"), Integer)) Then
                Me.pad_rebar_quantity_bottom_dir2 = CType(UnitBaseDataRow.Item("pad_rebar_quantity_bottom_dir2"), Integer)
            Else
                Me.pad_rebar_quantity_bottom_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_bottom_dir2 = Nothing
        End Try 'Pad Rebar Quantity (Bottom Direction 2)
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("pad_clear_cover"), Double)) Then
                Me.pad_clear_cover = CType(UnitBaseDataRow.Item("pad_clear_cover"), Double)
            Else
                Me.pad_clear_cover = Nothing
            End If
        Catch
            Me.pad_clear_cover = Nothing
        End Try 'Pad Clear Cover

        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("total_soil_unit_weight"), Double)) Then
                Me.total_soil_unit_weight = CType(UnitBaseDataRow.Item("total_soil_unit_weight"), Double)
            Else
                Me.total_soil_unit_weight = Nothing
            End If
        Catch
            Me.total_soil_unit_weight = Nothing
        End Try 'Total Soil Unit Weight
        Try
            Me.bearing_type = CType(UnitBaseDataRow.Item("bearing_type"), String)
        Catch
            Me.bearing_type = "Ultimate Gross Bearing, Qult:"
        End Try 'Bearing Type
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("nominal_bearing_capacity"), Double)) Then
                Me.nominal_bearing_capacity = CType(UnitBaseDataRow.Item("nominal_bearing_capacity"), Double)
            Else
                Me.nominal_bearing_capacity = Nothing
            End If
        Catch
            Me.nominal_bearing_capacity = Nothing
        End Try 'Nominal Bearing Capacity
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("cohesion"), Double)) Then
                Me.cohesion = CType(UnitBaseDataRow.Item("cohesion"), Double)
            Else
                Me.cohesion = Nothing
            End If
        Catch
            Me.cohesion = Nothing
        End Try 'Cohesion
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("friction_angle"), Double)) Then
                Me.friction_angle = CType(UnitBaseDataRow.Item("friction_angle"), Double)
            Else
                Me.friction_angle = Nothing
            End If
        Catch
            Me.friction_angle = Nothing
        End Try 'Friction Angle
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("spt_blow_count"), Integer)) Then
                Me.spt_blow_count = CType(UnitBaseDataRow.Item("spt_blow_count"), Integer)
            Else
                Me.spt_blow_count = Nothing
            End If
        Catch
            Me.spt_blow_count = Nothing
        End Try 'STP Blow Count
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("base_friction_factor"), Double)) Then
                Me.base_friction_factor = CType(UnitBaseDataRow.Item("base_friction_factor"), Double)
            Else
                Me.base_friction_factor = Nothing
            End If
        Catch
            Me.base_friction_factor = Nothing
        End Try 'Base Friction Factor
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("neglect_depth"), Double)) Then
                Me.neglect_depth = CType(UnitBaseDataRow.Item("neglect_depth"), Double)
            Else
                Me.neglect_depth = Nothing
            End If
        Catch
            Me.neglect_depth = Nothing
        End Try 'Neglect Depth
        Try
            Me.bearing_distribution_type = CType(UnitBaseDataRow.Item("bearing_distribution_type"), Boolean)
        Catch
            Me.bearing_distribution_type = True
        End Try 'Bearing Distribution Type
        Try
            If Not IsDBNull(CType(UnitBaseDataRow.Item("groundwater_depth"), Double)) Then
                Me.groundwater_depth = CType(UnitBaseDataRow.Item("groundwater_depth"), Double)
            Else
                Me.groundwater_depth = Nothing
            End If
        Catch
            Me.groundwater_depth = -1
        End Try 'Groundwater Depth
        Try
            Me.basic_soil_check = CType(UnitBaseDataRow.Item("SoilInteractionBoolean"), Boolean)
        Catch
            Me.basic_soil_check = False
        End Try 'Basic Soil Interaction up to 110% Acceptable?
        Try
            Me.structural_check = CType(UnitBaseDataRow.Item("StructuralCheckBoolean"), Boolean)
        Catch
            Me.structural_check = False
        End Try 'Structural Checks up to 105.0% Acceptable?
        Try
            Me.tool_version = CType(UnitBaseDataRow.Item("vnum"), String)
        Catch
            Me.tool_version = ""
        End Try 'Tool Version

        'If Me.modified = True Then
        '    For Each ModifiedRangeDataRow As DataRow In ds.Tables("Pier and Pad Modified Ranges SQL").Rows
        '        Dim modRefID As Integer = CType(ModifiedRangeDataRow.Item("modified_id"), Integer)
        '        If modRefID = refID Then
        '            Me.ModifiedRanges.Add(New ModifiedRange(ModifiedRangeDataRow))
        '        End If
        '    Next 'Add Modified Ranges to Modified Range Object
        'End If

    End Sub 'Generate a Unit Base object from EDS

    Public Sub New(ByVal path As String)
        Try
            Me.unit_base_id = CType(GetOneExcelRange(path, "ID"), Integer)
        Catch
            Me.unit_base_id = 0
        End Try 'Unit Base ID
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "E"), Double)) Then
                Me.extension_above_grade = CType(GetOneExcelRange(path, "E"), Double)
            Else
                Me.extension_above_grade = Nothing
            End If
        Catch
            Me.extension_above_grade = Nothing
        End Try 'Extension Above Grade
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D"), Double)) Then
                Me.foundation_depth = CType(GetOneExcelRange(path, "D"), Double)
            Else
                Me.foundation_depth = Nothing
            End If
        Catch
            Me.foundation_depth = Nothing
        End Try 'Foundation Depth
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "F\c"), Double)) Then
                Me.concrete_compressive_strength = CType(GetOneExcelRange(path, "F\c"), Double)
            Else
                Me.concrete_compressive_strength = Nothing
            End If
        Catch
            Me.concrete_compressive_strength = Nothing
        End Try 'Concrete Compressive Strength
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "ConcreteDensity"), Double)) Then
                Me.dry_concrete_density = CType(GetOneExcelRange(path, "ConcreteDensity"), Double)
            Else
                Me.dry_concrete_density = Nothing
            End If
        Catch
            Me.dry_concrete_density = Nothing
        End Try 'Dry Concrete Density
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Fy"), Double)) Then
                Me.rebar_grade = CType(GetOneExcelRange(path, "Fy"), Double)
            Else
                Me.rebar_grade = Nothing
            End If
        Catch
            Me.rebar_grade = Nothing
        End Try 'Rebar Grade
        Try
            Me.top_and_bottom_rebar_different = CType(GetOneExcelRange(path, "DifferentReinforcementBoolean"), Boolean)
        Catch
            Me.top_and_bottom_rebar_different = False
        End Try 'Top and Bottom Rebar Different
        Try
            Me.block_foundation = CType(GetOneExcelRange(path, "BlockFoundationBoolean"), Boolean)
        Catch
            Me.block_foundation = False
        End Try 'Block Foundation 
        Try
            Me.rectangular_foundation = CType(GetOneExcelRange(path, "RectangularPadBoolean"), Boolean)
        Catch
            Me.rectangular_foundation = False
        End Try 'Rectangular Foundation
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "bpdist"), Double)) Then
                Me.base_plate_distance_above_foundation = CType(GetOneExcelRange(path, "bpdist"), Double)
            Else
                Me.base_plate_distance_above_foundation = Nothing
            End If
        Catch
            Me.base_plate_distance_above_foundation = Nothing
        End Try 'Base Plate Distance Above Foundation
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "BC"), Double)) Then
                Me.bolt_circle_bearing_plate_width = CType(GetOneExcelRange(path, "BC"), Double)
            Else
                Me.bolt_circle_bearing_plate_width = Nothing
            End If
        Catch
            Me.bolt_circle_bearing_plate_width = Nothing
        End Try 'Bolt Circle Bearing Plate Width
        Try
            Me.tower_centroid_offset = CType(GetOneExcelRange(path, "TowerCentroidOffsetBoolean"), Boolean)
        Catch
            Me.tower_centroid_offset = False
        End Try 'Tower Centroid Offset 

        Try
            Me.pier_shape = CType(GetOneExcelRange(path, "shape"), String)
        Catch
            Me.pier_shape = "Circular"
        End Try 'Pier Shape
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "dpier"), Double)) Then
                Me.pier_diameter = CType(GetOneExcelRange(path, "dpier"), Double)
            Else
                Me.pier_diameter = Nothing
            End If
        Catch
            Me.pier_diameter = Nothing
        End Try 'Pier Diameter
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "mc"), Integer)) Then
                Me.pier_rebar_quantity = CType(GetOneExcelRange(path, "mc"), Integer)
            Else
                Me.pier_rebar_quantity = Nothing
            End If
        Catch
            Me.pier_rebar_quantity = Nothing
        End Try 'Pier Rebar Quantity
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Sc"), Integer)) Then
                Me.pier_rebar_size = CType(GetOneExcelRange(path, "Sc"), Integer)
            Else
                Me.pier_rebar_size = Nothing
            End If
        Catch
            Me.pier_rebar_size = Nothing
        End Try 'Pier Rebar Size
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "mt"), Integer)) Then
                Me.pier_tie_quantity = CType(GetOneExcelRange(path, "mt"), Integer)
            Else
                Me.pier_tie_quantity = Nothing
            End If
        Catch
            Me.pier_tie_quantity = Nothing
        End Try 'Pier Tie Quantity
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "St"), Integer)) Then
                Me.pier_tie_size = CType(GetOneExcelRange(path, "St"), Integer)
            Else
                Me.pier_tie_size = Nothing
            End If
        Catch
            Me.pier_tie_size = Nothing
        End Try 'Pier Tie Size
        Try
            Me.pier_reinforcement_type = CType(GetOneExcelRange(path, "PierReinfType"), String)
        Catch
            Me.pier_reinforcement_type = "Tie"
        End Try 'Pier Reinforcement Type
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "ccpier"), Double)) Then
                Me.pier_clear_cover = CType(GetOneExcelRange(path, "ccpier"), Double)
            Else
                Me.pier_clear_cover = Nothing
            End If
        Catch
            Me.pier_clear_cover = Nothing
        End Try 'Pier Clear Cover

        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "W"), Double)) Then
                Me.pad_width_1 = CType(GetOneExcelRange(path, "W"), Double)
            Else
                Me.pad_width_1 = Nothing
            End If
        Catch
            Me.pad_width_1 = Nothing
        End Try 'Pad Width 1
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "W.dir2"), Double)) Then
                Me.pad_width_2 = CType(GetOneExcelRange(path, "W.dir2"), Double)
            Else
                Me.pad_width_2 = Nothing
            End If
        Catch
            Me.pad_width_2 = Nothing
        End Try 'Pad Width 2
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "T"), Double)) Then
                Me.pad_thickness = CType(GetOneExcelRange(path, "T"), Double)
            Else
                Me.pad_thickness = Nothing
            End If
        Catch
            Me.pad_thickness = Nothing
        End Try 'Pad Thickness
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "sptop"), Integer)) Then
                Me.pad_rebar_size_top_dir1 = CType(GetOneExcelRange(path, "sptop"), Integer)
            Else
                Me.pad_rebar_size_top_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_size_top_dir1 = Nothing
        End Try 'Pad Rebar Size (Top Direction 1)
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Sp"), Integer)) Then
                Me.pad_rebar_size_bottom_dir1 = CType(GetOneExcelRange(path, "Sp"), Integer)
            Else
                Me.pad_rebar_size_bottom_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_size_bottom_dir1 = Nothing
        End Try 'Pad Rebar Size (Bottom Direction 1)
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "sptop2"), Integer)) Then
                Me.pad_rebar_size_top_dir2 = CType(GetOneExcelRange(path, "sptop2"), Integer)
            Else
                Me.pad_rebar_size_top_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_size_top_dir2 = Nothing
        End Try 'Pad Rebar Size (Top Direction 2)
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "sp_2"), Integer)) Then
                Me.pad_rebar_size_bottom_dir2 = CType(GetOneExcelRange(path, "sp_2"), Integer)
            Else
                Me.pad_rebar_size_bottom_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_size_bottom_dir2 = Nothing
        End Try 'Pad Rebar Size (Bottom Direction 2)
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "mptop"), Integer)) Then
                Me.pad_rebar_quantity_top_dir1 = CType(GetOneExcelRange(path, "mptop"), Integer)
            Else
                Me.pad_rebar_quantity_top_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_top_dir1 = Nothing
        End Try 'Pad Rebar Quantity (Top Direction 1)
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "mp"), Integer)) Then
                Me.pad_rebar_quantity_bottom_dir1 = CType(GetOneExcelRange(path, "mp"), Integer)
            Else
                Me.pad_rebar_quantity_bottom_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_bottom_dir1 = Nothing
        End Try 'Pad Rebar Quantity (Bottom Direction 1)
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "mptop2"), Integer)) Then
                Me.pad_rebar_quantity_top_dir2 = CType(GetOneExcelRange(path, "mptop2"), Integer)
            Else
                Me.pad_rebar_quantity_top_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_top_dir2 = Nothing
        End Try 'Pad Rebar Quantity (Top Direction 2)
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "mp_2"), Integer)) Then
                Me.pad_rebar_quantity_bottom_dir2 = CType(GetOneExcelRange(path, "mp_2"), Integer)
            Else
                Me.pad_rebar_quantity_bottom_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_bottom_dir2 = Nothing
        End Try 'Pad Rebar Quantity (Bottom Direction 2)
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "ccpad"), Double)) Then
                Me.pad_clear_cover = CType(GetOneExcelRange(path, "ccpad"), Double)
            Else
                Me.pad_clear_cover = Nothing
            End If
        Catch
            Me.pad_clear_cover = Nothing
        End Try 'Pad Clear Cover

        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "γ"), Double)) Then
                Me.total_soil_unit_weight = CType(GetOneExcelRange(path, "γ"), Double)
            Else
                Me.total_soil_unit_weight = Nothing
            End If
        Catch
            Me.total_soil_unit_weight = Nothing
        End Try 'Total Soil Unit Weight
        Try
            Me.bearing_type = CType(GetOneExcelRange(path, "BearingType"), String)
        Catch
            Me.bearing_type = "Ultimate Gross Bearing, Qult:"
        End Try 'Bearing Type
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Qinput"), Double)) Then
                Me.nominal_bearing_capacity = CType(GetOneExcelRange(path, "Qinput"), Double)
            Else
                Me.nominal_bearing_capacity = Nothing
            End If
        Catch
            Me.nominal_bearing_capacity = Nothing
        End Try 'Nominal Bearing Capacity
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Cu"), Double)) Then
                Me.cohesion = CType(GetOneExcelRange(path, "Cu"), Double)
            Else
                Me.cohesion = Nothing
            End If
        Catch
            Me.cohesion = Nothing
        End Try 'Cohesion
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "ϕ"), Double)) Then
                Me.friction_angle = CType(GetOneExcelRange(path, "ϕ"), Double)
            Else
                Me.friction_angle = Nothing
            End If
        Catch
            Me.friction_angle = Nothing
        End Try 'Friction Angle
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "N_blows"), Integer)) Then
                Me.spt_blow_count = CType(GetOneExcelRange(path, "N_blows"), Integer)
            Else
                Me.spt_blow_count = Nothing
            End If
        Catch
            Me.spt_blow_count = Nothing
        End Try 'STP Blow Count
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "μ"), Double)) Then
                Me.base_friction_factor = CType(GetOneExcelRange(path, "μ"), Double)
            Else
                Me.base_friction_factor = Nothing
            End If
        Catch
            Me.base_friction_factor = Nothing
        End Try 'Base Friction Factor
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "N"), Double)) Then
                Me.neglect_depth = CType(GetOneExcelRange(path, "N"), Double)
            Else
                Me.neglect_depth = Nothing
            End If
        Catch
            Me.neglect_depth = Nothing
        End Try 'Neglect Depth
        Try
            If CType(GetOneExcelRange(path, "Rock"), String) = "Yes" Then
                Me.bearing_distribution_type = False
            Else
                Me.bearing_distribution_type = True
            End If
        Catch
            Me.bearing_distribution_type = True
        End Try 'Bearing Distribution Type
        Try
            Me.groundwater_depth = CType(GetOneExcelRange(path, "gw"), Double)
        Catch
            Me.groundwater_depth = -1
        End Try 'Groundwater Depth
        Try
            Me.basic_soil_check = CType(GetOneExcelRange(path, "SoilInteractionBoolean"), Boolean)
        Catch
            Me.basic_soil_check = False
        End Try 'Basic Soil Interaction up to 110% Acceptable?
        Try
            Me.structural_check = CType(GetOneExcelRange(path, "StructuralCheckBoolean"), Boolean)
        Catch
            Me.structural_check = False
        End Try 'Structural Checks up to 105.0% Acceptable?
        Try
            Me.tool_version = CType(GetOneExcelRange(path, "vnum"), String)
        Catch
            Me.tool_version = Nothing
        End Try 'Tool Version

    End Sub 'Generate a Unit Base object from Excel

#End Region

End Class
