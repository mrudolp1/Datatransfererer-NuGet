Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
Partial Public Class Pile

#Region "Define"
    Private prop_pile_id As Integer?
    Private prop_load_eccentricity As Double?
    Private prop_bolt_circle_bearing_plate_width As Double?
    Private prop_pile_shape As String
    Private prop_pile_material As String
    Private prop_pile_length As Double?
    Private prop_pile_diameter_width As Double?
    Private prop_pile_pipe_thickness As Double?
    Private prop_pile_soil_capacity_given As Boolean
    Private prop_steel_yield_strength As Double?
    Private prop_pile_type_option As String
    Private prop_rebar_quantity As Integer?
    Private prop_pile_group_config As String
    Private prop_foundation_depth As Double?
    Private prop_pad_thickness As Double?
    Private prop_pad_width_dir1 As Double?
    Private prop_pad_width_dir2 As Double?
    Private prop_pad_rebar_size_bottom As Integer?
    Private prop_pad_rebar_size_top As Integer?
    Private prop_pad_rebar_quantity_bottom_dir1 As Integer?
    Private prop_pad_rebar_quantity_top_dir1 As Integer?
    Private prop_pad_rebar_quantity_bottom_dir2 As Integer?
    Private prop_pad_rebar_quantity_top_dir2 As Integer?
    Private prop_pier_shape As String
    Private prop_pier_diameter As Integer?
    Private prop_extension_above_grade As Double?
    Private prop_pier_rebar_size As Integer?
    Private prop_pier_rebar_quantity As Integer?
    Private prop_pier_tie_size As Integer?
    'Private prop_pier_tie_quantity As Integer? '(Remove, not applicable for Piles. Need to remove from SQL Database)
    Private prop_rebar_grade As Double?
    Private prop_concrete_compressive_strength As Double?
    Private prop_groundwater_depth As Double?
    Private prop_total_soil_unit_weight As Double?
    Private prop_cohesion As Double?
    Private prop_friction_angle As Double?
    Private prop_neglect_depth As Double?
    Private prop_spt_blow_count As Integer?
    Private prop_pile_negative_friction_force As Double?
    Private prop_pile_ultimate_compression As Double?
    Private prop_pile_ultimate_tension As Double?
    Private prop_top_and_bottom_rebar_different As Boolean
    'Private prop_top_and_bottom_rebar_different As String
    Private prop_ultimate_gross_end_bearing As Double?
    Private prop_skin_friction_given As Boolean
    Private prop_pile_quantity_circular As Integer?
    Private prop_group_diameter_circular As Double?
    Private prop_pile_column_quantity As Integer?
    Private prop_pile_row_quantity As Integer?
    Private prop_pile_columns_spacing As Double?
    Private prop_pile_row_spacing As Double?
    Private prop_group_efficiency_factor_given As Boolean
    Private prop_group_efficiency_factor As Double?
    Private prop_cap_type As String
    Private prop_pile_quantity_asymmetric As Integer?
    Private prop_pile_spacing_min_asymmetric As Double?
    Private prop_quantity_piles_surrounding As Integer?
    Public Property soil_layers As New List(Of PileSoilLayer)
    Public Property pile_locations As New List(Of PileLocation)
    <Category("Pile Details"), Description(""), DisplayName("Pile_Id")>
    Public Property pile_id() As Integer?
        Get
            Return Me.prop_pile_id
        End Get
        Set
            Me.prop_pile_id = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Load_Eccentricity")>
    Public Property load_eccentricity() As Double?
        Get
            Return Me.prop_load_eccentricity
        End Get
        Set
            Me.prop_load_eccentricity = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Bolt_Circle_Bearing_Plate_Width")>
    Public Property bolt_circle_bearing_plate_width() As Double?
        Get
            Return Me.prop_bolt_circle_bearing_plate_width
        End Get
        Set
            Me.prop_bolt_circle_bearing_plate_width = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Shape")>
    Public Property pile_shape() As String
        Get
            Return Me.prop_pile_shape
        End Get
        Set
            Me.prop_pile_shape = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Material")>
    Public Property pile_material() As String
        Get
            Return Me.prop_pile_material
        End Get
        Set
            Me.prop_pile_material = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Length")>
    Public Property pile_length() As Double?
        Get
            Return Me.prop_pile_length
        End Get
        Set
            Me.prop_pile_length = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Diameter_Width")>
    Public Property pile_diameter_width() As Double?
        Get
            Return Me.prop_pile_diameter_width
        End Get
        Set
            Me.prop_pile_diameter_width = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Pipe_Thickness")>
    Public Property pile_pipe_thickness() As Double?
        Get
            Return Me.prop_pile_pipe_thickness
        End Get
        Set
            Me.prop_pile_pipe_thickness = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Soil_Capacity_Given")>
    Public Property pile_soil_capacity_given() As Boolean
        Get
            Return Me.prop_pile_soil_capacity_given
        End Get
        Set
            Me.prop_pile_soil_capacity_given = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Steel_Yield_Strength")>
    Public Property steel_yield_strength() As Double?
        Get
            Return Me.prop_steel_yield_strength
        End Get
        Set
            Me.prop_steel_yield_strength = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Type_Option")>
    Public Property pile_type_option() As String
        Get
            Return Me.prop_pile_type_option
        End Get
        Set
            Me.prop_pile_type_option = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Rebar_Quantity")>
    Public Property rebar_quantity() As Integer?
        Get
            Return Me.prop_rebar_quantity
        End Get
        Set
            Me.prop_rebar_quantity = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Group_Config")>
    Public Property pile_group_config() As String
        Get
            Return Me.prop_pile_group_config
        End Get
        Set
            Me.prop_pile_group_config = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Foundation_Depth")>
    Public Property foundation_depth() As Double?
        Get
            Return Me.prop_foundation_depth
        End Get
        Set
            Me.prop_foundation_depth = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pad_Thickness")>
    Public Property pad_thickness() As Double?
        Get
            Return Me.prop_pad_thickness
        End Get
        Set
            Me.prop_pad_thickness = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pad_Width_Dir1")>
    Public Property pad_width_dir1() As Double?
        Get
            Return Me.prop_pad_width_dir1
        End Get
        Set
            Me.prop_pad_width_dir1 = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pad_Width_Dir2")>
    Public Property pad_width_dir2() As Double?
        Get
            Return Me.prop_pad_width_dir2
        End Get
        Set
            Me.prop_pad_width_dir2 = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pad_Rebar_Size_Bottom")>
    Public Property pad_rebar_size_bottom() As Integer?
        Get
            Return Me.prop_pad_rebar_size_bottom
        End Get
        Set
            Me.prop_pad_rebar_size_bottom = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pad_Rebar_Size_Top")>
    Public Property pad_rebar_size_top() As Integer?
        Get
            Return Me.prop_pad_rebar_size_top
        End Get
        Set
            Me.prop_pad_rebar_size_top = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pad_Rebar_Quantity_Bottom_Dir1")>
    Public Property pad_rebar_quantity_bottom_dir1() As Integer?
        Get
            Return Me.prop_pad_rebar_quantity_bottom_dir1
        End Get
        Set
            Me.prop_pad_rebar_quantity_bottom_dir1 = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pad_Rebar_Quantity_Top_Dir1")>
    Public Property pad_rebar_quantity_top_dir1() As Integer?
        Get
            Return Me.prop_pad_rebar_quantity_top_dir1
        End Get
        Set
            Me.prop_pad_rebar_quantity_top_dir1 = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pad_Rebar_Quantity_Bottom_Dir2")>
    Public Property pad_rebar_quantity_bottom_dir2() As Integer?
        Get
            Return Me.prop_pad_rebar_quantity_bottom_dir2
        End Get
        Set
            Me.prop_pad_rebar_quantity_bottom_dir2 = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pad_Rebar_Quantity_Top_Dir2")>
    Public Property pad_rebar_quantity_top_dir2() As Integer?
        Get
            Return Me.prop_pad_rebar_quantity_top_dir2
        End Get
        Set
            Me.prop_pad_rebar_quantity_top_dir2 = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pier_Shape")>
    Public Property pier_shape() As String
        Get
            Return Me.prop_pier_shape
        End Get
        Set
            Me.prop_pier_shape = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pier_Diameter")>
    Public Property pier_diameter() As Integer?
        Get
            Return Me.prop_pier_diameter
        End Get
        Set
            Me.prop_pier_diameter = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Extension_Above_Grade")>
    Public Property extension_above_grade() As Double?
        Get
            Return Me.prop_extension_above_grade
        End Get
        Set
            Me.prop_extension_above_grade = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pier_Rebar_Size")>
    Public Property pier_rebar_size() As Integer?
        Get
            Return Me.prop_pier_rebar_size
        End Get
        Set
            Me.prop_pier_rebar_size = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pier_Rebar_Quantity")>
    Public Property pier_rebar_quantity() As Integer?
        Get
            Return Me.prop_pier_rebar_quantity
        End Get
        Set
            Me.prop_pier_rebar_quantity = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pier_Tie_Size")>
    Public Property pier_tie_size() As Integer?
        Get
            Return Me.prop_pier_tie_size
        End Get
        Set
            Me.prop_pier_tie_size = Value
        End Set
    End Property
    '<Category("Pile Details"), Description(""), DisplayName("Pier_Tie_Quantity")>
    'Public Property pier_tie_quantity() As Integer?
    '    Get
    '        Return Me.prop_pier_tie_quantity
    '    End Get
    '    Set
    '        Me.prop_pier_tie_quantity = Value
    '    End Set
    'End Property
    <Category("Pile Details"), Description(""), DisplayName("Rebar_Grade")>
    Public Property rebar_grade() As Double?
        Get
            Return Me.prop_rebar_grade
        End Get
        Set
            Me.prop_rebar_grade = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Concrete_Compressive_Strength")>
    Public Property concrete_compressive_strength() As Double?
        Get
            Return Me.prop_concrete_compressive_strength
        End Get
        Set
            Me.prop_concrete_compressive_strength = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Groundwater_Depth")>
    Public Property groundwater_depth() As Double?
        Get
            Return Me.prop_groundwater_depth
        End Get
        Set
            Me.prop_groundwater_depth = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Total_Soil_Unit_Weight")>
    Public Property total_soil_unit_weight() As Double?
        Get
            Return Me.prop_total_soil_unit_weight
        End Get
        Set
            Me.prop_total_soil_unit_weight = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Cohesion")>
    Public Property cohesion() As Double?
        Get
            Return Me.prop_cohesion
        End Get
        Set
            Me.prop_cohesion = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Friction_Angle")>
    Public Property friction_angle() As Double?
        Get
            Return Me.prop_friction_angle
        End Get
        Set
            Me.prop_friction_angle = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Neglect_Depth")>
    Public Property neglect_depth() As Double?
        Get
            Return Me.prop_neglect_depth
        End Get
        Set
            Me.prop_neglect_depth = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Spt_Blow_Count")>
    Public Property spt_blow_count() As Integer?
        Get
            Return Me.prop_spt_blow_count
        End Get
        Set
            Me.prop_spt_blow_count = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Negative_Friction_Force")>
    Public Property pile_negative_friction_force() As Double?
        Get
            Return Me.prop_pile_negative_friction_force
        End Get
        Set
            Me.prop_pile_negative_friction_force = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Ultimate_Compression")>
    Public Property pile_ultimate_compression() As Double?
        Get
            Return Me.prop_pile_ultimate_compression
        End Get
        Set
            Me.prop_pile_ultimate_compression = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Ultimate_Tension")>
    Public Property pile_ultimate_tension() As Double?
        Get
            Return Me.prop_pile_ultimate_tension
        End Get
        Set
            Me.prop_pile_ultimate_tension = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Top_And_Bottom_Rebar_Different")>
    Public Property top_and_bottom_rebar_different() As Boolean
        Get
            Return Me.prop_top_and_bottom_rebar_different
        End Get
        Set
            Me.prop_top_and_bottom_rebar_different = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Ultimate_Gross_End_Bearing")>
    Public Property ultimate_gross_end_bearing() As Double?
        Get
            Return Me.prop_ultimate_gross_end_bearing
        End Get
        Set
            Me.prop_ultimate_gross_end_bearing = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Skin_Friction_Given")>
    Public Property skin_friction_given() As Boolean
        Get
            Return Me.prop_skin_friction_given
        End Get
        Set
            Me.prop_skin_friction_given = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Quantity_Circular")>
    Public Property pile_quantity_circular() As Integer?
        Get
            Return Me.prop_pile_quantity_circular
        End Get
        Set
            Me.prop_pile_quantity_circular = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Group_Diameter_Circular")>
    Public Property group_diameter_circular() As Double?
        Get
            Return Me.prop_group_diameter_circular
        End Get
        Set
            Me.prop_group_diameter_circular = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Column_Quantity")>
    Public Property pile_column_quantity() As Integer?
        Get
            Return Me.prop_pile_column_quantity
        End Get
        Set
            Me.prop_pile_column_quantity = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Row_Quantity")>
    Public Property pile_row_quantity() As Integer?
        Get
            Return Me.prop_pile_row_quantity
        End Get
        Set
            Me.prop_pile_row_quantity = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Columns_Spacing")>
    Public Property pile_columns_spacing() As Double?
        Get
            Return Me.prop_pile_columns_spacing
        End Get
        Set
            Me.prop_pile_columns_spacing = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Row_Spacing")>
    Public Property pile_row_spacing() As Double?
        Get
            Return Me.prop_pile_row_spacing
        End Get
        Set
            Me.prop_pile_row_spacing = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Group_Efficiency_Factor_Given")>
    Public Property group_efficiency_factor_given() As Boolean
        Get
            Return Me.prop_group_efficiency_factor_given
        End Get
        Set
            Me.prop_group_efficiency_factor_given = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Group_Efficiency_Factor")>
    Public Property group_efficiency_factor() As Double?
        Get
            Return Me.prop_group_efficiency_factor
        End Get
        Set
            Me.prop_group_efficiency_factor = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Cap_Type")>
    Public Property cap_type() As String
        Get
            Return Me.prop_cap_type
        End Get
        Set
            Me.prop_cap_type = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Quantity_Asymmetric")>
    Public Property pile_quantity_asymmetric() As Integer?
        Get
            Return Me.prop_pile_quantity_asymmetric
        End Get
        Set
            Me.prop_pile_quantity_asymmetric = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Pile_Spacing_Min_Asymmetric")>
    Public Property pile_spacing_min_asymmetric() As Double?
        Get
            Return Me.prop_pile_spacing_min_asymmetric
        End Get
        Set
            Me.prop_pile_spacing_min_asymmetric = Value
        End Set
    End Property
    <Category("Pile Details"), Description(""), DisplayName("Quantity_Piles_Surrounding")>
    Public Property quantity_piles_surrounding() As Integer?
        Get
            Return Me.prop_quantity_piles_surrounding
        End Get
        Set
            Me.prop_quantity_piles_surrounding = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal PileDataRow As DataRow, refID As Integer)
        Try
            Me.pile_id = refID
        Catch
            Me.pile_id = 0
        End Try 'Pile_Id
        Try
            If Not IsDBNull(CType(PileDataRow.Item("load_eccentricity"), Double)) Then
                Me.load_eccentricity = CType(PileDataRow.Item("load_eccentricity"), Double)
            Else
                Me.load_eccentricity = Nothing
            End If
        Catch
            Me.load_eccentricity = Nothing
        End Try 'Load_Eccentricity
        Try
            If Not IsDBNull(CType(PileDataRow.Item("bolt_circle_bearing_plate_width"), Double)) Then
                Me.bolt_circle_bearing_plate_width = CType(PileDataRow.Item("bolt_circle_bearing_plate_width"), Double)
            Else
                Me.bolt_circle_bearing_plate_width = Nothing
            End If
        Catch
            Me.bolt_circle_bearing_plate_width = Nothing
        End Try 'Bolt_Circle_Bearing_Plate_Width
        Try
            Me.pile_shape = CType(PileDataRow.Item("pile_shape"), String)
        Catch
            Me.pile_shape = ""
        End Try 'Pile_Shape
        Try
            Me.pile_material = CType(PileDataRow.Item("pile_material"), String)
        Catch
            Me.pile_material = ""
        End Try 'Pile_Material
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_length"), Double)) Then
                Me.pile_length = CType(PileDataRow.Item("pile_length"), Double)
            Else
                Me.pile_length = Nothing
            End If
        Catch
            Me.pile_length = Nothing
        End Try 'Pile_Length
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_diameter_width"), Double)) Then
                Me.pile_diameter_width = CType(PileDataRow.Item("pile_diameter_width"), Double)
            Else
                Me.pile_diameter_width = Nothing
            End If
        Catch
            Me.pile_diameter_width = Nothing
        End Try 'Pile_Diameter_Width
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_pipe_thickness"), Double)) Then
                Me.pile_pipe_thickness = CType(PileDataRow.Item("pile_pipe_thickness"), Double)
            Else
                Me.pile_pipe_thickness = Nothing
            End If
        Catch
            Me.pile_pipe_thickness = Nothing
        End Try 'Pile_Pipe_Thickness
        Try
            Me.pile_soil_capacity_given = CType(PileDataRow.Item("pile_soil_capacity_given"), Boolean)
        Catch
            Me.pile_soil_capacity_given = False
        End Try 'Pile_Soil_Capacity_Given
        Try
            If Not IsDBNull(CType(PileDataRow.Item("steel_yield_strength"), Double)) Then
                Me.steel_yield_strength = CType(PileDataRow.Item("steel_yield_strength"), Double)
            Else
                Me.steel_yield_strength = Nothing
            End If
        Catch
            Me.steel_yield_strength = Nothing
        End Try 'Steel_Yield_Strength
        Try
            Me.pile_type_option = CType(PileDataRow.Item("pile_type_option"), String)
        Catch
            Me.pile_type_option = ""
        End Try 'Pile_Type_Option
        Try
            If Not IsDBNull(CType(PileDataRow.Item("rebar_quantity"), Integer)) Then
                Me.rebar_quantity = CType(PileDataRow.Item("rebar_quantity"), Integer)
            Else
                Me.rebar_quantity = Nothing
            End If
        Catch
            Me.rebar_quantity = Nothing
        End Try 'Rebar_Quantity
        Try
            Me.pile_group_config = CType(PileDataRow.Item("pile_group_config"), String)
        Catch
            Me.pile_group_config = ""
        End Try 'Pile_Group_Config
        Try
            If Not IsDBNull(CType(PileDataRow.Item("foundation_depth"), Double)) Then
                Me.foundation_depth = CType(PileDataRow.Item("foundation_depth"), Double)
            Else
                Me.foundation_depth = Nothing
            End If
        Catch
            Me.foundation_depth = Nothing
        End Try 'Foundation_Depth
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pad_thickness"), Double)) Then
                Me.pad_thickness = CType(PileDataRow.Item("pad_thickness"), Double)
            Else
                Me.pad_thickness = Nothing
            End If
        Catch
            Me.pad_thickness = Nothing
        End Try 'Pad_Thickness
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pad_width_dir1"), Double)) Then
                Me.pad_width_dir1 = CType(PileDataRow.Item("pad_width_dir1"), Double)
            Else
                Me.pad_width_dir1 = Nothing
            End If
        Catch
            Me.pad_width_dir1 = Nothing
        End Try 'Pad_Width_Dir1
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pad_width_dir2"), Double)) Then
                Me.pad_width_dir2 = CType(PileDataRow.Item("pad_width_dir2"), Double)
            Else
                Me.pad_width_dir2 = Nothing
            End If
        Catch
            Me.pad_width_dir2 = Nothing
        End Try 'Pad_Width_Dir2
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pad_rebar_size_bottom"), Integer)) Then
                Me.pad_rebar_size_bottom = CType(PileDataRow.Item("pad_rebar_size_bottom"), Integer)
            Else
                Me.pad_rebar_size_bottom = Nothing
            End If
        Catch
            Me.pad_rebar_size_bottom = Nothing
        End Try 'Pad_Rebar_Size_Bottom
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pad_rebar_size_top"), Integer)) Then
                Me.pad_rebar_size_top = CType(PileDataRow.Item("pad_rebar_size_top"), Integer)
            Else
                Me.pad_rebar_size_top = Nothing
            End If
        Catch
            Me.pad_rebar_size_top = Nothing
        End Try 'Pad_Rebar_Size_Top
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pad_rebar_quantity_bottom_dir1"), Integer)) Then
                Me.pad_rebar_quantity_bottom_dir1 = CType(PileDataRow.Item("pad_rebar_quantity_bottom_dir1"), Integer)
            Else
                Me.pad_rebar_quantity_bottom_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_bottom_dir1 = Nothing
        End Try 'Pad_Rebar_Quantity_Bottom_Dir1
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pad_rebar_quantity_top_dir1"), Integer)) Then
                Me.pad_rebar_quantity_top_dir1 = CType(PileDataRow.Item("pad_rebar_quantity_top_dir1"), Integer)
            Else
                Me.pad_rebar_quantity_top_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_top_dir1 = Nothing
        End Try 'Pad_Rebar_Quantity_Top_Dir1
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pad_rebar_quantity_bottom_dir2"), Integer)) Then
                Me.pad_rebar_quantity_bottom_dir2 = CType(PileDataRow.Item("pad_rebar_quantity_bottom_dir2"), Integer)
            Else
                Me.pad_rebar_quantity_bottom_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_bottom_dir2 = Nothing
        End Try 'Pad_Rebar_Quantity_Bottom_Dir2
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pad_rebar_quantity_top_dir2"), Integer)) Then
                Me.pad_rebar_quantity_top_dir2 = CType(PileDataRow.Item("pad_rebar_quantity_top_dir2"), Integer)
            Else
                Me.pad_rebar_quantity_top_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_top_dir2 = Nothing
        End Try 'Pad_Rebar_Quantity_Top_Dir2
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pier_shape"), String)) Then
                Me.pier_shape = CType(PileDataRow.Item("pier_shape"), String)
            Else
                Me.pier_shape = ""
            End If
        Catch
            Me.pier_shape = ""
        End Try 'Pier_Shape
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pier_diameter"), Integer)) Then
                Me.pier_diameter = CType(PileDataRow.Item("pier_diameter"), Integer)
            Else
                Me.pier_diameter = Nothing
            End If
        Catch
            Me.pier_diameter = Nothing
        End Try 'Pier_Diameter
        Try
            If Not IsDBNull(CType(PileDataRow.Item("extension_above_grade"), Double)) Then
                Me.extension_above_grade = CType(PileDataRow.Item("extension_above_grade"), Double)
            Else
                Me.extension_above_grade = Nothing
            End If
        Catch
            Me.extension_above_grade = Nothing
        End Try 'Extension_Above_Grade
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pier_rebar_size"), Integer)) Then
                Me.pier_rebar_size = CType(PileDataRow.Item("pier_rebar_size"), Integer)
            Else
                Me.pier_rebar_size = Nothing
            End If
        Catch
            Me.pier_rebar_size = Nothing
        End Try 'Pier_Rebar_Size
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pier_rebar_quantity"), Integer)) Then
                Me.pier_rebar_quantity = CType(PileDataRow.Item("pier_rebar_quantity"), Integer)
            Else
                Me.pier_rebar_quantity = Nothing
            End If
        Catch
            Me.pier_rebar_quantity = Nothing
        End Try 'Pier_Rebar_Quantity
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pier_tie_size"), Integer)) Then
                Me.pier_tie_size = CType(PileDataRow.Item("pier_tie_size"), Integer)
            Else
                Me.pier_tie_size = Nothing
            End If
        Catch
            Me.pier_tie_size = Nothing
        End Try 'Pier_Tie_Size
        'Try
        '    If Not IsDBNull(CType(PileDataRow.Item("pier_tie_quantity"), Integer)) Then
        '        Me.pier_tie_quantity = CType(PileDataRow.Item("pier_tie_quantity"), Integer)
        '    Else
        '        Me.pier_tie_quantity = Nothing
        '    End If
        'Catch
        '    Me.pier_tie_quantity = Nothing
        'End Try 'Pier_Tie_Quantity
        Try
            If Not IsDBNull(CType(PileDataRow.Item("rebar_grade"), Double)) Then
                Me.rebar_grade = CType(PileDataRow.Item("rebar_grade"), Double)
            Else
                Me.rebar_grade = Nothing
            End If
        Catch
            Me.rebar_grade = Nothing
        End Try 'Rebar_Grade
        Try
            If Not IsDBNull(CType(PileDataRow.Item("concrete_compressive_strength"), Double)) Then
                Me.concrete_compressive_strength = CType(PileDataRow.Item("concrete_compressive_strength"), Double)
            Else
                Me.concrete_compressive_strength = Nothing
            End If
        Catch
            Me.concrete_compressive_strength = Nothing
        End Try 'Concrete_Compressive_Strength
        Try
            If Not IsDBNull(CType(PileDataRow.Item("groundwater_depth"), Double)) Then
                Me.groundwater_depth = CType(PileDataRow.Item("groundwater_depth"), Double)
            Else
                Me.groundwater_depth = Nothing
            End If
        Catch
            Me.groundwater_depth = Nothing
        End Try 'Groundwater_Depth
        Try
            If Not IsDBNull(CType(PileDataRow.Item("total_soil_unit_weight"), Double)) Then
                Me.total_soil_unit_weight = CType(PileDataRow.Item("total_soil_unit_weight"), Double)
            Else
                Me.total_soil_unit_weight = Nothing
            End If
        Catch
            Me.total_soil_unit_weight = Nothing
        End Try 'Total_Soil_Unit_Weight
        Try
            If Not IsDBNull(CType(PileDataRow.Item("cohesion"), Double)) Then
                Me.cohesion = CType(PileDataRow.Item("cohesion"), Double)
            Else
                Me.cohesion = Nothing
            End If
        Catch
            Me.cohesion = Nothing
        End Try 'Cohesion
        Try
            If Not IsDBNull(CType(PileDataRow.Item("friction_angle"), Double)) Then
                Me.friction_angle = CType(PileDataRow.Item("friction_angle"), Double)
            Else
                Me.friction_angle = Nothing
            End If
        Catch
            Me.friction_angle = Nothing
        End Try 'Friction_Angle
        Try
            If Not IsDBNull(CType(PileDataRow.Item("neglect_depth"), Double)) Then
                Me.neglect_depth = CType(PileDataRow.Item("neglect_depth"), Double)
            Else
                Me.neglect_depth = Nothing
            End If
        Catch
            Me.neglect_depth = Nothing
        End Try 'Neglect_Depth
        Try
            If Not IsDBNull(CType(PileDataRow.Item("spt_blow_count"), Integer)) Then
                Me.spt_blow_count = CType(PileDataRow.Item("spt_blow_count"), Integer)
            Else
                Me.spt_blow_count = Nothing
            End If
        Catch
            Me.spt_blow_count = Nothing
        End Try 'Spt_Blow_Count
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_negative_friction_force"), Double)) Then
                Me.pile_negative_friction_force = CType(PileDataRow.Item("pile_negative_friction_force"), Double)
            Else
                Me.pile_negative_friction_force = Nothing
            End If
        Catch
            Me.pile_negative_friction_force = Nothing
        End Try 'Pile_Negative_Friction_Force
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_ultimate_compression"), Double)) Then
                Me.pile_ultimate_compression = CType(PileDataRow.Item("pile_ultimate_compression"), Double)
            Else
                Me.pile_ultimate_compression = Nothing
            End If
        Catch
            Me.pile_ultimate_compression = Nothing
        End Try 'Pile_Ultimate_Compression
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_ultimate_tension"), Double)) Then
                Me.pile_ultimate_tension = CType(PileDataRow.Item("pile_ultimate_tension"), Double)
            Else
                Me.pile_ultimate_tension = Nothing
            End If
        Catch
            Me.pile_ultimate_tension = Nothing
        End Try 'Pile_Ultimate_Tension
        Try
            Me.top_and_bottom_rebar_different = CType(PileDataRow.Item("top_and_bottom_rebar_different"), Boolean)
        Catch
            Me.top_and_bottom_rebar_different = False
        End Try 'Top_And_Bottom_Rebar_Different
        Try
            If Not IsDBNull(CType(PileDataRow.Item("ultimate_gross_end_bearing"), Double)) Then
                Me.ultimate_gross_end_bearing = CType(PileDataRow.Item("ultimate_gross_end_bearing"), Double)
            Else
                Me.ultimate_gross_end_bearing = Nothing
            End If
        Catch
            Me.ultimate_gross_end_bearing = Nothing
        End Try 'Ultimate_Gross_End_Bearing
        Try
            Me.skin_friction_given = CType(PileDataRow.Item("skin_friction_given"), Boolean)
        Catch
            Me.skin_friction_given = False
        End Try 'Skin_Friction_Given
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_quantity_circular"), Integer)) Then
                Me.pile_quantity_circular = CType(PileDataRow.Item("pile_quantity_circular"), Integer)
            Else
                Me.pile_quantity_circular = Nothing
            End If
        Catch
            Me.pile_quantity_circular = Nothing
        End Try 'Pile_Quantity_Circular
        Try
            If Not IsDBNull(CType(PileDataRow.Item("group_diameter_circular"), Double)) Then
                Me.group_diameter_circular = CType(PileDataRow.Item("group_diameter_circular"), Double)
            Else
                Me.group_diameter_circular = Nothing
            End If
        Catch
            Me.group_diameter_circular = Nothing
        End Try 'Group_Diameter_Circular
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_column_quantity"), Integer)) Then
                Me.pile_column_quantity = CType(PileDataRow.Item("pile_column_quantity"), Integer)
            Else
                Me.pile_column_quantity = Nothing
            End If
        Catch
            Me.pile_column_quantity = Nothing
        End Try 'Pile_Column_Quantity
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_row_quantity"), Integer)) Then
                Me.pile_row_quantity = CType(PileDataRow.Item("pile_row_quantity"), Integer)
            Else
                Me.pile_row_quantity = Nothing
            End If
        Catch
            Me.pile_row_quantity = Nothing
        End Try 'Pile_Row_Quantity
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_columns_spacing"), Double)) Then
                Me.pile_columns_spacing = CType(PileDataRow.Item("pile_columns_spacing"), Double)
            Else
                Me.pile_columns_spacing = Nothing
            End If
        Catch
            Me.pile_columns_spacing = Nothing
        End Try 'Pile_Columns_Spacing
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_row_spacing"), Double)) Then
                Me.pile_row_spacing = CType(PileDataRow.Item("pile_row_spacing"), Double)
            Else
                Me.pile_row_spacing = Nothing
            End If
        Catch
            Me.pile_row_spacing = Nothing
        End Try 'Pile_Row_Spacing
        Try
            Me.group_efficiency_factor_given = CType(PileDataRow.Item("group_efficiency_factor_given"), Boolean)
        Catch
            Me.group_efficiency_factor_given = False
        End Try 'Group_Efficiency_Factor_Given
        Try
            If Not IsDBNull(CType(PileDataRow.Item("group_efficiency_factor"), Double)) Then
                Me.group_efficiency_factor = CType(PileDataRow.Item("group_efficiency_factor"), Double)
            Else
                Me.group_efficiency_factor = Nothing
            End If
        Catch
            Me.group_efficiency_factor = Nothing
        End Try 'Group_Efficiency_Factor
        Try
            Me.cap_type = CType(PileDataRow.Item("cap_type"), String)
        Catch
            Me.cap_type = ""
        End Try 'Cap_Type
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_quantity_asymmetric"), Integer)) Then
                Me.pile_quantity_asymmetric = CType(PileDataRow.Item("pile_quantity_asymmetric"), Integer)
            Else
                Me.pile_quantity_asymmetric = Nothing
            End If
        Catch
            Me.pile_quantity_asymmetric = Nothing
        End Try 'Pile_Quantity_Asymmetric
        Try
            If Not IsDBNull(CType(PileDataRow.Item("pile_spacing_min_asymmetric"), Double)) Then
                Me.pile_spacing_min_asymmetric = CType(PileDataRow.Item("pile_spacing_min_asymmetric"), Double)
            Else
                Me.pile_spacing_min_asymmetric = Nothing
            End If
        Catch
            Me.pile_spacing_min_asymmetric = Nothing
        End Try 'Pile_Spacing_Min_Asymmetric
        Try
            If Not IsDBNull(CType(PileDataRow.Item("quantity_piles_surrounding"), Integer)) Then
                Me.quantity_piles_surrounding = CType(PileDataRow.Item("quantity_piles_surrounding"), Integer)
            Else
                Me.quantity_piles_surrounding = Nothing
            End If
        Catch
            Me.quantity_piles_surrounding = Nothing
        End Try 'Quantity_Piles_Surrounding

        For Each SoilLayerDataRow As DataRow In ds.Tables("Pile Soil SQL").Rows
            Dim soilRefID As Integer = CType(SoilLayerDataRow.Item("pile_fnd_id"), Integer)
            If soilRefID = refID Then
                Me.soil_layers.Add(New PileSoilLayer(SoilLayerDataRow))
            End If
        Next 'Add Soild Layers to to Pile Soil Layer Object

        For Each LocationDataRow As DataRow In ds.Tables("Pile Location SQL").Rows
            Dim locRefID As Integer = CType(LocationDataRow.Item("pile_fnd_id"), Integer)
            If locRefID = refID Then
                Me.pile_locations.Add(New PileLocation(LocationDataRow))
            End If
        Next 'Add Soild Layers to to Pile Location Object

    End Sub 'Generate from EDS

    Public Sub New(ByVal path As String)
        Try
            Me.pile_id = CType(GetOneExcelRange(path, "ID"), Integer)
        Catch
            Me.pile_id = 0
        End Try 'Pile_Id
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Ecc"), Double)) Then
                Me.load_eccentricity = CType(GetOneExcelRange(path, "Ecc"), Double)
            Else
                Me.load_eccentricity = Nothing
            End If
        Catch
            Me.load_eccentricity = Nothing
        End Try 'Load_Eccentricity
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "BC"), Double)) Then
                Me.bolt_circle_bearing_plate_width = CType(GetOneExcelRange(path, "BC"), Double)
            Else
                Me.bolt_circle_bearing_plate_width = Nothing
            End If
        Catch
            Me.bolt_circle_bearing_plate_width = Nothing
        End Try 'Bolt_Circle_Bearing_Plate_Width
        Try
            Me.pile_shape = CType(GetOneExcelRange(path, "D23", "Input"), String)
        Catch
            Me.pile_shape = ""
        End Try 'Pile_Shape
        Try
            Me.pile_material = CType(GetOneExcelRange(path, "D24", "Input"), String)
        Catch
            Me.pile_material = ""
        End Try 'Pile_Material
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Lpile"), Double)) Then
                Me.pile_length = CType(GetOneExcelRange(path, "Lpile"), Double)
            Else
                Me.pile_length = Nothing
            End If
        Catch
            Me.pile_length = Nothing
        End Try 'Pile_Length
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D26", "Input"), Double)) Then
                Me.pile_diameter_width = CType(GetOneExcelRange(path, "D26", "Input"), Double)
            Else
                Me.pile_diameter_width = Nothing
            End If
        Catch
            Me.pile_diameter_width = Nothing
        End Try 'Pile_Diameter_Width
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D27", "Input"), Double)) Then
                Me.pile_pipe_thickness = CType(GetOneExcelRange(path, "D27", "Input"), Double)
            Else
                Me.pile_pipe_thickness = Nothing
            End If
        Catch
            Me.pile_pipe_thickness = Nothing
        End Try 'Pile_Pipe_Thickness
        Try
            If CType(GetOneExcelRange(path, "D29", "Input"), String) = "Yes" Then
                Me.pile_soil_capacity_given = True
            Else
                Me.pile_soil_capacity_given = False
            End If
        Catch
            Me.pile_soil_capacity_given = False
        End Try 'Pile_Soil_Capacity_Given
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D30", "Input"), Double)) Then
                Me.steel_yield_strength = CType(GetOneExcelRange(path, "D30", "Input"), Double)
            Else
                Me.steel_yield_strength = Nothing
            End If
        Catch
            Me.steel_yield_strength = Nothing
        End Try 'Steel_Yield_Strength
        Try
            Me.pile_type_option = CType(GetOneExcelRange(path, "Psize"), String)
        Catch
            Me.pile_type_option = ""
        End Try 'Pile_Type_Option
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Pquan"), Integer)) Then
                Me.rebar_quantity = CType(GetOneExcelRange(path, "Pquan"), Integer)
            Else
                Me.rebar_quantity = Nothing
            End If
        Catch
            Me.rebar_quantity = Nothing
        End Try 'Rebar_Quantity
        Try
            Me.pile_group_config = CType(GetOneExcelRange(path, "Config"), String)
        Catch
            Me.pile_group_config = ""
        End Try 'Pile_Group_Config
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D"), Double)) Then
                Me.foundation_depth = CType(GetOneExcelRange(path, "D"), Double)
            Else
                Me.foundation_depth = Nothing
            End If
        Catch
            Me.foundation_depth = Nothing
        End Try 'Foundation_Depth
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "T"), Double)) Then
                Me.pad_thickness = CType(GetOneExcelRange(path, "T"), Double)
            Else
                Me.pad_thickness = Nothing
            End If
        Catch
            Me.pad_thickness = Nothing
        End Try 'Pad_Thickness
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Wx"), Double)) Then
                Me.pad_width_dir1 = CType(GetOneExcelRange(path, "Wx"), Double)
            Else
                Me.pad_width_dir1 = Nothing
            End If
        Catch
            Me.pad_width_dir1 = Nothing
        End Try 'Pad_Width_Dir1
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Wy"), Double)) Then
                Me.pad_width_dir2 = CType(GetOneExcelRange(path, "Wy"), Double)
            Else
                Me.pad_width_dir2 = Nothing
            End If
        Catch
            Me.pad_width_dir2 = Nothing
        End Try 'Pad_Width_Dir2
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Spad"), Integer)) Then
                Me.pad_rebar_size_bottom = CType(GetOneExcelRange(path, "Spad"), Integer)
            Else
                Me.pad_rebar_size_bottom = Nothing
            End If
        Catch
            Me.pad_rebar_size_bottom = Nothing
        End Try 'Pad_Rebar_Size_Bottom
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Spad_top"), Integer)) Then
                Me.pad_rebar_size_top = CType(GetOneExcelRange(path, "Spad_top"), Integer)
            Else
                Me.pad_rebar_size_top = Nothing
            End If
        Catch
            Me.pad_rebar_size_top = Nothing
        End Try 'Pad_Rebar_Size_Top
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Mpad"), Integer)) Then
                Me.pad_rebar_quantity_bottom_dir1 = CType(GetOneExcelRange(path, "Mpad"), Integer)
            Else
                Me.pad_rebar_quantity_bottom_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_bottom_dir1 = Nothing
        End Try 'Pad_Rebar_Quantity_Bottom_Dir1
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Mpad_top"), Integer)) Then
                Me.pad_rebar_quantity_top_dir1 = CType(GetOneExcelRange(path, "Mpad_top"), Integer)
            Else
                Me.pad_rebar_quantity_top_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_top_dir1 = Nothing
        End Try 'Pad_Rebar_Quantity_Top_Dir1
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Mpad_y"), Integer)) Then
                Me.pad_rebar_quantity_bottom_dir2 = CType(GetOneExcelRange(path, "Mpad_y"), Integer)
            Else
                Me.pad_rebar_quantity_bottom_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_bottom_dir2 = Nothing
        End Try 'Pad_Rebar_Quantity_Bottom_Dir2
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Mpad_y_top"), Integer)) Then
                Me.pad_rebar_quantity_top_dir2 = CType(GetOneExcelRange(path, "Mpad_y_top"), Integer)
            Else
                Me.pad_rebar_quantity_top_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_top_dir2 = Nothing
        End Try 'Pad_Rebar_Quantity_Top_Dir2
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D57", "Input"), String)) Then
                Me.pier_shape = CType(GetOneExcelRange(path, "D57", "Input"), String)
            Else
                Me.pier_shape = ""
            End If
        Catch
            Me.pier_shape = ""
        End Try 'Pier_Shape
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "di"), Integer)) Then
                Me.pier_diameter = CType(GetOneExcelRange(path, "di"), Integer)
            Else
                Me.pier_diameter = Nothing
            End If
        Catch
            Me.pier_diameter = Nothing
        End Try 'Pier_Diameter
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "E"), Double)) Then
                Me.extension_above_grade = CType(GetOneExcelRange(path, "E"), Double)
            Else
                Me.extension_above_grade = Nothing
            End If
        Catch
            Me.extension_above_grade = Nothing
        End Try 'Extension_Above_Grade
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Rs"), Integer)) Then
                Me.pier_rebar_size = CType(GetOneExcelRange(path, "Rs"), Integer)
            Else
                Me.pier_rebar_size = Nothing
            End If
        Catch
            Me.pier_rebar_size = Nothing
        End Try 'Pier_Rebar_Size
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "mc"), Integer)) Then
                Me.pier_rebar_quantity = CType(GetOneExcelRange(path, "mc"), Integer)
            Else
                Me.pier_rebar_quantity = Nothing
            End If
        Catch
            Me.pier_rebar_quantity = Nothing
        End Try 'Pier_Rebar_Quantity
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "St"), Integer)) Then
                Me.pier_tie_size = CType(GetOneExcelRange(path, "St"), Integer)
            Else
                Me.pier_tie_size = Nothing
            End If
        Catch
            Me.pier_tie_size = Nothing
        End Try 'Pier_Tie_Size
        'Try
        '    If Not IsNothing(CType(GetOneExcelRange(path, "", ""), Integer)) Then
        '        Me.pier_tie_quantity = CType(GetOneExcelRange(path, "", ""), Integer)
        '    Else
        '        Me.pier_tie_quantity = Nothing
        '    End If
        'Catch
        '    Me.pier_tie_quantity = Nothing
        'End Try 'Pier_Tie_Quantity
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Fy"), Double)) Then
                Me.rebar_grade = CType(GetOneExcelRange(path, "Fy"), Double)
            Else
                Me.rebar_grade = Nothing
            End If
        Catch
            Me.rebar_grade = Nothing
        End Try 'Rebar_Grade
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Fc"), Double)) Then
                Me.concrete_compressive_strength = CType(GetOneExcelRange(path, "Fc"), Double)
            Else
                Me.concrete_compressive_strength = Nothing
            End If
        Catch
            Me.concrete_compressive_strength = Nothing
        End Try 'Concrete_Compressive_Strength
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D69", "Input"), Double)) Then
                Me.groundwater_depth = CType(GetOneExcelRange(path, "D69", "Input"), Double)
            Else
                Me.groundwater_depth = Nothing
            End If
        Catch
            Me.groundwater_depth = Nothing
        End Try 'Groundwater_Depth
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "γsoil_dry"), Double)) Then
                Me.total_soil_unit_weight = CType(GetOneExcelRange(path, "γsoil_dry"), Double)
            Else
                Me.total_soil_unit_weight = Nothing
            End If
        Catch
            Me.total_soil_unit_weight = Nothing
        End Try 'Total_Soil_Unit_Weight
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Co"), Double)) Then
                Me.cohesion = CType(GetOneExcelRange(path, "Co"), Double)
            Else
                Me.cohesion = Nothing
            End If
        Catch
            Me.cohesion = Nothing
        End Try 'Cohesion
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "ɸ"), Double)) Then
                Me.friction_angle = CType(GetOneExcelRange(path, "ɸ"), Double)
            Else
                Me.friction_angle = Nothing
            End If
        Catch
            Me.friction_angle = Nothing
        End Try 'Friction_Angle
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "ND"), Double)) Then
                Me.neglect_depth = CType(GetOneExcelRange(path, "ND"), Double)
            Else
                Me.neglect_depth = Nothing
            End If
        Catch
            Me.neglect_depth = Nothing
        End Try 'Neglect_Depth
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "N_blows"), Integer)) Then
                Me.spt_blow_count = CType(GetOneExcelRange(path, "N_blows"), Integer)
            Else
                Me.spt_blow_count = Nothing
            End If
        Catch
            Me.spt_blow_count = Nothing
        End Try 'Spt_Blow_Count
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "Sw"), Double)) Then
                Me.pile_negative_friction_force = CType(GetOneExcelRange(path, "Sw"), Double)
            Else
                Me.pile_negative_friction_force = Nothing
            End If
        Catch
            Me.pile_negative_friction_force = Nothing
        End Try 'Pile_Negative_Friction_Force
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "K45", "Input"), Double)) Then
                Me.pile_ultimate_compression = CType(GetOneExcelRange(path, "K45", "Input"), Double)
            Else
                Me.pile_ultimate_compression = Nothing
            End If
        Catch
            Me.pile_ultimate_compression = Nothing
        End Try 'Pile_Ultimate_Compression
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "K46", "Input"), Double)) Then
                Me.pile_ultimate_tension = CType(GetOneExcelRange(path, "K46", "Input"), Double)
            Else
                Me.pile_ultimate_tension = Nothing
            End If
        Catch
            Me.pile_ultimate_tension = Nothing
        End Try 'Pile_Ultimate_Tension
        Try
            Me.top_and_bottom_rebar_different = CType(GetOneExcelRange(path, "Z10", "Input"), Boolean)

        Catch
            Me.top_and_bottom_rebar_different = False
        End Try 'Top_And_Bottom_Rebar_Different
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "M71", "Input"), Double)) Then
                Me.ultimate_gross_end_bearing = CType(GetOneExcelRange(path, "M71", "Input"), Double)
            Else
                Me.ultimate_gross_end_bearing = Nothing
            End If
        Catch
            Me.ultimate_gross_end_bearing = Nothing
        End Try 'Ultimate_Gross_End_Bearing
        Try
            If CType(GetOneExcelRange(path, "N54", "Input"), String) = "Yes" Then
                Me.skin_friction_given = True
            Else
                Me.skin_friction_given = False
            End If
        Catch
            Me.skin_friction_given = False
        End Try 'Skin_Friction_Given
        If Me.pile_group_config = "Circular" Then
            Try
                If Not IsNothing(CType(GetOneExcelRange(path, "D36", "Input"), Integer)) Then
                    Me.pile_quantity_circular = CType(GetOneExcelRange(path, "D36", "Input"), Integer)
                Else
                    Me.pile_quantity_circular = Nothing
                End If
            Catch
                Me.pile_quantity_circular = Nothing
            End Try 'Pile_Quantity_Circular
            Try
                If Not IsNothing(CType(GetOneExcelRange(path, "D37", "Input"), Double)) Then
                    Me.group_diameter_circular = CType(GetOneExcelRange(path, "D37", "Input"), Double)
                Else
                    Me.group_diameter_circular = Nothing
                End If
            Catch
                Me.group_diameter_circular = Nothing
            End Try 'Group_Diameter_Circular
        End If
        If Me.pile_group_config = "Rectangular" Then
            Try
                If Not IsNothing(CType(GetOneExcelRange(path, "D36", "Input"), Integer)) Then
                    Me.pile_column_quantity = CType(GetOneExcelRange(path, "D36", "Input"), Integer)
                Else
                    Me.pile_column_quantity = Nothing
                End If
            Catch
                Me.pile_column_quantity = Nothing
            End Try 'Pile_Column_Quantity
            Try
                If Not IsNothing(CType(GetOneExcelRange(path, "D37", "Input"), Integer)) Then
                    Me.pile_row_quantity = CType(GetOneExcelRange(path, "D37", "Input"), Integer)
                Else
                    Me.pile_row_quantity = Nothing
                End If
            Catch
                Me.pile_row_quantity = Nothing
            End Try 'Pile_Row_Quantity
        End If
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D38", "Input"), Double)) Then
                Me.pile_columns_spacing = CType(GetOneExcelRange(path, "D38", "Input"), Double)
            Else
                Me.pile_columns_spacing = Nothing
            End If
        Catch
            Me.pile_columns_spacing = Nothing
        End Try 'Pile_Columns_Spacing
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D39", "Input"), Double)) Then
                Me.pile_row_spacing = CType(GetOneExcelRange(path, "D39", "Input"), Double)
            Else
                Me.pile_row_spacing = Nothing
            End If
        Catch
            Me.pile_row_spacing = Nothing
        End Try 'Pile_Row_Spacing
        Try
            If CType(GetOneExcelRange(path, "D41", "Input"), String) = "Yes" Then
                Me.group_efficiency_factor_given = True
            Else
                Me.group_efficiency_factor_given = False
            End If
        Catch
            Me.group_efficiency_factor_given = False
        End Try 'Group_Efficiency_Factor_Given
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D42", "Input"), Double)) Then
                Me.group_efficiency_factor = CType(GetOneExcelRange(path, "D42", "Input"), Double)
            Else
                Me.group_efficiency_factor = Nothing
            End If
        Catch
            Me.group_efficiency_factor = Nothing
        End Try 'Group_Efficiency_Factor
        Try
            Me.cap_type = CType(GetOneExcelRange(path, "D45", "Input"), String)
        Catch
            Me.cap_type = ""
        End Try 'Cap_Type
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D10", "Moment of Inertia"), Integer)) Then
                Me.pile_quantity_asymmetric = CType(GetOneExcelRange(path, "D10", "Moment of Inertia"), Integer)
            Else
                Me.pile_quantity_asymmetric = Nothing
            End If
        Catch
            Me.pile_quantity_asymmetric = Nothing
        End Try 'Pile_Quantity_Asymmetric
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D11", "Moment of Inertia"), Double)) Then
                Me.pile_spacing_min_asymmetric = CType(GetOneExcelRange(path, "D11", "Moment of Inertia"), Double)
            Else
                Me.pile_spacing_min_asymmetric = Nothing
            End If
        Catch
            Me.pile_spacing_min_asymmetric = Nothing
        End Try 'Pile_Spacing_Min_Asymmetric
        Try
            If Not IsNothing(CType(GetOneExcelRange(path, "D12", "Moment of Inertia"), Integer)) Then
                Me.quantity_piles_surrounding = CType(GetOneExcelRange(path, "D12", "Moment of Inertia"), Integer)
            Else
                Me.quantity_piles_surrounding = Nothing
            End If
        Catch
            Me.quantity_piles_surrounding = Nothing
        End Try 'Quantity_Piles_Surrounding

        For Each SoilLayerDataRow As DataRow In ds.Tables("Pile Soil EXCEL").Rows
            Me.soil_layers.Add(New PileSoilLayer(SoilLayerDataRow))
        Next 'Add Soil Layers to to Pile Soil Layer Object

        For Each LocationDataRow As DataRow In ds.Tables("Pile Location EXCEL").Rows
            Me.pile_locations.Add(New PileLocation(LocationDataRow))
        Next 'Add Location to to Pile Location Object

    End Sub 'Generate from Excel
#End Region

End Class

Partial Public Class PileLocation

#Region "Define"
    Private prop_location_id As Integer
    Private prop_pile_x_coordinate As Double?
    Private prop_pile_y_coordinate As Double?
    <Category("Pile Location"), Description(""), DisplayName("Location_Id")>
    Public Property location_id() As Integer
        Get
            Return Me.prop_location_id
        End Get
        Set
            Me.prop_location_id = Value
        End Set
    End Property
    <Category("Pile Location"), Description(""), DisplayName("Pile_X_Coordinate")>
    Public Property pile_x_coordinate() As Double?
        Get
            Return Me.prop_pile_x_coordinate
        End Get
        Set
            Me.prop_pile_x_coordinate = Value
        End Set
    End Property
    <Category("Pile Location"), Description(""), DisplayName("Pile_Y_Coordinate")>
    Public Property pile_y_coordinate() As Double?
        Get
            Return Me.prop_pile_y_coordinate
        End Get
        Set
            Me.prop_pile_y_coordinate = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal LocationDataRow As DataRow)
        Try
            Me.location_id = CType(LocationDataRow.Item("location_id"), Integer)
        Catch
            Me.location_id = 0
        End Try 'Location_Id
        Try
            If Not IsDBNull(CType(LocationDataRow.Item("pile_x_coordinate"), Double)) Then
                Me.pile_x_coordinate = CType(LocationDataRow.Item("pile_x_coordinate"), Double)
            Else
                Me.pile_x_coordinate = Nothing
            End If
        Catch
            Me.pile_x_coordinate = Nothing
        End Try 'Pile_X_Coordinate
        Try
            If Not IsDBNull(CType(LocationDataRow.Item("pile_y_coordinate"), Double)) Then
                Me.pile_y_coordinate = CType(LocationDataRow.Item("pile_y_coordinate"), Double)
            Else
                Me.pile_y_coordinate = Nothing
            End If
        Catch
            Me.pile_y_coordinate = Nothing
        End Try 'Pile_Y_Coordinate
    End Sub 'Add a pile location to a pile
#End Region

End Class
Partial Public Class PileSoilLayer

#Region "Define"
    Private prop_soil_layer_id As Integer
    Private prop_bottom_depth As Double?
    Private prop_effective_soil_density As Double?
    Private prop_cohesion As Double?
    Private prop_friction_angle As Double?
    'Private prop_skin_friction_override_uplift As Double?
    Private prop_spt_blow_count As Integer?
    Private prop_ultimate_skin_friction_comp As Double?
    Private prop_ultimate_skin_friction_uplift As Double?
    <Category("Pile Soil Layer"), Description(""), DisplayName("Soil_Layer_Id")>
    Public Property soil_layer_id() As Integer
        Get
            Return Me.prop_soil_layer_id
        End Get
        Set
            Me.prop_soil_layer_id = Value
        End Set
    End Property
    <Category("Pile Soil Layer"), Description(""), DisplayName("Bottom_Depth")>
    Public Property bottom_depth() As Double?
        Get
            Return Me.prop_bottom_depth
        End Get
        Set
            Me.prop_bottom_depth = Value
        End Set
    End Property
    <Category("Pile Soil Layer"), Description(""), DisplayName("Effective_Soil_Density")>
    Public Property effective_soil_density() As Double?
        Get
            Return Me.prop_effective_soil_density
        End Get
        Set
            Me.prop_effective_soil_density = Value
        End Set
    End Property
    <Category("Pile Soil Layer"), Description(""), DisplayName("Cohesion")>
    Public Property cohesion() As Double?
        Get
            Return Me.prop_cohesion
        End Get
        Set
            Me.prop_cohesion = Value
        End Set
    End Property
    <Category("Pile Soil Layer"), Description(""), DisplayName("Friction_Angle")>
    Public Property friction_angle() As Double?
        Get
            Return Me.prop_friction_angle
        End Get
        Set
            Me.prop_friction_angle = Value
        End Set
    End Property
    '<Category("Pile Soil Layer"), Description(""), DisplayName("Skin_Friction_Override_Uplift")>
    'Public Property skin_friction_override_uplift() As Double?
    '    Get
    '        Return Me.prop_skin_friction_override_uplift
    '    End Get
    '    Set
    '        Me.prop_skin_friction_override_uplift = Value
    '    End Set
    'End Property
    <Category("Pile Soil Layer"), Description(""), DisplayName("Spt_Blow_Count")>
    Public Property spt_blow_count() As Integer?
        Get
            Return Me.prop_spt_blow_count
        End Get
        Set
            Me.prop_spt_blow_count = Value
        End Set
    End Property
    <Category("Pile Soil Layer"), Description(""), DisplayName("Ultimate_Skin_Friction_Comp")>
    Public Property ultimate_skin_friction_comp() As Double?
        Get
            Return Me.prop_ultimate_skin_friction_comp
        End Get
        Set
            Me.prop_ultimate_skin_friction_comp = Value
        End Set
    End Property
    <Category("Pile Soil Layer"), Description(""), DisplayName("Ultimate_Skin_Friction_Uplift")>
    Public Property ultimate_skin_friction_uplift() As Double?
        Get
            Return Me.prop_ultimate_skin_friction_uplift
        End Get
        Set
            Me.prop_ultimate_skin_friction_uplift = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave Method Empty
    End Sub

    Public Sub New(ByVal SoilLayerDataRow As DataRow)
        Try
            Me.soil_layer_id = CType(SoilLayerDataRow.Item("soil_layer_id"), Integer)
        Catch
            Me.soil_layer_id = 0
        End Try 'Soil_Layer_Id
        Try
            If Not IsDBNull(CType(SoilLayerDataRow.Item("bottom_depth"), Double)) Then
                Me.bottom_depth = CType(SoilLayerDataRow.Item("bottom_depth"), Double)
            Else
                Me.bottom_depth = Nothing
            End If
        Catch
            Me.bottom_depth = Nothing
        End Try 'Bottom_Depth
        Try
            If Not IsDBNull(CType(SoilLayerDataRow.Item("effective_soil_density"), Double)) Then
                Me.effective_soil_density = CType(SoilLayerDataRow.Item("effective_soil_density"), Double)
            Else
                Me.effective_soil_density = Nothing
            End If
        Catch
            Me.effective_soil_density = Nothing
        End Try 'Effective_Soil_Density
        Try
            If Not IsDBNull(CType(SoilLayerDataRow.Item("cohesion"), Double)) Then
                Me.cohesion = CType(SoilLayerDataRow.Item("cohesion"), Double)
            Else
                Me.cohesion = Nothing
            End If
        Catch
            Me.cohesion = Nothing
        End Try 'Cohesion
        Try
            If Not IsDBNull(CType(SoilLayerDataRow.Item("friction_angle"), Double)) Then
                Me.friction_angle = CType(SoilLayerDataRow.Item("friction_angle"), Double)
            Else
                Me.friction_angle = Nothing
            End If
        Catch
            Me.friction_angle = Nothing
        End Try 'Friction_Angle
        'Try
        '    If Not IsDBNull(CType(SoilLayerDataRow.Item("skin_friction_override_uplift"), Double)) Then
        '        Me.skin_friction_override_uplift = CType(SoilLayerDataRow.Item("skin_friction_override_uplift"), Double)
        '    Else
        '        Me.skin_friction_override_uplift = Nothing
        '    End If
        'Catch
        '    Me.skin_friction_override_uplift = Nothing
        'End Try 'Skin_Friction_Override_Uplift
        Try
            If Not IsDBNull(CType(SoilLayerDataRow.Item("spt_blow_count"), Integer)) Then
                Me.spt_blow_count = CType(SoilLayerDataRow.Item("spt_blow_count"), Integer)
            Else
                Me.spt_blow_count = Nothing
            End If
        Catch
            Me.spt_blow_count = Nothing
        End Try 'Spt_Blow_Count
        Try
            If Not IsDBNull(CType(SoilLayerDataRow.Item("ultimate_skin_friction_comp"), Double)) Then
                Me.ultimate_skin_friction_comp = CType(SoilLayerDataRow.Item("ultimate_skin_friction_comp"), Double)
            Else
                Me.ultimate_skin_friction_comp = Nothing
            End If
        Catch
            Me.ultimate_skin_friction_comp = Nothing
        End Try 'Ultimate_Skin_Friction_Comp
        Try
            If Not IsDBNull(CType(SoilLayerDataRow.Item("ultimate_skin_friction_uplift"), Double)) Then
                Me.ultimate_skin_friction_uplift = CType(SoilLayerDataRow.Item("ultimate_skin_friction_uplift"), Double)
            Else
                Me.ultimate_skin_friction_uplift = Nothing
            End If
        Catch
            Me.ultimate_skin_friction_uplift = Nothing
        End Try 'Ultimate_Skin_Friction_Uplift
    End Sub 'Add a soil layer to a pile

#End Region

End Class