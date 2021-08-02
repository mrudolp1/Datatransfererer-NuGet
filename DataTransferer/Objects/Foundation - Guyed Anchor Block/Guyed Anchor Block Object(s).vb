Option Strict On
Imports System.ComponentModel

Partial Public Class GuyedAnchorBlock

#Region "Define"

    Private prop_anchor_id As Integer
    Private prop_local_anchor_id As Integer
    Private prop_anchor_location As String
    Private prop_guy_anchor_radius As Double
    Private prop_anchor_depth As Double
    Private prop_anchor_width As Double
    Private prop_anchor_thickness As Double
    Private prop_anchor_length As Double
    Private prop_anchor_toe_width As Double
    Private prop_anchor_top_rebar_size As Integer
    Private prop_anchor_top_rebar_quantity As Integer
    Private prop_anchor_bottom_rebar_size As Integer
    Private prop_anchor_bottom_rebar_quantity As Integer
    Private prop_anchor_stirrup_size As Integer
    Private prop_anchor_shaft_diameter As Double
    Private prop_anchor_shaft_quantity As Integer
    Private prop_anchor_shaft_area_override As Double
    Private prop_anchor_shaft_shear_lag_factor As Double
    Private prop_anchor_rebar_grade As Double
    Private prop_concrete_compressive_strength As Double
    Private prop_clear_cover As Double
    Private prop_anchor_shaft_yield_strength As Double
    Private prop_anchor_shaft_ultimate_strength As Double
    Private prop_neglect_depth As Double
    Private prop_groundwater_depth As Double
    Private prop_soil_layer_quantity As Integer
    Public Property soil_layers As New List(Of GuyedAnchorBlockSoilLayer)

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Global ID")>
    Public Property anchor_id() As Integer
        Get
            Return Me.prop_anchor_id
        End Get
        Set
            Me.prop_anchor_id = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Local ID")>
    Public Property local_anchor_id() As Integer
        Get
            Return Me.prop_local_anchor_id
        End Get
        Set
            Me.prop_local_anchor_id = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Location")>
    Public Property anchor_location() As String
        Get
            Return Me.prop_anchor_location
        End Get
        Set
            Me.prop_anchor_location = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Radius")>
    Public Property guy_anchor_radius() As Double
        Get
            Return Me.prop_guy_anchor_radius
        End Get
        Set
            Me.prop_guy_anchor_radius = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Depth")>
    Public Property anchor_depth() As Double
        Get
            Return Me.prop_anchor_depth
        End Get
        Set
            Me.prop_anchor_depth = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Width")>
    Public Property anchor_width() As Double
        Get
            Return Me.prop_anchor_width
        End Get
        Set
            Me.prop_anchor_width = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Thickness")>
    Public Property anchor_thickness() As Double
        Get
            Return Me.prop_anchor_thickness
        End Get
        Set
            Me.prop_anchor_thickness = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Length")>
    Public Property anchor_length() As Double
        Get
            Return Me.prop_anchor_length
        End Get
        Set
            Me.prop_anchor_length = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Toe Width")>
    Public Property anchor_toe_width() As Double
        Get
            Return Me.prop_anchor_toe_width
        End Get
        Set
            Me.prop_anchor_toe_width = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Top Rebar Size")>
    Public Property anchor_top_rebar_size() As Integer
        Get
            Return Me.prop_anchor_top_rebar_size
        End Get
        Set
            Me.prop_anchor_top_rebar_size = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Top Rebar Quantity")>
    Public Property anchor_top_rebar_quantity() As Integer
        Get
            Return Me.prop_anchor_top_rebar_quantity
        End Get
        Set
            Me.prop_anchor_top_rebar_quantity = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Bottom Rebar Size")>
    Public Property anchor_bottom_rebar_size() As Integer
        Get
            Return Me.prop_anchor_bottom_rebar_size
        End Get
        Set
            Me.prop_anchor_bottom_rebar_size = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Bottom Rebar Quantity")>
    Public Property anchor_bottom_rebar_quantity() As Integer
        Get
            Return Me.prop_anchor_bottom_rebar_quantity
        End Get
        Set
            Me.prop_anchor_bottom_rebar_quantity = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Stirrup Size")>
    Public Property anchor_stirrup_size() As Integer
        Get
            Return Me.prop_anchor_stirrup_size
        End Get
        Set
            Me.prop_anchor_stirrup_size = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Diameter")>
    Public Property anchor_shaft_diameter() As Double
        Get
            Return Me.prop_anchor_shaft_diameter
        End Get
        Set
            Me.prop_anchor_shaft_diameter = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Quantity")>
    Public Property anchor_shaft_quantity() As Integer
        Get
            Return Me.prop_anchor_shaft_quantity
        End Get
        Set
            Me.prop_anchor_shaft_quantity = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Area Override")>
    Public Property anchor_shaft_area_override() As Double
        Get
            Return Me.prop_anchor_shaft_area_override
        End Get
        Set
            Me.prop_anchor_shaft_area_override = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Shear Lag Factor")>
    Public Property anchor_shaft_shear_lag_factor() As Double
        Get
            Return Me.prop_anchor_shaft_shear_lag_factor
        End Get
        Set
            Me.prop_anchor_shaft_shear_lag_factor = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Guyed Anchor Block Rebar Grade")>
    Public Property anchor_rebar_grade() As Double
        Get
            Return Me.prop_anchor_rebar_grade
        End Get
        Set
            Me.prop_anchor_rebar_grade = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Concrete Compressive Strength")>
    Public Property concrete_compressive_strength() As Double
        Get
            Return Me.prop_concrete_compressive_strength
        End Get
        Set
            Me.prop_concrete_compressive_strength = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Clear Cover")>
    Public Property clear_cover() As Double
        Get
            Return Me.prop_clear_cover
        End Get
        Set
            Me.prop_clear_cover = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Yield Strength")>
    Public Property anchor_shaft_yield_strength() As Double
        Get
            Return Me.prop_anchor_shaft_yield_strength
        End Get
        Set
            Me.prop_anchor_shaft_yield_strength = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Anchor Shaft Ultimate Strength")>
    Public Property anchor_shaft_ultimate_strength() As Double
        Get
            Return Me.prop_anchor_shaft_ultimate_strength
        End Get
        Set
            Me.prop_anchor_shaft_ultimate_strength = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Neglect Depth")>
    Public Property neglect_depth() As Double
        Get
            Return Me.prop_neglect_depth
        End Get
        Set
            Me.prop_neglect_depth = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Groundwater Depth")>
    Public Property groundwater_depth() As Double
        Get
            Return Me.prop_groundwater_depth
        End Get
        Set
            Me.prop_groundwater_depth = Value
        End Set
    End Property

    <Category("Guyed Anchor Block Details"), Description(""), DisplayName("Soil Layer Quantity")>
    Public Property soil_layer_quantity() As Integer
        Get
            Return Me.prop_soil_layer_quantity
        End Get
        Set
            Me.prop_soil_layer_quantity = Value
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
            Me.anchor_id = CType(GuyedAnchorBlockDataRow.Item("guyed_anchor_block_id"), Integer)
        Catch
            Me.anchor_id = 0
        End Try 'Guyed Anchor Block ID
        Try
            Me.anchor_location = CType(GuyedAnchorBlockDataRow.Item("anchor_location"), String)
        Catch
            Me.anchor_location = ""
        End Try 'Guyed Anchor Block Location
        Try
            Me.guy_anchor_radius = CType(GuyedAnchorBlockDataRow.Item("guy_anchor_radius"), Double)
        Catch
            Me.guy_anchor_radius = -1 'Set stored value of -1 to "" (Empty) in tool
        End Try 'Guyed Anchor Block Radius
        Try
            Me.anchor_depth = CType(GuyedAnchorBlockDataRow.Item("anchor_depth"), Double)
        Catch
            Me.anchor_depth = 0
        End Try 'Guyed Anchor Block Depth
        Try
            Me.anchor_width = CType(GuyedAnchorBlockDataRow.Item("anchor_width"), Double)
        Catch
            Me.anchor_width = 0
        End Try 'Guyed Anchor Block Width
        Try
            Me.anchor_thickness = CType(GuyedAnchorBlockDataRow.Item("anchor_thickness"), Double)
        Catch
            Me.anchor_thickness = 0
        End Try 'Guyed Anchor Block Thickness
        Try
            Me.anchor_length = CType(GuyedAnchorBlockDataRow.Item("anchor_length"), Double)
        Catch
            Me.anchor_length = 0
        End Try 'Guyed Anchor Block Length
        Try
            Me.anchor_toe_width = CType(GuyedAnchorBlockDataRow.Item("anchor_toe_width"), Double)
        Catch
            Me.anchor_toe_width = 0
        End Try 'Guyed Anchor Block Toe Width
        Try
            Me.anchor_top_rebar_size = CType(GuyedAnchorBlockDataRow.Item("anchor_top_rebar_size"), Integer)
        Catch
            Me.anchor_top_rebar_size = 3
        End Try 'Guyed Anchor Block Top Rebar Size
        Try
            Me.anchor_top_rebar_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_top_rebar_quantity"), Integer)
        Catch
            Me.anchor_top_rebar_quantity = 0
        End Try 'Guyed Anchor Block Top Rebar Quantity
        Try
            Me.anchor_bottom_rebar_size = CType(GuyedAnchorBlockDataRow.Item("anchor_bottom_rebar_size"), Integer)
        Catch
            Me.anchor_bottom_rebar_size = 3
        End Try 'Guyed Anchor Block Bottom Rebar Size
        Try
            Me.anchor_bottom_rebar_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_bottom_rebar_quantity"), Integer)
        Catch
            Me.anchor_bottom_rebar_quantity = 0
        End Try 'Guyed Anchor Block Bottom Rebar Quantity
        Try
            Me.anchor_stirrup_size = CType(GuyedAnchorBlockDataRow.Item("anchor_stirrup_size"), Integer)
        Catch
            Me.anchor_stirrup_size = 4
        End Try 'Guyed Anchor Block Stirrup Size
        Try
            Me.anchor_shaft_diameter = CType(GuyedAnchorBlockDataRow.Item("anchor_stirrup_size"), Double)
        Catch
            Me.anchor_shaft_diameter = 0
        End Try 'Guyed Anchor Block Anchor Shaft Diameter
        Try
            Me.anchor_shaft_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_quantity"), Integer)
        Catch
            Me.anchor_shaft_quantity = 1
        End Try 'Guyed Anchor Block Anchor Shaft Quantity
        Try
            Me.anchor_shaft_area_override = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_area_override"), Double)
        Catch
            Me.anchor_shaft_area_override = -1 'Set stored value of -1 to "" (Empty) in tool
        End Try 'Guyed Anchor Block Anchor Area Override
        Try
            Me.anchor_shaft_shear_lag_factor = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_shear_lag_factor"), Double)
        Catch
            Me.anchor_shaft_shear_lag_factor = 1
        End Try 'Guyed Anchor Block Anchor Shear Lag Factor
        Try
            Me.anchor_rebar_grade = CType(GuyedAnchorBlockDataRow.Item("anchor_rebar_grade"), Double)
        Catch
            Me.anchor_rebar_grade = 60
        End Try 'Guyed Anchor Block Rebar Grade
        Try
            Me.concrete_compressive_strength = CType(GuyedAnchorBlockDataRow.Item("concrete_compressive_strength"), Double)
        Catch
            Me.concrete_compressive_strength = 3
        End Try 'Guyed Anchor Block Concrete Compressive Strength
        Try
            Me.clear_cover = CType(GuyedAnchorBlockDataRow.Item("clear_cover"), Double)
        Catch
            Me.clear_cover = 3
        End Try 'Guyed Anchor Block Clear Cover
        Try
            Me.anchor_shaft_yield_strength = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_yield_strength"), Double)
        Catch
            Me.anchor_shaft_yield_strength = 50
        End Try 'Guyed Anchor Block Anchor Shaft Yield Strength
        Try
            Me.anchor_shaft_ultimate_strength = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_ultimate_strength"), Double)
        Catch
            Me.anchor_shaft_ultimate_strength = 65
        End Try 'Guyed Anchor Block Anchor Shaft Ultimate Strength
        Try
            Me.neglect_depth = CType(GuyedAnchorBlockDataRow.Item("neglect_depth"), Double)
        Catch
            Me.neglect_depth = 0
        End Try 'Guyed Anchor Block Anchor Neglect Depth
        Try
            Me.groundwater_depth = CType(GuyedAnchorBlockDataRow.Item("groundwater_depth"), Double)
        Catch
            Me.groundwater_depth = -1 'Set stored value of -1 to "N/A" in tool
        End Try 'Guyed Anchor Block Anchor Groundwater Depth
        Try
            Me.soil_layer_quantity = CType(GuyedAnchorBlockDataRow.Item("soil_layer_quantity"), Integer)
        Catch
            Me.soil_layer_quantity = 1
        End Try 'Guyed Anchor Block Anchor Soil Layer Quantity

        For Each SoilLayerDataRow As DataRow In ds.Tables("Guyed Anchor Block Soil SQL").Rows
            Dim soilRefID As Integer = CType(SoilLayerDataRow.Item("anchor_id"), Integer)

            If soilRefID = refID Then
                Me.soil_layers.Add(New GuyedAnchorBlockSoilLayer(SoilLayerDataRow))
            End If
        Next 'Add Soil Layers to to Guyed Anchor Block Object



    End Sub 'Generate a guyed anchor block from EDS

    Public Sub New(ByVal GuyedAnchorBlockDataRow As DataRow, ByVal refID As Integer, ByVal refcol As String, ByVal constants As List(Of EXCELRngParameter))
        'General Guyed Anchor Block Details
        Try
            Me.anchor_id = CType(GuyedAnchorBlockDataRow.Item("guyed_anchor_block_id"), Integer)
        Catch
            Me.anchor_id = 0
        End Try 'Guyed Anchor Block ID
        Try
            Me.anchor_location = CType(GuyedAnchorBlockDataRow.Item("anchor_location"), String)
        Catch
            Me.anchor_location = ""
        End Try 'Guyed Anchor Block Location
        Try
            Me.guy_anchor_radius = CType(GuyedAnchorBlockDataRow.Item("guy_anchor_radius"), Double)
        Catch
            Me.guy_anchor_radius = -1 'Set stored value of -1 to "" (Empty) in tool
        End Try 'Guyed Anchor Block Radius
        Try
            Me.anchor_depth = CType(GuyedAnchorBlockDataRow.Item("anchor_depth"), Double)
        Catch
            Me.anchor_depth = 0
        End Try 'Guyed Anchor Block Depth
        Try
            Me.anchor_width = CType(GuyedAnchorBlockDataRow.Item("anchor_width"), Double)
        Catch
            Me.anchor_width = 0
        End Try 'Guyed Anchor Block Width
        Try
            Me.anchor_thickness = CType(GuyedAnchorBlockDataRow.Item("anchor_thickness"), Double)
        Catch
            Me.anchor_thickness = 0
        End Try 'Guyed Anchor Block Thickness
        Try
            Me.anchor_length = CType(GuyedAnchorBlockDataRow.Item("anchor_length"), Double)
        Catch
            Me.anchor_length = 0
        End Try 'Guyed Anchor Block Length
        Try
            Me.anchor_toe_width = CType(GuyedAnchorBlockDataRow.Item("anchor_toe_width"), Double)
        Catch
            Me.anchor_toe_width = 0
        End Try 'Guyed Anchor Block Toe Width
        Try
            Me.anchor_top_rebar_size = CType(GuyedAnchorBlockDataRow.Item("anchor_top_rebar_size"), Integer)
        Catch
            Me.anchor_top_rebar_size = 3
        End Try 'Guyed Anchor Block Top Rebar Size
        Try
            Me.anchor_top_rebar_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_top_rebar_quantity"), Integer)
        Catch
            Me.anchor_top_rebar_quantity = 0
        End Try 'Guyed Anchor Block Top Rebar Quantity
        Try
            Me.anchor_bottom_rebar_size = CType(GuyedAnchorBlockDataRow.Item("anchor_bottom_rebar_size"), Integer)
        Catch
            Me.anchor_bottom_rebar_size = 3
        End Try 'Guyed Anchor Block Bottom Rebar Size
        Try
            Me.anchor_bottom_rebar_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_bottom_rebar_quantity"), Integer)
        Catch
            Me.anchor_bottom_rebar_quantity = 0
        End Try 'Guyed Anchor Block Bottom Rebar Quantity
        Try
            Me.anchor_stirrup_size = CType(GuyedAnchorBlockDataRow.Item("anchor_stirrup_size"), Integer)
        Catch
            Me.anchor_stirrup_size = 4
        End Try 'Guyed Anchor Block Stirrup Size
        Try
            Me.anchor_shaft_diameter = CType(GuyedAnchorBlockDataRow.Item("anchor_stirrup_size"), Double)
        Catch
            Me.anchor_shaft_diameter = 0
        End Try 'Guyed Anchor Block Anchor Shaft Diameter
        Try
            Me.anchor_shaft_quantity = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_quantity"), Integer)
        Catch
            Me.anchor_shaft_quantity = 1
        End Try 'Guyed Anchor Block Anchor Shaft Quantity
        Try
            Me.anchor_shaft_area_override = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_area_override"), Double)
        Catch
            Me.anchor_shaft_area_override = -1 'Set stored value of -1 to "" (Empty) in tool
        End Try 'Guyed Anchor Block Anchor Area Override
        Try
            Me.anchor_shaft_shear_lag_factor = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_shear_lag_factor"), Double)
        Catch
            Me.anchor_shaft_shear_lag_factor = 1
        End Try 'Guyed Anchor Block Anchor Shear Lag Factor
        Try
            Me.anchor_rebar_grade = CType(GuyedAnchorBlockDataRow.Item("anchor_rebar_grade"), Double)
        Catch
            Me.anchor_rebar_grade = 60
        End Try 'Guyed Anchor Block Rebar Grade
        Try
            Me.concrete_compressive_strength = CType(GuyedAnchorBlockDataRow.Item("concrete_compressive_strength"), Double)
        Catch
            Me.concrete_compressive_strength = 3
        End Try 'Guyed Anchor Block Concrete Compressive Strength
        Try
            Me.clear_cover = CType(GuyedAnchorBlockDataRow.Item("clear_cover"), Double)
        Catch
            Me.clear_cover = 3
        End Try 'Guyed Anchor Block Clear Cover
        Try
            Me.anchor_shaft_yield_strength = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_yield_strength"), Double)
        Catch
            Me.anchor_shaft_yield_strength = 50
        End Try 'Guyed Anchor Block Anchor Shaft Yield Strength
        Try
            Me.anchor_shaft_ultimate_strength = CType(GuyedAnchorBlockDataRow.Item("anchor_shaft_ultimate_strength"), Double)
        Catch
            Me.anchor_shaft_ultimate_strength = 65
        End Try 'Guyed Anchor Block Anchor Shaft Ultimate Strength
        Try
            Me.neglect_depth = CType(GuyedAnchorBlockDataRow.Item("neglect_depth"), Double)
        Catch
            Me.neglect_depth = 0
        End Try 'Guyed Anchor Block Anchor Neglect Depth
        Try
            Me.groundwater_depth = CType(GuyedAnchorBlockDataRow.Item("groundwater_depth"), Double)
        Catch
            Me.groundwater_depth = -1 'Set stored value of -1 to "N/A" in tool
        End Try 'Guyed Anchor Block Anchor Groundwater Depth
        Try
            Me.soil_layer_quantity = CType(GuyedAnchorBlockDataRow.Item("soil_layer_quantity"), Integer)
        Catch
            Me.soil_layer_quantity = 1
        End Try 'Guyed Anchor Block Anchor Soil Layer Quantity

        For Each SoilLayerDataRow As DataRow In ds.Tables("Guyed Anchor Block Soil EXCEL").Rows
            Dim soilRefID As Integer = CType(SoilLayerDataRow.Item(refcol), Integer)

            If soilRefID = refID Then
                Me.soil_layers.Add(New GuyedAnchorBlockSoilLayer(SoilLayerDataRow))
            End If
        Next 'Add Soil Layers to to Guyed Anchor Block Object
    End Sub 'Generate a guyed anchor block from Excel

#End Region

End Class


#Region "Guyed Anchor Block Extras"

Partial Public Class GuyedAnchorBlockSoilLayer

    Private prop_soil_layer_id As Integer
    Private prop_bottom_depth As Double
    Private prop_effective_soil_density As Double
    Private prop_cohesion As Double
    Private prop_friction_angle As Double
    Private prop_skin_friction_override_uplift As Double
    Private prop_spt_blow_count As Integer

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
    Public Property bottom_depth() As Double
        Get
            Return Me.prop_bottom_depth
        End Get
        Set
            Me.prop_bottom_depth = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil Effective Soil Density"), Description(""), DisplayName("Effective Soil Density")>
    Public Property effective_soil_density() As Double
        Get
            Return Me.prop_effective_soil_density
        End Get
        Set
            Me.prop_effective_soil_density = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil Cohesion"), Description(""), DisplayName("Soil Cohesion")>
    Public Property cohesion() As Double
        Get
            Return Me.prop_cohesion
        End Get
        Set
            Me.prop_cohesion = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil Friction Angle"), Description(""), DisplayName("Soil Friction Angle")>
    Public Property friction_angle() As Double
        Get
            Return Me.prop_friction_angle
        End Get
        Set
            Me.prop_friction_angle = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil Skin Friction Override"), Description(""), DisplayName("Soil Skin Friction Override")>
    Public Property skin_friction_override_uplift() As Double
        Get
            Return Me.prop_skin_friction_override_uplift
        End Get
        Set
            Me.prop_skin_friction_override_uplift = Value
        End Set
    End Property
    <Category("Guyed Anchor Block Soil SPT Blow Count"), Description(""), DisplayName("Soil SPT Blow Count")>
    Public Property spt_blow_count() As Integer
        Get
            Return Me.prop_spt_blow_count
        End Get
        Set
            Me.prop_spt_blow_count = Value
        End Set
    End Property

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal SoilLayerDataRow As DataRow)
        Try
            Me.soil_layer_id = CType(SoilLayerDataRow.Item("soil_layer_id"), Integer)
        Catch
            Me.soil_layer_id = 1
        End Try 'Soil Layer ID
        Try
            Me.bottom_depth = CType(SoilLayerDataRow.Item("bottom_depth"), Double)
        Catch
            Me.bottom_depth = 0
        End Try 'Soil Layer Depth
        Try
            Me.effective_soil_density = CType(SoilLayerDataRow.Item("effective_soil_density"), Double)
        Catch
            Me.effective_soil_density = 0
        End Try 'Soil Layer Effective Density
        Try
            Me.cohesion = CType(SoilLayerDataRow.Item("cohesion"), Double)
        Catch
            Me.cohesion = 0
        End Try 'Soil Layer Cohesion
        Try
            Me.friction_angle = CType(SoilLayerDataRow.Item("friction_angle"), Double)
        Catch
            Me.friction_angle = Nothing
        End Try 'Soil Layer Friction Angle
        Try
            Me.skin_friction_override_uplift = CType(SoilLayerDataRow.Item("skin_friction_override_uplift"), Double)
        Catch
            'Me.skin_friction_override_uplift = Nothing (WRITE CODE ON BACKEND TO DEFAULT NULL VALUE AS BLANK IN TOOL)
            Console.WriteLine("Skin friction is null")
        End Try 'Soil Layer Skin Friction Override
        Try
            Me.spt_blow_count = CType(SoilLayerDataRow.Item("spt_blow_count"), Integer)
        Catch
            'Me.spt_blow_count = Nothing (WRITE CODE ON BACKEND TO DEFAULT NULL VALUE AS BLANK IN TOOL)
            Console.WriteLine("SPT Blow Count is null")
        End Try 'Soil Layer Blow Count
    End Sub

End Class

#End Region
