﻿Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet

Partial Public Class UnitBase
    Inherits EDSFoundation

#Region "Inheritted"
    '''Must override these inherited properties

    Public Overrides ReadOnly Property EDSObjectName As String = "Unit Base"
    Public Overrides ReadOnly Property foundationType As String = "Unit Base"
    Public Overrides ReadOnly Property EDSTableName As String = "fnd.unit_base"
    Public Overrides ReadOnly Property templatePath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "SST Unit Base Foundation.xlsm")
    Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
        Get
            Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("Unit Base General Details EXCEL", "A1:AT2", "Details (SAPI)")}
        End Get
    End Property
    Private _Insert As String
    Private _Update As String
    Private _Delete As String
    Public Overrides ReadOnly Property Insert() As String
        Get
            If _Insert = "" Then
                _Insert = QueryBuilderFromFile(queryPath & "Unit Base\Unit Base (INSERT).sql")
            End If
            Dim InsertString As String = _Insert
            InsertString = InsertString.Replace("[BU NUMBER]", Me.bus_unit.FormatDBValue)
            InsertString = InsertString.Replace("[STRUCTURE ID]", Me.structure_id.FormatDBValue)
            InsertString = InsertString.Replace("[FOUNDATION VALUES]", Me.SQLInsertValues)
            InsertString = InsertString.Replace("[FOUNDATION FIELDS]", Me.SQLInsertFields)
            Return InsertString
        End Get
    End Property

    Public Overrides ReadOnly Property Update() As String
        Get
            If _Update = "" Then
                _Update = QueryBuilderFromFile(queryPath & "Unit Base\Unit Base (UPDATE).sql")
            End If
            Dim UpdateString As String = _Update
            UpdateString = UpdateString.Replace("[ID]", Me.ID.ToString.FormatDBValue)
            UpdateString = UpdateString.Replace("[UPDATE]", Me.SQLUpdate)
            Return UpdateString
        End Get
    End Property

    Public Overrides ReadOnly Property Delete() As String
        Get
            If _Delete = "" Then
                _Delete = QueryBuilderFromFile(queryPath & "Unit Base\Unit Base (DELETE).sql")
            End If
            Dim DeleteString As String = _Delete
            DeleteString = DeleteString.Replace("[ID]", Me.ID.ToString.FormatDBValue)
            Return DeleteString
        End Get
    End Property

#End Region
#Region "Define"
    'General
    Private prop_extension_above_grade As Double?
    Private prop_foundation_depth As Double?
    Private prop_concrete_compressive_strength As Double?
    Private prop_dry_concrete_density As Double?
    Private prop_rebar_grade As Double?
    Private prop_top_and_bottom_rebar_different As Boolean?
    Private prop_block_foundation As Boolean?
    Private prop_rectangular_foundation As Boolean?
    Private prop_base_plate_distance_above_foundation As Double?
    Private prop_bolt_circle_bearing_plate_width As Double?
    Private prop_tower_centroid_offset As Boolean? '**
    Private prop_basic_soil_check As Boolean?
    Private prop_structural_check As Boolean?
    Private prop_tool_version As String
    Private prop_modified As Boolean?
    'Pier
    Private prop_pier_shape As String
    Private prop_pier_diameter As Double?
    Private prop_pier_rebar_quantity As Double?
    Private prop_pier_rebar_size As Integer?
    Private prop_pier_tie_quantity As Double?
    Private prop_pier_tie_size As Integer?
    Private prop_pier_reinforcement_type As String
    Private prop_pier_clear_cover As Double?
    'Pad
    Private prop_pad_width_1 As Double?
    Private prop_pad_width_2 As Double?
    Private prop_pad_thickness As Double?
    Private prop_pad_rebar_size_top_dir1 As Integer?
    Private prop_pad_rebar_size_bottom_dir1 As Integer?
    Private prop_pad_rebar_size_top_dir2 As Integer?
    Private prop_pad_rebar_size_bottom_dir2 As Integer?
    Private prop_pad_rebar_quantity_top_dir1 As Double?
    Private prop_pad_rebar_quantity_bottom_dir1 As Double?
    Private prop_pad_rebar_quantity_top_dir2 As Double?
    Private prop_pad_rebar_quantity_bottom_dir2 As Double?
    Private prop_pad_clear_cover As Double?
    'Soil
    Private prop_total_soil_unit_weight As Double?
    Private prop_bearing_type As String
    Private prop_nominal_bearing_capacity As Double?
    Private prop_cohesion As Double?
    Private prop_friction_angle As Double?
    Private prop_spt_blow_count As Double?
    Private prop_base_friction_factor As Double?
    Private prop_neglect_depth As Double?
    Private prop_bearing_distribution_type As Boolean?
    Private prop_groundwater_depth As Double?
    'Non FND Specific Inputs
    'BU
    'Site Name
    'App Number
    'Private prop_tia_current As String ' TIA
    'Section 15.5
    'Load Z
    'Tower Height
    'Base Face Width
    'BP Dist Above FND?
    'AR Bolt Circle?
    'Extension Above Grade? - FND can be source
    'SDC

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
    Public Property top_and_bottom_rebar_different() As Boolean?
        Get
            Return Me.prop_top_and_bottom_rebar_different
        End Get
        Set
            Me.prop_top_and_bottom_rebar_different = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Block Foundation")>
    Public Property block_foundation() As Boolean?
        Get
            Return Me.prop_block_foundation
        End Get
        Set
            Me.prop_block_foundation = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Rectangular Foundation")>
    Public Property rectangular_foundation() As Boolean?
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
    Public Property tower_centroid_offset() As Boolean?
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
    Public Property pier_rebar_quantity() As Double?
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
    Public Property pier_tie_quantity() As Double?
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
    Public Property pad_rebar_quantity_top_dir1() As Double?
        Get
            Return Me.prop_pad_rebar_quantity_top_dir1
        End Get
        Set
            Me.prop_pad_rebar_quantity_top_dir1 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Quantity Bottom Direction 1")>
    Public Property pad_rebar_quantity_bottom_dir1() As Double?
        Get
            Return Me.prop_pad_rebar_quantity_bottom_dir1
        End Get
        Set
            Me.prop_pad_rebar_quantity_bottom_dir1 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Quantity Top Direction 2")>
    Public Property pad_rebar_quantity_top_dir2() As Double?
        Get
            Return Me.prop_pad_rebar_quantity_top_dir2
        End Get
        Set
            Me.prop_pad_rebar_quantity_top_dir2 = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Pad Rebar Quantity Bottom Direction 2")>
    Public Property pad_rebar_quantity_bottom_dir2() As Double?
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
    Public Property spt_blow_count() As Double?
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
    Public Property bearing_distribution_type() As Boolean?
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
    Public Property basic_soil_check() As Boolean?
        Get
            Return Me.prop_basic_soil_check
        End Get
        Set
            Me.prop_basic_soil_check = Value
        End Set
    End Property
    <Category("Unit Base Details"), Description(""), DisplayName("Structural Checks up to 105% Acceptable?")>
    Public Property structural_check() As Boolean?
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
    <Category("Unit Base Details"), Description(""), DisplayName("Modified")>
    Public Property modified() As Boolean?
        Get
            Return Me.prop_modified
        End Get
        Set
            Me.prop_modified = Value
        End Set
    End Property

    'Non FND Specific Inputs
    'BU
    'Site Name
    'App Number
    '<Category("Unit Base Details"), Description(""), DisplayName("TIA")>
    'Public Property tia_current() As String
    '    Get
    '        Return If(Me.ParentStructure.structureCodeCriteria.tia_current, Me.prop_tia_current)
    '    End Get
    '    Set
    '        Me.prop_tia_current = Value
    '    End Set
    'End Property
    'Section 15.5
    'Load Z
    'Tower Height
    'Base Face Width
    'BP Dist Above FND?
    'AR Bolt Circle?
    'Extension Above Grade? - FND can be source
    'SDC
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal ubDr As DataRow, Optional ByRef Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        ''''''Customize for each foundation type'''''

        Me.ID = DBtoNullableInt(ubDr.Item("ID"))
        Me.bus_unit = DBtoStr(ubDr.Item("bus_unit"))
        Me.structure_id = DBtoStr(ubDr.Item("structure_id"))

        Try
            If Not IsDBNull(CType(ubDr.Item("extension_above_grade"), Double)) Then
                Me.extension_above_grade = CType(ubDr.Item("extension_above_grade"), Double)
            Else
                Me.extension_above_grade = Nothing
            End If
        Catch
            Me.extension_above_grade = Nothing
        End Try 'Extension Above Grade
        Try
            If Not IsDBNull(CType(ubDr.Item("foundation_depth"), Double)) Then
                Me.foundation_depth = CType(ubDr.Item("foundation_depth"), Double)
            Else
                Me.foundation_depth = Nothing
            End If
        Catch
            Me.foundation_depth = Nothing
        End Try 'Foundation Depth
        Try
            If Not IsDBNull(CType(ubDr.Item("concrete_compressive_strength"), Double)) Then
                Me.concrete_compressive_strength = CType(ubDr.Item("concrete_compressive_strength"), Double)
            Else
                Me.concrete_compressive_strength = Nothing
            End If
        Catch
            Me.concrete_compressive_strength = Nothing
        End Try 'Concrete Compressive Strength
        Try
            If Not IsDBNull(CType(ubDr.Item("dry_concrete_density"), Double)) Then
                Me.dry_concrete_density = CType(ubDr.Item("dry_concrete_density"), Double)
            Else
                Me.dry_concrete_density = Nothing
            End If
        Catch
            Me.dry_concrete_density = Nothing
        End Try 'Dry Concrete Density
        Try
            If Not IsDBNull(CType(ubDr.Item("rebar_grade"), Double)) Then
                Me.rebar_grade = CType(ubDr.Item("rebar_grade"), Double)
            Else
                Me.rebar_grade = Nothing
            End If
        Catch
            Me.rebar_grade = Nothing
        End Try 'Rebar Grade
        Try
            Me.top_and_bottom_rebar_different = CType(ubDr.Item("top_and_bottom_rebar_different"), Boolean)
        Catch
            Me.top_and_bottom_rebar_different = False
        End Try 'Top and Bottom Rebar Different
        Try
            Me.block_foundation = CType(ubDr.Item("block_foundation"), Boolean)
        Catch
            Me.block_foundation = False
        End Try 'Block Foundation 
        Try
            Me.rectangular_foundation = CType(ubDr.Item("rectangular_foundation"), Boolean)
        Catch
            Me.rectangular_foundation = False
        End Try 'Rectangular Foundation
        Try
            If Not IsDBNull(CType(ubDr.Item("base_plate_distance_above_foundation"), Double)) Then
                Me.base_plate_distance_above_foundation = CType(ubDr.Item("base_plate_distance_above_foundation"), Double)
            Else
                Me.base_plate_distance_above_foundation = Nothing
            End If
        Catch
            Me.base_plate_distance_above_foundation = Nothing
        End Try 'Base Plate Distance Above Foundation
        Try
            If Not IsDBNull(CType(ubDr.Item("bolt_circle_bearing_plate_width"), Double)) Then
                Me.bolt_circle_bearing_plate_width = CType(ubDr.Item("bolt_circle_bearing_plate_width"), Double)
            Else
                Me.bolt_circle_bearing_plate_width = Nothing
            End If
        Catch
            Me.bolt_circle_bearing_plate_width = Nothing
        End Try 'Bolt Circle Bearing Plate Width
        Try
            If Not IsDBNull(CType(ubDr.Item("tower_centroid_offset"), Boolean)) Then
                Me.tower_centroid_offset = CType(ubDr.Item("tower_centroid_offset"), Boolean)
            Else
                Me.tower_centroid_offset = Nothing
            End If
        Catch
            Me.tower_centroid_offset = Nothing
        End Try 'Tower Centroid Offset
        Try
            Me.pier_shape = CType(ubDr.Item("pier_shape"), String)
        Catch
            Me.pier_shape = ""
        End Try 'Pier Shape
        Try
            If Not IsDBNull(CType(ubDr.Item("pier_diameter"), Double)) Then
                Me.pier_diameter = CType(ubDr.Item("pier_diameter"), Double)
            Else
                Me.pier_diameter = Nothing
            End If
        Catch
            Me.pier_diameter = Nothing
        End Try 'Pier Diameter
        Try
            If Not IsDBNull(CType(ubDr.Item("pier_rebar_quantity"), Double)) Then
                Me.pier_rebar_quantity = CType(ubDr.Item("pier_rebar_quantity"), Double)
            Else
                Me.pier_rebar_quantity = Nothing
            End If
        Catch
            Me.pier_rebar_quantity = Nothing
        End Try 'Pier Rebar Quantity
        Try
            If Not IsDBNull(CType(ubDr.Item("pier_rebar_size"), Integer)) Then
                Me.pier_rebar_size = CType(ubDr.Item("pier_rebar_size"), Integer)
            Else
                Me.pier_rebar_size = Nothing
            End If
        Catch
            Me.pier_rebar_size = Nothing
        End Try 'Pier Rebar Size
        Try
            If Not IsDBNull(CType(ubDr.Item("pier_tie_quantity"), Double)) Then
                Me.pier_tie_quantity = CType(ubDr.Item("pier_tie_quantity"), Double)
            Else
                Me.pier_tie_quantity = Nothing
            End If
        Catch
            Me.pier_tie_quantity = Nothing
        End Try 'Pier Tie Quantity
        Try
            If Not IsDBNull(CType(ubDr.Item("pier_tie_size"), Integer)) Then
                Me.pier_tie_size = CType(ubDr.Item("pier_tie_size"), Integer)
            Else
                Me.pier_tie_size = Nothing
            End If
        Catch
            Me.pier_tie_size = Nothing
        End Try 'Pier Tie Size
        Try
            Me.pier_reinforcement_type = CType(ubDr.Item("pier_reinforcement_type"), String)
        Catch
            Me.pier_reinforcement_type = ""
        End Try 'Pier Reinforcement Type
        Try
            If Not IsDBNull(CType(ubDr.Item("pier_clear_cover"), Double)) Then
                Me.pier_clear_cover = CType(ubDr.Item("pier_clear_cover"), Double)
            Else
                Me.pier_clear_cover = Nothing
            End If
        Catch
            Me.pier_clear_cover = Nothing
        End Try 'Pier Clear Cover
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_width_1"), Double)) Then
                Me.pad_width_1 = CType(ubDr.Item("pad_width_1"), Double)
            Else
                Me.pad_width_1 = Nothing
            End If
        Catch
            Me.pad_width_1 = Nothing
        End Try 'Pad Width 1
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_width_2"), Double)) Then
                Me.pad_width_2 = CType(ubDr.Item("pad_width_2"), Double)
            Else
                Me.pad_width_2 = Nothing
            End If
        Catch
            Me.pad_width_2 = Nothing
        End Try 'Pad Width 2
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_thickness"), Double)) Then
                Me.pad_thickness = CType(ubDr.Item("pad_thickness"), Double)
            Else
                Me.pad_thickness = Nothing
            End If
        Catch
            Me.pad_thickness = Nothing
        End Try 'Pad Thickness
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_rebar_size_top_dir1"), Integer)) Then
                Me.pad_rebar_size_top_dir1 = CType(ubDr.Item("pad_rebar_size_top_dir1"), Integer)
            Else
                Me.pad_rebar_size_top_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_size_top_dir1 = Nothing
        End Try 'Pad Rebar Size (Top Direction 1)
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_rebar_size_bottom_dir1"), Integer)) Then
                Me.pad_rebar_size_bottom_dir1 = CType(ubDr.Item("pad_rebar_size_bottom_dir1"), Integer)
            Else
                Me.pad_rebar_size_bottom_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_size_bottom_dir1 = Nothing
        End Try 'Pad Rebar Size (Bottom Direction 1)
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_rebar_size_top_dir2"), Integer)) Then
                Me.pad_rebar_size_top_dir2 = CType(ubDr.Item("pad_rebar_size_top_dir2"), Integer)
            Else
                Me.pad_rebar_size_top_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_size_top_dir2 = Nothing
        End Try 'Pad Rebar Size (Top Direction 2)
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_rebar_size_bottom_dir2"), Integer)) Then
                Me.pad_rebar_size_bottom_dir2 = CType(ubDr.Item("pad_rebar_size_bottom_dir2"), Integer)
            Else
                Me.pad_rebar_size_bottom_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_size_bottom_dir2 = Nothing
        End Try 'Pad Rebar Size (Bottom Direction 2)
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_rebar_quantity_top_dir1"), Integer)) Then
                Me.pad_rebar_quantity_top_dir1 = CType(ubDr.Item("pad_rebar_quantity_top_dir1"), Integer)
            Else
                Me.pad_rebar_quantity_top_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_top_dir1 = Nothing
        End Try 'Pad Rebar Quantity (Top Direction 1)
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_rebar_quantity_bottom_dir1"), Double)) Then
                Me.pad_rebar_quantity_bottom_dir1 = CType(ubDr.Item("pad_rebar_quantity_bottom_dir1"), Double)
            Else
                Me.pad_rebar_quantity_bottom_dir1 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_bottom_dir1 = Nothing
        End Try 'Pad Rebar Quantity (Bottom Direction 1)
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_rebar_quantity_top_dir2"), Double)) Then
                Me.pad_rebar_quantity_top_dir2 = CType(ubDr.Item("pad_rebar_quantity_top_dir2"), Double)
            Else
                Me.pad_rebar_quantity_top_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_top_dir2 = Nothing
        End Try 'Pad Rebar Quantity (Top Direction 2)
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_rebar_quantity_bottom_dir2"), Double)) Then
                Me.pad_rebar_quantity_bottom_dir2 = CType(ubDr.Item("pad_rebar_quantity_bottom_dir2"), Double)
            Else
                Me.pad_rebar_quantity_bottom_dir2 = Nothing
            End If
        Catch
            Me.pad_rebar_quantity_bottom_dir2 = Nothing
        End Try 'Pad Rebar Quantity (Bottom Direction 2)
        Try
            If Not IsDBNull(CType(ubDr.Item("pad_clear_cover"), Double)) Then
                Me.pad_clear_cover = CType(ubDr.Item("pad_clear_cover"), Double)
            Else
                Me.pad_clear_cover = Nothing
            End If
        Catch
            Me.pad_clear_cover = Nothing
        End Try 'Pad Clear Cover
        Try
            If Not IsDBNull(CType(ubDr.Item("total_soil_unit_weight"), Double)) Then
                Me.total_soil_unit_weight = CType(ubDr.Item("total_soil_unit_weight"), Double)
            Else
                Me.total_soil_unit_weight = Nothing
            End If
        Catch
            Me.total_soil_unit_weight = Nothing
        End Try 'Total Soil Unit Weight
        Try
            Me.bearing_type = CType(ubDr.Item("bearing_type"), String)
        Catch
            Me.bearing_type = "Ultimate Gross Bearing, Qult:"
        End Try 'Bearing Type
        Try
            If Not IsDBNull(CType(ubDr.Item("nominal_bearing_capacity"), Double)) Then
                Me.nominal_bearing_capacity = CType(ubDr.Item("nominal_bearing_capacity"), Double)
            Else
                Me.nominal_bearing_capacity = Nothing
            End If
        Catch
            Me.nominal_bearing_capacity = Nothing
        End Try 'Nominal Bearing Capacity
        Try
            If Not IsDBNull(CType(ubDr.Item("cohesion"), Double)) Then
                Me.cohesion = CType(ubDr.Item("cohesion"), Double)
            Else
                Me.cohesion = Nothing
            End If
        Catch
            Me.cohesion = Nothing
        End Try 'Cohesion
        Try
            If Not IsDBNull(CType(ubDr.Item("friction_angle"), Double)) Then
                Me.friction_angle = CType(ubDr.Item("friction_angle"), Double)
            Else
                Me.friction_angle = Nothing
            End If
        Catch
            Me.friction_angle = Nothing
        End Try 'Friction Angle
        Try
            If Not IsDBNull(CType(ubDr.Item("spt_blow_count"), Double)) Then
                Me.spt_blow_count = CType(ubDr.Item("spt_blow_count"), Double)
            Else
                Me.spt_blow_count = Nothing
            End If
        Catch
            Me.spt_blow_count = Nothing
        End Try 'STP Blow Count
        Try
            If Not IsDBNull(CType(ubDr.Item("base_friction_factor"), Double)) Then
                Me.base_friction_factor = CType(ubDr.Item("base_friction_factor"), Double)
            Else
                Me.base_friction_factor = Nothing
            End If
        Catch
            Me.base_friction_factor = Nothing
        End Try 'Base Friction Factor
        Try
            If Not IsDBNull(CType(ubDr.Item("neglect_depth"), Double)) Then
                Me.neglect_depth = CType(ubDr.Item("neglect_depth"), Double)
            Else
                Me.neglect_depth = Nothing
            End If
        Catch
            Me.neglect_depth = Nothing
        End Try 'Neglect Depth
        Try
            Me.bearing_distribution_type = CType(ubDr.Item("bearing_distribution_type"), Boolean)
        Catch
            Me.bearing_distribution_type = True
        End Try 'Bearing Distribution Type
        Try
            If Not IsDBNull(CType(ubDr.Item("groundwater_depth"), Double)) Then
                Me.groundwater_depth = CType(ubDr.Item("groundwater_depth"), Double)
            Else
                Me.groundwater_depth = Nothing
            End If
        Catch
            Me.groundwater_depth = -1
        End Try 'Groundwater Depth
        Try
            Me.basic_soil_check = CType(ubDr.Item("basic_soil_check"), Boolean)
        Catch
            Me.basic_soil_check = False
        End Try 'Basic Soil Interaction up to 110% Acceptable?
        Try
            Me.structural_check = CType(ubDr.Item("structural_check"), Boolean)
        Catch
            Me.structural_check = False
        End Try 'Structural Checks up to 105.0% Acceptable?
        Try
            Me.tool_version = CType(ubDr.Item("tool_version"), String)
        Catch
            Me.tool_version = ""
        End Try 'Tool Version

        'If Me.modified = True Then
        '    For Each ModifiedRangeDataRow As DataRow In ds.Tables("Unit Base Modified Ranges SQL").Rows
        '        Dim modRefID As Integer = CType(ModifiedRangeDataRow.Item("modified_id"), Integer)
        '        If modRefID = refID Then
        '            Me.ModifiedRanges.Add(New ModifiedRange(ModifiedRangeDataRow))
        '        End If
        '    Next 'Add Modified Ranges to Modified Range Object
        'End If

    End Sub 'Generate a ub from EDS

    'Public Sub New(ExcelFilePath As String, Optional BU As String = Nothing, Optional structureID As String = Nothing)
    Public Sub New(ExcelFilePath As String, Optional ByRef Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet

        'This can be cleaned up. Do we need a list of ExcelDTParameters if there is only one? - DHS
        For Each item As EXCELDTParameter In excelDTParams
            'Get additional tables from excel file 
            excelDS.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
        Next

        If excelDS.Tables.Contains("Unit Base General Details EXCEL") Then
            Dim dr = excelDS.Tables("Unit Base General Details EXCEL").Rows(0)
            'Need to dimension DataRow from GenStructure/TNX and anywhere else inputs may come from as well - MRR

            Try
                Me.ID = CType(dr.Item("unit_base_id"), Integer)
            Catch
                Me.ID = 0
            End Try 'Unit Base ID
            Try
                Me.extension_above_grade = CType(dr.Item("extension_above_grade"), Double)
            Catch
                Me.extension_above_grade = Nothing
            End Try 'Extension Above Grade
            Try
                Me.foundation_depth = CType(dr.Item("foundation_depth"), Double)
            Catch
                Me.foundation_depth = Nothing
            End Try 'Foundation Depth
            Try
                Me.concrete_compressive_strength = CType(dr.Item("concrete_compressive_strength"), Double)
            Catch
                Me.concrete_compressive_strength = Nothing
            End Try 'Concrete Compressive Strength
            Try
                Me.dry_concrete_density = CType(dr.Item("dry_concrete_density"), Double)
            Catch
                Me.dry_concrete_density = Nothing
            End Try 'Dry Concrete Density
            Try
                Me.rebar_grade = CType(dr.Item("rebar_grade"), Double)
            Catch
                Me.rebar_grade = Nothing
            End Try 'Rebar Grade
            Try
                Me.top_and_bottom_rebar_different = CType(dr.Item("top_and_bottom_rebar_different"), Boolean)
            Catch
                Me.top_and_bottom_rebar_different = False
            End Try 'Top and Bottom Rebar Different
            Try
                Me.block_foundation = CType(dr.Item("block_foundation"), Boolean)
            Catch
                Me.block_foundation = False
            End Try 'Block Foundation 
            Try
                Me.rectangular_foundation = CType(dr.Item("rectangular_foundation"), Boolean)
            Catch
                Me.rectangular_foundation = False
            End Try 'Rectangular Foundation
            Try
                Me.base_plate_distance_above_foundation = CType(dr.Item("base_plate_distance_above_foundation"), Double)
            Catch
                Me.base_plate_distance_above_foundation = Nothing
            End Try 'Base Plate Distance Above Foundation
            Try
                Me.bolt_circle_bearing_plate_width = CType(dr.Item("bolt_circle_bearing_plate_width"), Double)
            Catch
                Me.bolt_circle_bearing_plate_width = Nothing
            End Try 'Bolt Circle Bearing Plate Width
            Try
                Me.tower_centroid_offset = CType(dr.Item("tower_centroid_offset"), Boolean)
            Catch
                Me.tower_centroid_offset = Nothing
            End Try 'Tower Centroid Offset
            Try
                Me.pier_shape = CType(dr.Item("pier_shape"), String)
            Catch
                'Me.pier_shape = Nothing
                Me.pier_shape = ""
            End Try 'Pier Shape
            Try
                Me.pier_diameter = CType(dr.Item("pier_diameter"), Double)
            Catch
                Me.pier_diameter = Nothing
            End Try 'Pier Diameter
            Try
                Me.pier_rebar_quantity = CType(dr.Item("pier_rebar_quantity"), Double)
            Catch
                Me.pier_rebar_quantity = Nothing
            End Try 'Pier Rebar Quantity
            Try
                Me.pier_rebar_size = CType(dr.Item("pier_rebar_size"), Integer)
            Catch
                Me.pier_rebar_size = Nothing
            End Try 'Pier Rebar Size
            Try
                Me.pier_tie_quantity = CType(dr.Item("pier_tie_quantity"), Double)
            Catch
                Me.pier_tie_quantity = Nothing
            End Try 'Pier Tie Quantity
            Try
                Me.pier_tie_size = CType(dr.Item("pier_tie_size"), Integer)
            Catch
                Me.pier_tie_size = Nothing
            End Try 'Pier Tie Size
            Try
                Me.pier_reinforcement_type = CType(dr.Item("pier_reinforcement_type"), String)
            Catch
                Me.pier_reinforcement_type = ""
            End Try 'Pier Reinforcement Type
            Try
                Me.pier_clear_cover = CType(dr.Item("pier_clear_cover"), Double)
            Catch
                Me.pier_clear_cover = Nothing
            End Try 'Pier Clear Cover
            Try
                Me.pad_width_1 = CType(dr.Item("pad_width_1"), Double)
            Catch
                Me.pad_width_1 = Nothing
            End Try 'Pad Width 1
            Try
                Me.pad_width_2 = CType(dr.Item("pad_width_2"), Double)
            Catch
                Me.pad_width_2 = Nothing
            End Try 'Pad Width 2
            Try
                Me.pad_thickness = CType(dr.Item("pad_thickness"), Double)
            Catch
                Me.pad_thickness = Nothing
            End Try 'Pad Thickness
            Try
                Me.pad_rebar_size_top_dir1 = CType(dr.Item("pad_rebar_size_top_dir1"), Integer)
            Catch
                Me.pad_rebar_size_top_dir1 = Nothing
            End Try 'Pad Rebar Size (Top Direction 1)
            Try
                Me.pad_rebar_size_bottom_dir1 = CType(dr.Item("pad_rebar_size_bottom_dir1"), Integer)
            Catch
                Me.pad_rebar_size_bottom_dir1 = Nothing
            End Try 'Pad Rebar Size (Bottom Direction 1)
            Try
                Me.pad_rebar_size_top_dir2 = CType(dr.Item("pad_rebar_size_top_dir2"), Integer)
            Catch
                Me.pad_rebar_size_top_dir2 = Nothing
            End Try 'Pad Rebar Size (Top Direction 2)
            Try
                Me.pad_rebar_size_bottom_dir2 = CType(dr.Item("pad_rebar_size_bottom_dir2"), Integer)
            Catch
                Me.pad_rebar_size_bottom_dir2 = Nothing
            End Try 'Pad Rebar Size (Bottom Direction 2)
            Try
                Me.pad_rebar_quantity_top_dir1 = CType(dr.Item("pad_rebar_quantity_top_dir1"), Double)
            Catch
                Me.pad_rebar_quantity_top_dir1 = Nothing
            End Try 'Pad Rebar Quantity (Top Direction 1)
            Try
                Me.pad_rebar_quantity_bottom_dir1 = CType(dr.Item("pad_rebar_quantity_bottom_dir1"), Double)
            Catch
                Me.pad_rebar_quantity_bottom_dir1 = Nothing
            End Try 'Pad Rebar Quantity (Bottom Direction 1)
            Try
                Me.pad_rebar_quantity_top_dir2 = CType(dr.Item("pad_rebar_quantity_top_dir2"), Double)
            Catch
                Me.pad_rebar_quantity_top_dir2 = Nothing
            End Try 'Pad Rebar Quantity (Top Direction 2)
            Try
                Me.pad_rebar_quantity_bottom_dir2 = CType(dr.Item("pad_rebar_quantity_bottom_dir2"), Double)
            Catch
                Me.pad_rebar_quantity_bottom_dir2 = Nothing
            End Try 'Pad Rebar Quantity (Bottom Direction 2)
            Try
                Me.pad_clear_cover = CType(dr.Item("pad_clear_cover"), Double)
            Catch
                Me.pad_clear_cover = Nothing
            End Try 'Pad Clear Cover
            Try
                Me.total_soil_unit_weight = CType(dr.Item("total_soil_unit_weight"), Double)
            Catch
                Me.total_soil_unit_weight = Nothing
            End Try 'Total Soil Unit Weight
            Try
                Me.bearing_type = CType(dr.Item("bearing_type"), String)
            Catch
                Me.bearing_type = "Ultimate Gross Bearing, Qult:"
            End Try 'Bearing Type ******String Options "Ultimate Gross Bearing, Qult:" / "Ultimate Net Bearing, Qnet:" *******
            Try
                Me.nominal_bearing_capacity = CType(dr.Item("nominal_bearing_capacity"), Double)
            Catch
                Me.nominal_bearing_capacity = Nothing
            End Try 'Nominal Bearing Capacity
            Try
                Me.cohesion = CType(dr.Item("cohesion"), Double)
            Catch
                Me.cohesion = Nothing
            End Try 'Cohesion
            Try
                Me.friction_angle = CType(dr.Item("friction_angle"), Double)
            Catch
                Me.friction_angle = Nothing
            End Try 'Friction Angle
            Try
                Me.spt_blow_count = CType(dr.Item("spt_blow_count"), Double)
            Catch
                Me.spt_blow_count = Nothing
            End Try 'STP Blow Count
            Try
                Me.base_friction_factor = CType(dr.Item("base_friction_factor"), Double)
            Catch
                Me.base_friction_factor = Nothing
            End Try 'STP Blow Count
            Try
                Me.neglect_depth = CType(dr.Item("neglect_depth"), Double)
            Catch
                Me.neglect_depth = Nothing
            End Try 'Neglect Depth
            Try
                If CType(dr.Item("bearing_distribution_type"), String) = "No" Then
                    Me.bearing_distribution_type = False
                Else
                    Me.bearing_distribution_type = True
                End If
            Catch
                Me.bearing_distribution_type = True
            End Try 'Bearing Distribution Type
            Try
                If Not IsNothing(CType(dr.Item("groundwater_depth"), Double)) Then
                    Me.groundwater_depth = CType(dr.Item("groundwater_depth"), Double)
                Else
                    Me.groundwater_depth = Nothing
                End If
            Catch
                Me.groundwater_depth = -1
            End Try 'Groundwater Depth
            Try
                Me.basic_soil_check = CType(dr.Item("basic_soil_check"), Boolean)
            Catch
                Me.basic_soil_check = False
            End Try 'Basic Soil Interaction up to 110% Acceptable?
            Try
                Me.structural_check = CType(dr.Item("structural_check"), Boolean)
            Catch
                Me.structural_check = False
            End Try 'Structural Checks up to 105.0% Acceptable?
            Try
                Me.tool_version = CType(dr.Item("tool_version"), String)
            Catch
                Me.tool_version = ""
            End Try 'Tool Version
            'Try
            '    Me.modified = CType(dr.Item("modified"), Boolean)
            'Catch
            '    Me.modified = False
            'End Try 'Modified

            'For Each ModifiedRangeDataRow As DataRow In ds.Tables("Unit Base Modified Ranges EXCEL").Rows
            '    Me.ModifiedRanges.Add(New ModifiedRange(ModifiedRangeDataRow))
            'Next 'Add Modified Ranges to Modified Range Object

        End If

    End Sub

#End Region

#Region "Save to Excel"
    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        ''''''Customize for each foundation type'''''

        With wb
            .Worksheets("Input").Range("ID").Value = CType(Me.ID, Integer)

            If Not IsNothing(Me.pier_shape) Then
                .Worksheets("Input").Range("shape").Value = CType(Me.pier_shape, String)
            End If

            If Not IsNothing(Me.pier_diameter) Then
                .Worksheets("Input").Range("dpier").Value = CType(Me.pier_diameter, Double)
            Else
                .Worksheets("Input").Range("dpier").ClearContents
            End If

            If Not IsNothing(Me.extension_above_grade) Then
                .Worksheets("Input").Range("E").Value = CType(Me.extension_above_grade, Double)
            Else
                .Worksheets("Input").Range("E").ClearContents
            End If

            If Not IsNothing(Me.pier_rebar_size) Then
                .Worksheets("Input").Range("Sc").Value = CType(Me.pier_rebar_size, Integer)
            Else
                .Worksheets("Input").Range("Sc").ClearContents
            End If

            If Not IsNothing(Me.pier_rebar_quantity) Then
                .Worksheets("Input").Range("mc").Value = CType(Me.pier_rebar_quantity, Double)
            Else
                .Worksheets("Input").Range("mc").ClearContents
            End If

            If Not IsNothing(Me.pier_tie_size) Then
                .Worksheets("Input").Range("St").Value = CType(Me.pier_tie_size, Integer)
            Else
                .Worksheets("Input").Range("St").ClearContents
            End If

            If Not IsNothing(Me.pier_tie_quantity) Then
                .Worksheets("Input").Range("mt").Value = CType(Me.pier_tie_quantity, Double)
            Else
                .Worksheets("Input").Range("mt").ClearContents
            End If

            If Not IsNothing(Me.pier_reinforcement_type) Then
                .Worksheets("Input").Range("PierReinfType").Value = CType(Me.pier_reinforcement_type, String)
            End If

            If Not IsNothing(Me.pier_clear_cover) Then
                .Worksheets("Input").Range("ccpier").Value = CType(Me.pier_clear_cover, Double)
            Else
                .Worksheets("Input").Range("ccpier").ClearContents
            End If

            If Not IsNothing(Me.foundation_depth) Then
                .Worksheets("Input").Range("D").Value = CType(Me.foundation_depth, Double)
            Else
                .Worksheets("Input").Range("D").ClearContents
            End If

            If Not IsNothing(Me.pad_width_1) Then
                .Worksheets("Input").Range("W").Value = CType(Me.pad_width_1, Double)
            Else
                .Worksheets("Input").Range("W").ClearContents
            End If

            If Not IsNothing(Me.pad_width_2) Then
                .Worksheets("Input").Range("W.dir2").Value = CType(Me.pad_width_2, Double)
            Else
                .Worksheets("Input").Range("W.dir2").ClearContents
            End If

            If Not IsNothing(Me.pad_thickness) Then
                .Worksheets("Input").Range("T").Value = CType(Me.pad_thickness, Double)
            Else
                .Worksheets("Input").Range("T").ClearContents
            End If

            If Not IsNothing(Me.pad_rebar_size_top_dir1) Then
                .Worksheets("Input").Range("sptop").Value = CType(Me.pad_rebar_size_top_dir1, Integer)
            Else
                .Worksheets("Input").Range("sptop").ClearContents
            End If

            If Not IsNothing(Me.pad_rebar_size_bottom_dir1) Then
                .Worksheets("Input").Range("Sp").Value = CType(Me.pad_rebar_size_bottom_dir1, Integer)
            Else
                .Worksheets("Input").Range("Sp").ClearContents
            End If

            If Not IsNothing(Me.pad_rebar_size_top_dir2) Then
                .Worksheets("Input").Range("sptop2").Value = CType(Me.pad_rebar_size_top_dir2, Integer)
            Else
                .Worksheets("Input").Range("sptop2").ClearContents
            End If

            If Not IsNothing(Me.pad_rebar_size_bottom_dir2) Then
                .Worksheets("Input").Range("sp_2").Value = CType(Me.pad_rebar_size_bottom_dir2, Integer)
            Else
                .Worksheets("Input").Range("sp_2").ClearContents
            End If

            If Not IsNothing(Me.pad_rebar_quantity_top_dir1) Then
                .Worksheets("Input").Range("mptop").Value = CType(Me.pad_rebar_quantity_top_dir1, Double)
            Else
                .Worksheets("Input").Range("mptop").ClearContents
            End If

            If Not IsNothing(Me.pad_rebar_quantity_bottom_dir1) Then
                .Worksheets("Input").Range("mp").Value = CType(Me.pad_rebar_quantity_bottom_dir1, Double)
            Else
                .Worksheets("Input").Range("mp").ClearContents
            End If

            If Not IsNothing(Me.pad_rebar_quantity_top_dir2) Then
                .Worksheets("Input").Range("mptop2").Value = CType(Me.pad_rebar_quantity_top_dir2, Double)
            Else
                .Worksheets("Input").Range("mptop2").ClearContents
            End If

            If Not IsNothing(Me.pad_rebar_quantity_bottom_dir2) Then
                .Worksheets("Input").Range("mp_2").Value = CType(Me.pad_rebar_quantity_bottom_dir2, Double)
            Else
                .Worksheets("Input").Range("mp_2").ClearContents
            End If

            If Not IsNothing(Me.pad_clear_cover) Then
                .Worksheets("Input").Range("ccpad").Value = CType(Me.pad_clear_cover, Double)
            Else
                .Worksheets("Input").Range("ccpad").ClearContents
            End If

            If Not IsNothing(Me.rebar_grade) Then
                .Worksheets("Input").Range("Fy").Value = CType(Me.rebar_grade, Double)
            Else
                .Worksheets("Input").Range("Fy").ClearContents
            End If

            If Not IsNothing(Me.concrete_compressive_strength) Then
                .Worksheets("Input").Range("F\c").Value = CType(Me.concrete_compressive_strength, Double)
            Else
                .Worksheets("Input").Range("F\c").ClearContents
            End If

            If Not IsNothing(Me.dry_concrete_density) Then
                .Worksheets("Input").Range("ConcreteDensity").Value = CType(Me.dry_concrete_density, Double)
            Else
                .Worksheets("Input").Range("ConcreteDensity").ClearContents
            End If

            If Not IsNothing(Me.total_soil_unit_weight) Then
                .Worksheets("Input").Range("γ").Value = CType(Me.total_soil_unit_weight, Double)
            Else
                .Worksheets("Input").Range("γ").ClearContents
            End If

            If Not IsNothing(Me.bearing_type) Then
                .Worksheets("Input").Range("BearingType").Value = CType(Me.bearing_type, String)
            Else
                .Worksheets("Input").Range("BearingType").ClearContents
            End If

            If Not IsNothing(Me.nominal_bearing_capacity) Then
                .Worksheets("Input").Range("Qinput").Value = CType(Me.nominal_bearing_capacity, Double)
            Else
                .Worksheets("Input").Range("Qinput").ClearContents
            End If

            If Not IsNothing(Me.cohesion) Then
                .Worksheets("Input").Range("Cu").Value = CType(Me.cohesion, Double)
            Else
                .Worksheets("Input").Range("Cu").ClearContents
            End If

            If Not IsNothing(Me.friction_angle) Then
                .Worksheets("Input").Range("ϕ").Value = CType(Me.friction_angle, Double)
            Else
                .Worksheets("Input").Range("ϕ").ClearContents
            End If

            If Not IsNothing(Me.spt_blow_count) Then
                .Worksheets("Input").Range("N_blows").Value = CType(Me.spt_blow_count, Double)
            Else
                .Worksheets("Input").Range("N_blows").ClearContents
            End If

            If Not IsNothing(Me.base_friction_factor) Then
                .Worksheets("Input").Range("μ").Value = CType(Me.base_friction_factor, Double)
            Else
                .Worksheets("Input").Range("μ").ClearContents
            End If

            If Not IsNothing(Me.neglect_depth) Then
                .Worksheets("Input").Range("N").Value = CType(Me.neglect_depth, Double)
            End If

            If Me.bearing_distribution_type = False Then
                .Worksheets("Input").Range("Rock").Value = "No"
            Else
                .Worksheets("Input").Range("Rock").Value = "Yes"
            End If

            If Me.groundwater_depth = -1 Then
                .Worksheets("Input").Range("gw").Value = "N/A"
            Else
                .Worksheets("Input").Range("gw").Value = CType(Me.groundwater_depth, Double)
            End If

            If Not IsNothing(Me.top_and_bottom_rebar_different) Then
                .Worksheets("Input").Range("DifferentReinforcementBoolean").Value = CType(Me.top_and_bottom_rebar_different, Boolean)
            End If

            If Not IsNothing(Me.block_foundation) Then
                .Worksheets("Input").Range("BlockFoundationBoolean").Value = CType(Me.block_foundation, Boolean)
            End If

            If Not IsNothing(Me.rectangular_foundation) Then
                .Worksheets("Input").Range("RectangularPadBoolean").Value = CType(Me.rectangular_foundation, Boolean)
            End If

            If Not IsNothing(Me.base_plate_distance_above_foundation) Then
                .Worksheets("Input").Range("bpdist").Value = CType(Me.base_plate_distance_above_foundation, Double)
            Else
                .Worksheets("Input").Range("bpdist").ClearContents
            End If

            If Not IsNothing(Me.bolt_circle_bearing_plate_width) Then
                .Worksheets("Input").Range("BC").Value = CType(Me.bolt_circle_bearing_plate_width, Double)
            Else
                .Worksheets("Input").Range("BC").ClearContents
            End If

            If Not IsNothing(Me.tower_centroid_offset) Then
                .Worksheets("Input").Range("TowerCentroidOffsetBoolean").Value = CType(Me.tower_centroid_offset, Boolean)
            Else
                .Worksheets("Input").Range("TowerCentroidOffsetBoolean").ClearContents
            End If

            If Not IsNothing(Me.basic_soil_check) Then
                .Worksheets("Input").Range("SoilInteractionBoolean").Value = CType(Me.basic_soil_check, Boolean)
            End If

            If Not IsNothing(Me.structural_check) Then
                .Worksheets("Input").Range("StructuralCheckBoolean").Value = CType(Me.structural_check, Boolean)
            End If

            'If Not IsNothing(Me.tia_current) Then
            '    .Worksheets("Input").Range("TIA").Value = CType(Me.tia_current, String)
            'End If
        End With

    End Sub

#End Region

#Region "Save to EDS"

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.ID.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_shape.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_diameter.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.extension_above_grade.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_rebar_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_tie_size.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_tie_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_reinforcement_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_clear_cover.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.foundation_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_width_1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_width_2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_thickness.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_size_top_dir1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_size_bottom_dir1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_size_top_dir2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_size_bottom_dir2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_top_dir1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_bottom_dir1.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_top_dir2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_bottom_dir2.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_clear_cover.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_grade.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.concrete_compressive_strength.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.dry_concrete_density.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.total_soil_unit_weight.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bearing_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.nominal_bearing_capacity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cohesion.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.friction_angle.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.spt_blow_count.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.base_friction_factor.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.neglect_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bearing_distribution_type.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.groundwater_depth.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.top_and_bottom_rebar_different.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.block_foundation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rectangular_foundation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.base_plate_distance_above_foundation.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_circle_bearing_plate_width.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tower_centroid_offset.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_rebar_quantity.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.basic_soil_check.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structural_check.ToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.tool_version.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.valid_from.ToString.FormatDBValue)
        'SQLInsertValues = SQLInsertValues.AddtoDBString(Me.valid_to.ToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

        'SQLInsertFields = SQLInsertFields.AddtoDBString("ID")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bus_unit")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structure_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_shape")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_diameter")
        SQLInsertFields = SQLInsertFields.AddtoDBString("extension_above_grade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_rebar_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_tie_size")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_tie_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_reinforcement_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_clear_cover")
        SQLInsertFields = SQLInsertFields.AddtoDBString("foundation_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_width_1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_width_2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_thickness")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_size_top_dir1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_size_bottom_dir1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_size_top_dir2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_size_bottom_dir2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_quantity_top_dir1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_quantity_bottom_dir1")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_quantity_top_dir2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_rebar_quantity_bottom_dir2")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pad_clear_cover")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rebar_grade")
        SQLInsertFields = SQLInsertFields.AddtoDBString("concrete_compressive_strength")
        SQLInsertFields = SQLInsertFields.AddtoDBString("dry_concrete_density")
        SQLInsertFields = SQLInsertFields.AddtoDBString("total_soil_unit_weight")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bearing_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("nominal_bearing_capacity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("cohesion")
        SQLInsertFields = SQLInsertFields.AddtoDBString("friction_angle")
        SQLInsertFields = SQLInsertFields.AddtoDBString("spt_blow_count")
        SQLInsertFields = SQLInsertFields.AddtoDBString("base_friction_factor")
        SQLInsertFields = SQLInsertFields.AddtoDBString("neglect_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bearing_distribution_type")
        SQLInsertFields = SQLInsertFields.AddtoDBString("groundwater_depth")
        SQLInsertFields = SQLInsertFields.AddtoDBString("top_and_bottom_rebar_different")
        SQLInsertFields = SQLInsertFields.AddtoDBString("block_foundation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("rectangular_foundation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("base_plate_distance_above_foundation")
        SQLInsertFields = SQLInsertFields.AddtoDBString("bolt_circle_bearing_plate_width")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tower_centroid_offset")
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_rebar_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("basic_soil_check")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structural_check")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tool_version")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("valid_from")
        'SQLInsertFields = SQLInsertFields.AddtoDBString("valid_to")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdate() As String
        SQLUpdate = ""

        'SQLUpdate = SQLUpdate.AddtoDBString("ID = " & Me.ID.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("bus_unit = " & Me.bus_unit.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("structure_id = " & Me.structure_id.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_shape = " & Me.pier_shape.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_diameter = " & Me.pier_diameter.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("extension_above_grade = " & Me.extension_above_grade.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_rebar_size = " & Me.pier_rebar_size.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_tie_size = " & Me.pier_tie_size.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_tie_quantity = " & Me.pier_tie_quantity.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_reinforcement_type = " & Me.pier_reinforcement_type.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_clear_cover = " & Me.pier_clear_cover.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("foundation_depth = " & Me.foundation_depth.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_width_1 = " & Me.pad_width_1.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_width_2 = " & Me.pad_width_2.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_thickness = " & Me.pad_thickness.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_size_top_dir1 = " & Me.pad_rebar_size_top_dir1.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_size_bottom_dir1 = " & Me.pad_rebar_size_bottom_dir1.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_size_top_dir2 = " & Me.pad_rebar_size_top_dir2.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_size_bottom_dir2 = " & Me.pad_rebar_size_bottom_dir2.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_quantity_top_dir1 = " & Me.pad_rebar_quantity_top_dir1.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_quantity_bottom_dir1 = " & Me.pad_rebar_quantity_bottom_dir1.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_quantity_top_dir2 = " & Me.pad_rebar_quantity_top_dir2.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_rebar_quantity_bottom_dir2 = " & Me.pad_rebar_quantity_bottom_dir2.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pad_clear_cover = " & Me.pad_clear_cover.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("rebar_grade = " & Me.rebar_grade.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("concrete_compressive_strength = " & Me.concrete_compressive_strength.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("dry_concrete_density = " & Me.dry_concrete_density.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("total_soil_unit_weight = " & Me.total_soil_unit_weight.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("bearing_type = " & Me.bearing_type.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("nominal_bearing_capacity = " & Me.nominal_bearing_capacity.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("cohesion = " & Me.cohesion.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("friction_angle = " & Me.friction_angle.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("spt_blow_count = " & Me.spt_blow_count.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("base_friction_factor = " & Me.base_friction_factor.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("neglect_depth = " & Me.neglect_depth.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("bearing_distribution_type = " & Me.bearing_distribution_type.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("groundwater_depth = " & Me.groundwater_depth.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("top_and_bottom_rebar_different = " & Me.top_and_bottom_rebar_different.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("block_foundation = " & Me.block_foundation.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("rectangular_foundation = " & Me.rectangular_foundation.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("base_plate_distance_above_foundation = " & Me.base_plate_distance_above_foundation.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("bolt_circle_bearing_plate_width = " & Me.bolt_circle_bearing_plate_width.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("tower_centroid_offset = " & Me.tower_centroid_offset.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("pier_rebar_quantity = " & Me.pier_rebar_quantity.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("basic_soil_check = " & Me.basic_soil_check.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("structural_check = " & Me.structural_check.ToString.FormatDBValue)
        SQLUpdate = SQLUpdate.AddtoDBString("tool_version = " & Me.tool_version.ToString.FormatDBValue)
        'SQLUpdate = SQLUpdate.AddtoDBString("modified_person_id = " & Me.modified_person_id.ToString.FormatDBValue)
        'SQLUpdate = SQLUpdate.AddtoDBString("process_stage = " & Me.process_stage.ToString.FormatDBValue)
        'SQLUpdate = SQLUpdate.AddtoDBString("valid_from = " & Me.valid_from.ToString.FormatDBValue)
        'SQLUpdate = SQLUpdate.AddtoDBString("valid_to = " & Me.valid_to.ToString.FormatDBValue)


        Return SQLUpdate
    End Function

#End Region

#Region "Check Changes"

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Equals = True
        If changes Is Nothing Then changes = New List(Of AnalysisChange)

        'Makes sure you are comparing to the same object type
        'Customize this to the object type
        Dim otherToCompare As SST_Unit_Base = TryCast(other, SST_Unit_Base)
        If otherToCompare Is Nothing Then Return False

        Dim CategoryName As String = "Foundation - Unit Base"

        Equals = If(Me.pier_diameter.CheckChange(otherToCompare.pier_diameter, changes, CategoryName, "Pier Diameter"), Equals, False)

    End Function
    'Public Overrides Function CompareMe(Of T As EDSObject)(ByVal previous As T) As Boolean
    '    ''''''Customize for each foundation type'''''

    '    'We're overriding a generic function but we only want this override to be applicable to this specific child class
    '    'Try to cast the input object (previous) to this type of object
    '    Dim prevUnitBase As UnitBase = TryCast(previous, UnitBase)
    '    'If cast failed, new object will be nothing
    '    If prevUnitBase Is Nothing Then
    '        'Cast Failed
    '        Return False
    '    End If

    '    Dim comparer As New ObjectsComparer.Comparer(Of UnitBase)

    '    'Ignore these 
    '    comparer.IgnoreMember("ID")
    '    comparer.IgnoreMember("activeDatabase")
    '    comparer.IgnoreMember("databaseIdentity")
    '    comparer.IgnoreMember("differences")
    '    comparer.IgnoreMember("Insert")
    '    comparer.IgnoreMember("Update")
    '    comparer.IgnoreMember("Delete")
    '    comparer.IgnoreMember("workBookPath")
    '    comparer.IgnoreMember("templatePath")
    '    comparer.IgnoreMember("fileType")
    '    comparer.IgnoreMember("EDSTableName")
    '    comparer.IgnoreMember("excelDTParams")

    '    Dim basePath As String = "Foundations - " & Me.foundationType
    '    'Me.differences = New List(Of ObjectsComparer.Difference)
    '    Dim differences As IEnumerable(Of ObjectsComparer.Difference) = Nothing

    '    CompareMe = comparer.Compare(Me, prevUnitBase, differences)


    '    'If the current item doesn't have an ID it wasn't created from EDS and should need to be inserted
    '    'If we weren't able to find a previous item in EDS we need to insert
    '    'If Me.ID Is Nothing Then
    '    '    differences.Add(New ObjectsComparer.Difference(basePath, "New", "", ObjectsComparer.DifferenceTypes.MissedMemberInSecondObject))
    '    'End If

    '    'CompareMe = True

    '    ''Check Details
    '    'If Not comparer.Compare(Me.pier_shape, prevUnitBase.pier_shape, basePath & " - pier shape", differences) Then CompareMe = False
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pier_diameter", Me.pier_diameter.ToString, previous.pier_diameter.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "extension_above_grade", Me.extension_above_grade.ToString, previous.extension_above_grade.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pier_rebar_size", Me.pier_rebar_size.ToString, previous.pier_rebar_size.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pier_tie_size", Me.pier_tie_size.ToString, previous.pier_tie_size.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pier_tie_quantity", Me.pier_tie_quantity.ToString, previous.pier_tie_quantity.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pier_reinforcement_type", Me.pier_reinforcement_type, previous.pier_reinforcement_type) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pier_clear_cover", Me.pier_clear_cover.ToString, previous.pier_clear_cover.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "foundation_depth", Me.foundation_depth.ToString, previous.foundation_depth.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_width_1", Me.pad_width_1.ToString, previous.pad_width_1.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_width_2", Me.pad_width_2.ToString, previous.pad_width_2.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_thickness", Me.pad_thickness.ToString, previous.pad_thickness.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_rebar_size_top_dir1", Me.pad_rebar_size_top_dir1.ToString, previous.pad_rebar_size_top_dir1.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_rebar_size_bottom_dir1", Me.pad_rebar_size_bottom_dir1.ToString, previous.pad_rebar_size_bottom_dir1.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_rebar_size_top_dir2", Me.pad_rebar_size_top_dir2.ToString, previous.pad_rebar_size_top_dir2.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_rebar_size_bottom_dir2", Me.pad_rebar_size_bottom_dir2.ToString, previous.pad_rebar_size_bottom_dir2.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_rebar_quantity_top_dir1", Me.pad_rebar_quantity_top_dir1.ToString, previous.pad_rebar_quantity_top_dir1.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_rebar_quantity_bottom_dir1", Me.pad_rebar_quantity_bottom_dir1.ToString, previous.pad_rebar_quantity_bottom_dir1.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_rebar_quantity_top_dir2", Me.pad_rebar_quantity_top_dir2.ToString, previous.pad_rebar_quantity_top_dir2.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_rebar_quantity_bottom_dir2", Me.pad_rebar_quantity_bottom_dir2.ToString, previous.pad_rebar_quantity_bottom_dir2.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pad_clear_cover", Me.pad_clear_cover.ToString, previous.pad_clear_cover.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "rebar_grade", Me.rebar_grade.ToString, previous.rebar_grade.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "concrete_compressive_strength", Me.concrete_compressive_strength.ToString, previous.concrete_compressive_strength.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "dry_concrete_density", Me.dry_concrete_density.ToString, previous.dry_concrete_density.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "total_soil_unit_weight", Me.total_soil_unit_weight.ToString, previous.total_soil_unit_weight.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "bearing_type", Me.bearing_type, previous.bearing_type) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "nominal_bearing_capacity", Me.nominal_bearing_capacity.ToString, previous.nominal_bearing_capacity.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "cohesion", Me.cohesion.ToString, previous.cohesion.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "friction_angle", Me.friction_angle.ToString, previous.friction_angle.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "spt_blow_count", Me.spt_blow_count.ToString, previous.spt_blow_count.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "base_friction_factor", Me.base_friction_factor.ToString, previous.base_friction_factor.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "neglect_depth", Me.neglect_depth.ToString, previous.neglect_depth.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "bearing_distribution_type", Me.bearing_distribution_type.ToString, previous.bearing_distribution_type.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "groundwater_depth", Me.groundwater_depth.ToString, previous.groundwater_depth.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "top_and_bottom_rebar_different", Me.top_and_bottom_rebar_different.ToString, previous.top_and_bottom_rebar_different.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "block_foundation", Me.block_foundation.ToString, previous.block_foundation.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "rectangular_foundation", Me.rectangular_foundation.ToString, previous.rectangular_foundation.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "base_plate_distance_above_foundation", Me.base_plate_distance_above_foundation.ToString, previous.base_plate_distance_above_foundation.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "bolt_circle_bearing_plate_width", Me.bolt_circle_bearing_plate_width.ToString, previous.bolt_circle_bearing_plate_width.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "tower_centroid_offset", Me.tower_centroid_offset.ToString, previous.tower_centroid_offset.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "pier_rebar_quantity", Me.pier_rebar_quantity.ToString, previous.pier_rebar_quantity.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "basic_soil_check", Me.basic_soil_check.ToString, previous.basic_soil_check.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "structural_check", Me.structural_check.ToString, previous.structural_check.ToString) Then changesMade = True
    '    'If Me.dbComparison.Check1Change(Me.foundationType, "tool_version", Me.tool_version, previous.tool_version) Then changesMade = True

    '    Me.differences = differences.ToList

    '    Return CompareMe

    'End Function
#End Region

End Class