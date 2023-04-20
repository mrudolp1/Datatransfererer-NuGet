Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet

Partial Public Class PierAndPad
    Inherits EDSExcelObject

#Region "Inheritted"
    '''Must override these inherited properties
    Public Overrides ReadOnly Property EDSObjectName As String = "Pier And Pad Foundation"
    Public Overrides ReadOnly Property EDSTableName As String = "fnd.pier_pad"
    Public Overrides ReadOnly Property TemplatePath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "Pier and Pad Foundation.xlsm")
    Public Overrides ReadOnly Property Template As Byte() = CCI_Engineering_Templates.My.Resources.Pier_and_Pad_Foundation
    Public Overrides ReadOnly Property ExcelDTParams As List(Of EXCELDTParameter)
        Get
            Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("Pier and Pad General Details EXCEL", "A2:AR3", "Details (SAPI)"),
                                                        New EXCELDTParameter("Pier and Pad General Results EXCEL", "A2:C16", "Results (SAPI)")}
        End Get
    End Property

#End Region

#Region "Define"
    Private _pier_shape As String
    Private _pier_diameter As Double?
    Private _extension_above_grade As Double?
    Private _pier_rebar_size As Integer?
    Private _pier_tie_size As Integer?
    Private _pier_tie_quantity As Integer?
    Private _pier_reinforcement_type As String
    Private _pier_clear_cover As Double?
    Private _foundation_depth As Double?
    Private _pad_width_1 As Double?
    Private _pad_width_2 As Double?
    Private _pad_thickness As Double?
    Private _pad_rebar_size_top_dir1 As Integer?
    Private _pad_rebar_size_bottom_dir1 As Integer?
    Private _pad_rebar_size_top_dir2 As Integer?
    Private _pad_rebar_size_bottom_dir2 As Integer?
    Private _pad_rebar_quantity_top_dir1 As Integer?
    Private _pad_rebar_quantity_bottom_dir1 As Integer?
    Private _pad_rebar_quantity_top_dir2 As Integer?
    Private _pad_rebar_quantity_bottom_dir2 As Integer?
    Private _pad_clear_cover As Double?
    Private _rebar_grade As Double?
    Private _concrete_compressive_strength As Double?
    Private _dry_concrete_density As Double?
    Private _total_soil_unit_weight As Double?
    Private _bearing_type As String
    Private _nominal_bearing_capacity As Double?
    Private _cohesion As Double?
    Private _friction_angle As Double?
    Private _spt_blow_count As Integer?
    Private _base_friction_factor As Double?
    Private _neglect_depth As Double?
    Private _bearing_distribution_type As Boolean?
    Private _groundwater_depth As Double?
    Private _top_and_bottom_rebar_different As Boolean?
    Private _block_foundation As Boolean?
    Private _rectangular_foundation As Boolean?
    Private _base_plate_distance_above_foundation As Double?
    Private _bolt_circle_bearing_plate_width As Double?
    Private _pier_rebar_quantity As Integer?
    Private _basic_soil_check As Boolean?
    Private _structural_check As Boolean?

    <Category("Pier"), Description(""), DisplayName("Pier Shape")>
    Public Property pier_shape() As String
        Get
            Return Me._pier_shape
        End Get
        Set
            Me._pier_shape = Value
        End Set
    End Property
    <Category("Pier"), Description(""), DisplayName("Pier Diameter")>
    Public Property pier_diameter() As Double?
        Get
            Return Me._pier_diameter
        End Get
        Set
            Me._pier_diameter = Value
        End Set
    End Property
    <Category("Pier"), Description(""), DisplayName("Extension Above Grade")>
    Public Property extension_above_grade() As Double?
        Get
            Return Me._extension_above_grade
        End Get
        Set
            Me._extension_above_grade = Value
        End Set
    End Property
    <Category("Pier"), Description(""), DisplayName("Pier Rebar Size")>
    Public Property pier_rebar_size() As Integer?
        Get
            Return Me._pier_rebar_size
        End Get
        Set
            Me._pier_rebar_size = Value
        End Set
    End Property
    <Category("Pier"), Description(""), DisplayName("Pier Tie Size")>
    Public Property pier_tie_size() As Integer?
        Get
            Return Me._pier_tie_size
        End Get
        Set
            Me._pier_tie_size = Value
        End Set
    End Property
    <Category("Pier"), Description(""), DisplayName("Pier Tie Quantity")>
    Public Property pier_tie_quantity() As Integer?
        Get
            Return Me._pier_tie_quantity
        End Get
        Set
            Me._pier_tie_quantity = Value
        End Set
    End Property
    <Category("Pier"), Description(""), DisplayName("Pier Reinforcement Type")>
    Public Property pier_reinforcement_type() As String
        Get
            Return Me._pier_reinforcement_type
        End Get
        Set
            Me._pier_reinforcement_type = Value
        End Set
    End Property
    <Category("Pier"), Description(""), DisplayName("Pier Clear Cover")>
    Public Property pier_clear_cover() As Double?
        Get
            Return Me._pier_clear_cover
        End Get
        Set
            Me._pier_clear_cover = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Foundation Depth")>
    Public Property foundation_depth() As Double?
        Get
            Return Me._foundation_depth
        End Get
        Set
            Me._foundation_depth = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Width 1")>
    Public Property pad_width_1() As Double?
        Get
            Return Me._pad_width_1
        End Get
        Set
            Me._pad_width_1 = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Width 2")>
    Public Property pad_width_2() As Double?
        Get
            Return Me._pad_width_2
        End Get
        Set
            Me._pad_width_2 = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Thickness")>
    Public Property pad_thickness() As Double?
        Get
            Return Me._pad_thickness
        End Get
        Set
            Me._pad_thickness = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Rebar Size Top Dir1")>
    Public Property pad_rebar_size_top_dir1() As Integer?
        Get
            Return Me._pad_rebar_size_top_dir1
        End Get
        Set
            Me._pad_rebar_size_top_dir1 = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Rebar Size Bottom Dir1")>
    Public Property pad_rebar_size_bottom_dir1() As Integer?
        Get
            Return Me._pad_rebar_size_bottom_dir1
        End Get
        Set
            Me._pad_rebar_size_bottom_dir1 = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Rebar Size Top Dir2")>
    Public Property pad_rebar_size_top_dir2() As Integer?
        Get
            Return Me._pad_rebar_size_top_dir2
        End Get
        Set
            Me._pad_rebar_size_top_dir2 = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Rebar Size Bottom Dir2")>
    Public Property pad_rebar_size_bottom_dir2() As Integer?
        Get
            Return Me._pad_rebar_size_bottom_dir2
        End Get
        Set
            Me._pad_rebar_size_bottom_dir2 = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Rebar Quantity Top Dir1")>
    Public Property pad_rebar_quantity_top_dir1() As Integer?
        Get
            Return Me._pad_rebar_quantity_top_dir1
        End Get
        Set
            Me._pad_rebar_quantity_top_dir1 = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Rebar Quantity Bottom Dir1")>
    Public Property pad_rebar_quantity_bottom_dir1() As Integer?
        Get
            Return Me._pad_rebar_quantity_bottom_dir1
        End Get
        Set
            Me._pad_rebar_quantity_bottom_dir1 = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Rebar Quantity Top Dir2")>
    Public Property pad_rebar_quantity_top_dir2() As Integer?
        Get
            Return Me._pad_rebar_quantity_top_dir2
        End Get
        Set
            Me._pad_rebar_quantity_top_dir2 = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Rebar Quantity Bottom Dir2")>
    Public Property pad_rebar_quantity_bottom_dir2() As Integer?
        Get
            Return Me._pad_rebar_quantity_bottom_dir2
        End Get
        Set
            Me._pad_rebar_quantity_bottom_dir2 = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Pad Clear Cover")>
    Public Property pad_clear_cover() As Double?
        Get
            Return Me._pad_clear_cover
        End Get
        Set
            Me._pad_clear_cover = Value
        End Set
    End Property
    <Category("Pier and Pad "), Description(""), DisplayName("Rebar Grade")>
    Public Property rebar_grade() As Double?
        Get
            Return Me._rebar_grade
        End Get
        Set
            Me._rebar_grade = Value
        End Set
    End Property
    <Category("Pier and Pad "), Description(""), DisplayName("Concrete Compressive Strength")>
    Public Property concrete_compressive_strength() As Double?
        Get
            Return Me._concrete_compressive_strength
        End Get
        Set
            Me._concrete_compressive_strength = Value
        End Set
    End Property
    <Category("Pier and Pad "), Description(""), DisplayName("Dry Concrete Density")>
    Public Property dry_concrete_density() As Double?
        Get
            Return Me._dry_concrete_density
        End Get
        Set
            Me._dry_concrete_density = Value
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
    <Category("Soil"), Description(""), DisplayName("Bearing Type")>
    Public Property bearing_type() As String
        Get
            Return Me._bearing_type
        End Get
        Set
            Me._bearing_type = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Nominal Bearing Capacity")>
    Public Property nominal_bearing_capacity() As Double?
        Get
            Return Me._nominal_bearing_capacity
        End Get
        Set
            Me._nominal_bearing_capacity = Value
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
    <Category("Soil"), Description(""), DisplayName("Spt Blow Count")>
    Public Property spt_blow_count() As Integer?
        Get
            Return Me._spt_blow_count
        End Get
        Set
            Me._spt_blow_count = Value
        End Set
    End Property
    <Category("Soil"), Description(""), DisplayName("Base Friction Factor")>
    Public Property base_friction_factor() As Double?
        Get
            Return Me._base_friction_factor
        End Get
        Set
            Me._base_friction_factor = Value
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
    <Category("Soil"), Description(""), DisplayName("Bearing Distribution Type")>
    Public Property bearing_distribution_type() As Boolean?
        Get
            Return Me._bearing_distribution_type
        End Get
        Set
            Me._bearing_distribution_type = Value
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
    <Category("Pier and Pad "), Description(""), DisplayName("Top And Bottom Rebar Different")>
    Public Property top_and_bottom_rebar_different() As Boolean?
        Get
            Return Me._top_and_bottom_rebar_different
        End Get
        Set
            Me._top_and_bottom_rebar_different = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Block Foundation")>
    Public Property block_foundation() As Boolean?
        Get
            Return Me._block_foundation
        End Get
        Set
            Me._block_foundation = Value
        End Set
    End Property
    <Category("Pad"), Description(""), DisplayName("Rectangular Foundation")>
    Public Property rectangular_foundation() As Boolean?
        Get
            Return Me._rectangular_foundation
        End Get
        Set
            Me._rectangular_foundation = Value
        End Set
    End Property
    <Category("Pier"), Description(""), DisplayName("Base Plate Distance Above Foundation")>
    Public Property base_plate_distance_above_foundation() As Double?
        Get
            Return Me._base_plate_distance_above_foundation
        End Get
        Set
            Me._base_plate_distance_above_foundation = Value
        End Set
    End Property
    <Category("Pier"), Description(""), DisplayName("Bolt Circle Bearing Plate Width")>
    Public Property bolt_circle_bearing_plate_width() As Double?
        Get
            Return Me._bolt_circle_bearing_plate_width
        End Get
        Set
            Me._bolt_circle_bearing_plate_width = Value
        End Set
    End Property
    <Category("Pier"), Description(""), DisplayName("Pier Rebar Quantity")>
    Public Property pier_rebar_quantity() As Integer?
        Get
            Return Me._pier_rebar_quantity
        End Get
        Set
            Me._pier_rebar_quantity = Value
        End Set
    End Property
    <Category("Pier and Pad "), Description(""), DisplayName("Basic Soil Check")>
    Public Property basic_soil_check() As Boolean?
        Get
            Return Me._basic_soil_check
        End Get
        Set
            Me._basic_soil_check = Value
        End Set
    End Property
    <Category("Pier and Pad "), Description(""), DisplayName("Structural Check")>
    Public Property structural_check() As Boolean?
        Get
            Return Me._structural_check
        End Get
        Set
            Me._structural_check = Value
        End Set
    End Property



#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal dr As DataRow, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        ''''''Customize for each foundation type'''''
        Me.ID = DBtoNullableInt(dr.Item("ID"))
        Me.bus_unit = DBtoStr(dr.Item("bus_unit"))
        Me.structure_id = DBtoStr(dr.Item("structure_id"))
        Me.pier_shape = DBtoStr(dr.Item("pier_shape"))
        Me.pier_diameter = DBtoNullableDbl(dr.Item("pier_diameter"))
        Me.extension_above_grade = DBtoNullableDbl(dr.Item("extension_above_grade"))
        Me.pier_rebar_size = DBtoNullableInt(dr.Item("pier_rebar_size"))
        Me.pier_tie_size = DBtoNullableInt(dr.Item("pier_tie_size"))
        Me.pier_tie_quantity = DBtoNullableInt(dr.Item("pier_tie_quantity"))
        Me.pier_reinforcement_type = DBtoStr(dr.Item("pier_reinforcement_type"))
        Me.pier_clear_cover = DBtoNullableDbl(dr.Item("pier_clear_cover"))
        Me.foundation_depth = DBtoNullableDbl(dr.Item("foundation_depth"))
        Me.pad_width_1 = DBtoNullableDbl(dr.Item("pad_width_1"))
        Me.pad_width_2 = DBtoNullableDbl(dr.Item("pad_width_2"))
        Me.pad_thickness = DBtoNullableDbl(dr.Item("pad_thickness"))
        Me.pad_rebar_size_top_dir1 = DBtoNullableInt(dr.Item("pad_rebar_size_top_dir1"))
        Me.pad_rebar_size_bottom_dir1 = DBtoNullableInt(dr.Item("pad_rebar_size_bottom_dir1"))
        Me.pad_rebar_size_top_dir2 = DBtoNullableInt(dr.Item("pad_rebar_size_top_dir2"))
        Me.pad_rebar_size_bottom_dir2 = DBtoNullableInt(dr.Item("pad_rebar_size_bottom_dir2"))
        Me.pad_rebar_quantity_top_dir1 = DBtoNullableInt(dr.Item("pad_rebar_quantity_top_dir1"))
        Me.pad_rebar_quantity_bottom_dir1 = DBtoNullableInt(dr.Item("pad_rebar_quantity_bottom_dir1"))
        Me.pad_rebar_quantity_top_dir2 = DBtoNullableInt(dr.Item("pad_rebar_quantity_top_dir2"))
        Me.pad_rebar_quantity_bottom_dir2 = DBtoNullableInt(dr.Item("pad_rebar_quantity_bottom_dir2"))
        Me.pad_clear_cover = DBtoNullableDbl(dr.Item("pad_clear_cover"))
        Me.rebar_grade = DBtoNullableDbl(dr.Item("rebar_grade"))
        Me.concrete_compressive_strength = DBtoNullableDbl(dr.Item("concrete_compressive_strength"))
        Me.dry_concrete_density = DBtoNullableDbl(dr.Item("dry_concrete_density"))
        Me.total_soil_unit_weight = DBtoNullableDbl(dr.Item("total_soil_unit_weight"))
        Me.bearing_type = DBtoStr(dr.Item("bearing_type"))
        Me.nominal_bearing_capacity = DBtoNullableDbl(dr.Item("nominal_bearing_capacity"))
        Me.cohesion = DBtoNullableDbl(dr.Item("cohesion"))
        Me.friction_angle = DBtoNullableDbl(dr.Item("friction_angle"))
        Me.spt_blow_count = DBtoNullableInt(dr.Item("spt_blow_count"))
        Me.base_friction_factor = DBtoNullableDbl(dr.Item("base_friction_factor"))
        Me.neglect_depth = DBtoNullableDbl(dr.Item("neglect_depth"))
        Me.bearing_distribution_type = DBtoNullableBool(dr.Item("bearing_distribution_type"))
        Me.groundwater_depth = DBtoNullableDbl(dr.Item("groundwater_depth"))
        Me.top_and_bottom_rebar_different = DBtoNullableBool(dr.Item("top_and_bottom_rebar_different"))
        Me.block_foundation = DBtoNullableBool(dr.Item("block_foundation"))
        Me.rectangular_foundation = DBtoNullableBool(dr.Item("rectangular_foundation"))
        Me.base_plate_distance_above_foundation = DBtoNullableDbl(dr.Item("base_plate_distance_above_foundation"))
        Me.bolt_circle_bearing_plate_width = DBtoNullableDbl(dr.Item("bolt_circle_bearing_plate_width"))
        Me.pier_rebar_quantity = DBtoNullableInt(dr.Item("pier_rebar_quantity"))
        Me.basic_soil_check = DBtoNullableBool(dr.Item("basic_soil_check"))
        Me.structural_check = DBtoNullableBool(dr.Item("structural_check"))
        Me.Version = DBtoStr(dr.Item("tool_version"))
        Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
        Me.process_stage = DBtoStr(dr.Item("process_stage"))
    End Sub 'Generate a pp from EDS

    'Public Sub New(ExcelFilePath As String, Optional BU As String = Nothing, Optional structureID As String = Nothing)
    Public Sub New(ExcelFilePath As String, Optional ByVal Parent As EDSObject = Nothing)
        'If this is being created by another EDSObject (i.e. the Structure) this will pass along the most important identifying data
        If Parent IsNot Nothing Then Me.Absorb(Parent)

        ''''''Customize for each foundation type'''''
        Dim excelDS As New DataSet

        For Each item As EXCELDTParameter In ExcelDTParams
            'Get additional tables from excel file 
            Try
                excelDS.Tables.Add(ExcelDatasourceToDataTable(GetExcelDataSource(ExcelFilePath, item.xlsSheet, item.xlsRange), item.xlsDatatable))
            Catch ex As Exception
                Debug.Print(String.Format("Failed to create datatable for: {0}, {1}, {2}", IO.Path.GetFileName(ExcelFilePath), item.xlsSheet, item.xlsRange))
            End Try
        Next

        If excelDS.Tables.Contains("Pier and Pad General Details EXCEL") Then
            Dim dr = excelDS.Tables("Pier and Pad General Details EXCEL").Rows(0)

            Me.ID = DBtoNullableInt(dr.Item("pp_id"))
            Me.pier_shape = DBtoStr(dr.Item("pier_shape"))
            Me.pier_diameter = DBtoNullableDbl(dr.Item("pier_diameter"))
            Me.extension_above_grade = DBtoNullableDbl(dr.Item("extension_above_grade"))
            Me.pier_rebar_size = DBtoNullableInt(dr.Item("pier_rebar_size"))
            Me.pier_tie_size = DBtoNullableInt(dr.Item("pier_tie_size"))
            Me.pier_tie_quantity = DBtoNullableInt(dr.Item("pier_tie_quantity"))
            Me.pier_reinforcement_type = DBtoStr(dr.Item("pier_reinforcement_type"))
            Me.pier_clear_cover = DBtoNullableDbl(dr.Item("pier_clear_cover"))
            Me.foundation_depth = DBtoNullableDbl(dr.Item("foundation_depth"))
            Me.pad_width_1 = DBtoNullableDbl(dr.Item("pad_width_1"))
            Me.pad_width_2 = DBtoNullableDbl(dr.Item("pad_width_2"))
            Me.pad_thickness = DBtoNullableDbl(dr.Item("pad_thickness"))
            Me.pad_rebar_size_top_dir1 = DBtoNullableInt(dr.Item("pad_rebar_size_top_dir1"))
            Me.pad_rebar_size_bottom_dir1 = DBtoNullableInt(dr.Item("pad_rebar_size_bottom_dir1"))
            Me.pad_rebar_size_top_dir2 = DBtoNullableInt(dr.Item("pad_rebar_size_top_dir2"))
            Me.pad_rebar_size_bottom_dir2 = DBtoNullableInt(dr.Item("pad_rebar_size_bottom_dir2"))
            Me.pad_rebar_quantity_top_dir1 = DBtoNullableInt(dr.Item("pad_rebar_quantity_top_dir1"))
            Me.pad_rebar_quantity_bottom_dir1 = DBtoNullableInt(dr.Item("pad_rebar_quantity_bottom_dir1"))
            Me.pad_rebar_quantity_top_dir2 = DBtoNullableInt(dr.Item("pad_rebar_quantity_top_dir2"))
            Me.pad_rebar_quantity_bottom_dir2 = DBtoNullableInt(dr.Item("pad_rebar_quantity_bottom_dir2"))
            Me.pad_clear_cover = DBtoNullableDbl(dr.Item("pad_clear_cover"))
            Me.rebar_grade = DBtoNullableDbl(dr.Item("rebar_grade"))
            Me.concrete_compressive_strength = DBtoNullableDbl(dr.Item("concrete_compressive_strength"))
            Me.dry_concrete_density = DBtoNullableDbl(dr.Item("dry_concrete_density"))
            Me.total_soil_unit_weight = DBtoNullableDbl(dr.Item("total_soil_unit_weight"))
            Me.bearing_type = DBtoStr(dr.Item("bearing_type"))
            Me.nominal_bearing_capacity = DBtoNullableDbl(dr.Item("nominal_bearing_capacity"))
            Me.cohesion = DBtoNullableDbl(dr.Item("cohesion"))
            Me.friction_angle = DBtoNullableDbl(dr.Item("friction_angle"))
            Me.spt_blow_count = DBtoNullableInt(dr.Item("spt_blow_count"))
            Me.base_friction_factor = DBtoNullableDbl(dr.Item("base_friction_factor"))
            Me.neglect_depth = DBtoNullableDbl(dr.Item("neglect_depth"))
            Me.bearing_distribution_type = DBtoNullableBool(dr.Item("bearing_distribution_type"))
            Me.groundwater_depth = DBtoNullableDbl(dr.Item("groundwater_depth"))
            Me.top_and_bottom_rebar_different = DBtoNullableBool(dr.Item("top_and_bottom_rebar_different"))
            Me.block_foundation = DBtoNullableBool(dr.Item("block_foundation"))
            Me.rectangular_foundation = DBtoNullableBool(dr.Item("rectangular_foundation"))
            Me.base_plate_distance_above_foundation = DBtoNullableDbl(dr.Item("base_plate_distance_above_foundation"))
            Me.bolt_circle_bearing_plate_width = DBtoNullableDbl(dr.Item("bolt_circle_bearing_plate_width"))
            Me.pier_rebar_quantity = DBtoNullableInt(dr.Item("pier_rebar_quantity"))
            Me.basic_soil_check = DBtoNullableBool(dr.Item("basic_soil_check"))
            Me.structural_check = DBtoNullableBool(dr.Item("structural_check"))
            Me.Version = DBtoStr(dr.Item("tool_version"))
            'Me.modified_person_id = DBtoNullableInt(dr.Item("modified_person_id"))
            'Me.process_stage = DBtoStr(dr.Item("process_stage"))
        End If

        If excelDS.Tables.Contains("Pier and Pad General Results EXCEL") Then

            For Each Row As DataRow In excelDS.Tables("Pier and Pad General Results EXCEL").Rows

                'For Tools with multiple foundation or sub items, use Row.Item("ID") or add a local_ID column to filter which results should be associated with each foundation

                Me.Results.Add(New EDSResult(Row, Me))

            Next

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

            If IsNothing(Me.groundwater_depth) OrElse Me.groundwater_depth.Value = -1 Then
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

            If Not IsNothing(Me.basic_soil_check) Then
                .Worksheets("Input").Range("SoilInteractionBoolean").Value = CType(Me.basic_soil_check, Boolean)
            End If

            If Not IsNothing(Me.structural_check) Then
                .Worksheets("Input").Range("StructuralCheckBoolean").Value = CType(Me.structural_check, Boolean)
            End If
        End With

    End Sub

#End Region

#Region "Save to EDS"

    Public Overrides Function SQLInsertValues() As String
        SQLInsertValues = ""

        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bus_unit.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structure_id.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_shape.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_diameter.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.extension_above_grade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_rebar_size.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_tie_size.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_tie_quantity.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_reinforcement_type.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_clear_cover.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.foundation_depth.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_width_1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_width_2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_thickness.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_size_top_dir1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_size_bottom_dir1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_size_top_dir2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_size_bottom_dir2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_top_dir1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_bottom_dir1.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_top_dir2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_rebar_quantity_bottom_dir2.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pad_clear_cover.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rebar_grade.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.concrete_compressive_strength.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.dry_concrete_density.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.total_soil_unit_weight.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bearing_type.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.nominal_bearing_capacity.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.cohesion.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.friction_angle.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.spt_blow_count.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.base_friction_factor.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.neglect_depth.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bearing_distribution_type.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.groundwater_depth.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.top_and_bottom_rebar_different.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.block_foundation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.rectangular_foundation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.base_plate_distance_above_foundation.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.bolt_circle_bearing_plate_width.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.pier_rebar_quantity.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.basic_soil_check.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.structural_check.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.Version.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.modified_person_id.NullableToString.FormatDBValue)
        SQLInsertValues = SQLInsertValues.AddtoDBString(Me.process_stage.NullableToString.FormatDBValue)

        Return SQLInsertValues
    End Function

    Public Overrides Function SQLInsertFields() As String
        SQLInsertFields = ""

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
        SQLInsertFields = SQLInsertFields.AddtoDBString("pier_rebar_quantity")
        SQLInsertFields = SQLInsertFields.AddtoDBString("basic_soil_check")
        SQLInsertFields = SQLInsertFields.AddtoDBString("structural_check")
        SQLInsertFields = SQLInsertFields.AddtoDBString("tool_version")
        SQLInsertFields = SQLInsertFields.AddtoDBString("modified_person_id")
        SQLInsertFields = SQLInsertFields.AddtoDBString("process_stage")

        Return SQLInsertFields
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        SQLUpdateFieldsandValues = ""

        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bus_unit = " & Me.bus_unit.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structure_id = " & Me.structure_id.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_shape = " & Me.pier_shape.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_diameter = " & Me.pier_diameter.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("extension_above_grade = " & Me.extension_above_grade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_rebar_size = " & Me.pier_rebar_size.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_tie_size = " & Me.pier_tie_size.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_tie_quantity = " & Me.pier_tie_quantity.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_reinforcement_type = " & Me.pier_reinforcement_type.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_clear_cover = " & Me.pier_clear_cover.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("foundation_depth = " & Me.foundation_depth.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_width_1 = " & Me.pad_width_1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_width_2 = " & Me.pad_width_2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_thickness = " & Me.pad_thickness.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_rebar_size_top_dir1 = " & Me.pad_rebar_size_top_dir1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_rebar_size_bottom_dir1 = " & Me.pad_rebar_size_bottom_dir1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_rebar_size_top_dir2 = " & Me.pad_rebar_size_top_dir2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_rebar_size_bottom_dir2 = " & Me.pad_rebar_size_bottom_dir2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_rebar_quantity_top_dir1 = " & Me.pad_rebar_quantity_top_dir1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_rebar_quantity_bottom_dir1 = " & Me.pad_rebar_quantity_bottom_dir1.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_rebar_quantity_top_dir2 = " & Me.pad_rebar_quantity_top_dir2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_rebar_quantity_bottom_dir2 = " & Me.pad_rebar_quantity_bottom_dir2.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pad_clear_cover = " & Me.pad_clear_cover.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rebar_grade = " & Me.rebar_grade.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("concrete_compressive_strength = " & Me.concrete_compressive_strength.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("dry_concrete_density = " & Me.dry_concrete_density.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("total_soil_unit_weight = " & Me.total_soil_unit_weight.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bearing_type = " & Me.bearing_type.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("nominal_bearing_capacity = " & Me.nominal_bearing_capacity.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("cohesion = " & Me.cohesion.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("friction_angle = " & Me.friction_angle.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("spt_blow_count = " & Me.spt_blow_count.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("base_friction_factor = " & Me.base_friction_factor.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("neglect_depth = " & Me.neglect_depth.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bearing_distribution_type = " & Me.bearing_distribution_type.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("groundwater_depth = " & Me.groundwater_depth.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("top_and_bottom_rebar_different = " & Me.top_and_bottom_rebar_different.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("block_foundation = " & Me.block_foundation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("rectangular_foundation = " & Me.rectangular_foundation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("base_plate_distance_above_foundation = " & Me.base_plate_distance_above_foundation.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("bolt_circle_bearing_plate_width = " & Me.bolt_circle_bearing_plate_width.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("pier_rebar_quantity = " & Me.pier_rebar_quantity.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("basic_soil_check = " & Me.basic_soil_check.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("structural_check = " & Me.structural_check.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("tool_version = " & Me.Version.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("modified_person_id = " & Me.modified_person_id.NullableToString.FormatDBValue)
        SQLUpdateFieldsandValues = SQLUpdateFieldsandValues.AddtoDBString("process_stage = " & Me.process_stage.NullableToString.FormatDBValue)

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
        Dim otherToCompare As PierAndPad = TryCast(other, PierAndPad)
        If otherToCompare Is Nothing Then Return False

        Equals = If(Me.pier_shape.CheckChange(otherToCompare.pier_shape, changes, categoryName, "Pier Shape"), Equals, False)
        Equals = If(Me.pier_diameter.CheckChange(otherToCompare.pier_diameter, changes, categoryName, "Pier Diameter"), Equals, False)
        Equals = If(Me.extension_above_grade.CheckChange(otherToCompare.extension_above_grade, changes, categoryName, "Extension Above Grade"), Equals, False)
        Equals = If(Me.pier_rebar_size.CheckChange(otherToCompare.pier_rebar_size, changes, categoryName, "Pier Rebar Size"), Equals, False)
        Equals = If(Me.pier_tie_size.CheckChange(otherToCompare.pier_tie_size, changes, categoryName, "Pier Tie Size"), Equals, False)
        Equals = If(Me.pier_tie_quantity.CheckChange(otherToCompare.pier_tie_quantity, changes, categoryName, "Pier Tie Quantity"), Equals, False)
        Equals = If(Me.pier_reinforcement_type.CheckChange(otherToCompare.pier_reinforcement_type, changes, categoryName, "Pier Reinforcement Type"), Equals, False)
        Equals = If(Me.pier_clear_cover.CheckChange(otherToCompare.pier_clear_cover, changes, categoryName, "Pier Clear Cover"), Equals, False)
        Equals = If(Me.foundation_depth.CheckChange(otherToCompare.foundation_depth, changes, categoryName, "Foundation Depth"), Equals, False)
        Equals = If(Me.pad_width_1.CheckChange(otherToCompare.pad_width_1, changes, categoryName, "Pad Width 1"), Equals, False)
        Equals = If(Me.pad_width_2.CheckChange(otherToCompare.pad_width_2, changes, categoryName, "Pad Width 2"), Equals, False)
        Equals = If(Me.pad_thickness.CheckChange(otherToCompare.pad_thickness, changes, categoryName, "Pad Thickness"), Equals, False)
        Equals = If(Me.pad_rebar_size_top_dir1.CheckChange(otherToCompare.pad_rebar_size_top_dir1, changes, categoryName, "Pad Rebar Size Top Dir1"), Equals, False)
        Equals = If(Me.pad_rebar_size_bottom_dir1.CheckChange(otherToCompare.pad_rebar_size_bottom_dir1, changes, categoryName, "Pad Rebar Size Bottom Dir1"), Equals, False)
        Equals = If(Me.pad_rebar_size_top_dir2.CheckChange(otherToCompare.pad_rebar_size_top_dir2, changes, categoryName, "Pad Rebar Size Top Dir2"), Equals, False)
        Equals = If(Me.pad_rebar_size_bottom_dir2.CheckChange(otherToCompare.pad_rebar_size_bottom_dir2, changes, categoryName, "Pad Rebar Size Bottom Dir2"), Equals, False)
        Equals = If(Me.pad_rebar_quantity_top_dir1.CheckChange(otherToCompare.pad_rebar_quantity_top_dir1, changes, categoryName, "Pad Rebar Quantity Top Dir1"), Equals, False)
        Equals = If(Me.pad_rebar_quantity_bottom_dir1.CheckChange(otherToCompare.pad_rebar_quantity_bottom_dir1, changes, categoryName, "Pad Rebar Quantity Bottom Dir1"), Equals, False)
        Equals = If(Me.pad_rebar_quantity_top_dir2.CheckChange(otherToCompare.pad_rebar_quantity_top_dir2, changes, categoryName, "Pad Rebar Quantity Top Dir2"), Equals, False)
        Equals = If(Me.pad_rebar_quantity_bottom_dir2.CheckChange(otherToCompare.pad_rebar_quantity_bottom_dir2, changes, categoryName, "Pad Rebar Quantity Bottom Dir2"), Equals, False)
        Equals = If(Me.pad_clear_cover.CheckChange(otherToCompare.pad_clear_cover, changes, categoryName, "Pad Clear Cover"), Equals, False)
        Equals = If(Me.rebar_grade.CheckChange(otherToCompare.rebar_grade, changes, categoryName, "Rebar Grade"), Equals, False)
        Equals = If(Me.concrete_compressive_strength.CheckChange(otherToCompare.concrete_compressive_strength, changes, categoryName, "Concrete Compressive Strength"), Equals, False)
        Equals = If(Me.dry_concrete_density.CheckChange(otherToCompare.dry_concrete_density, changes, categoryName, "Dry Concrete Density"), Equals, False)
        Equals = If(Me.total_soil_unit_weight.CheckChange(otherToCompare.total_soil_unit_weight, changes, categoryName, "Total Soil Unit Weight"), Equals, False)
        Equals = If(Me.bearing_type.CheckChange(otherToCompare.bearing_type, changes, categoryName, "Bearing Type"), Equals, False)
        Equals = If(Me.nominal_bearing_capacity.CheckChange(otherToCompare.nominal_bearing_capacity, changes, categoryName, "Nominal Bearing Capacity"), Equals, False)
        Equals = If(Me.cohesion.CheckChange(otherToCompare.cohesion, changes, categoryName, "Cohesion"), Equals, False)
        Equals = If(Me.friction_angle.CheckChange(otherToCompare.friction_angle, changes, categoryName, "Friction Angle"), Equals, False)
        Equals = If(Me.spt_blow_count.CheckChange(otherToCompare.spt_blow_count, changes, categoryName, "Spt Blow Count"), Equals, False)
        Equals = If(Me.base_friction_factor.CheckChange(otherToCompare.base_friction_factor, changes, categoryName, "Base Friction Factor"), Equals, False)
        Equals = If(Me.neglect_depth.CheckChange(otherToCompare.neglect_depth, changes, categoryName, "Neglect Depth"), Equals, False)
        Equals = If(Me.bearing_distribution_type.CheckChange(otherToCompare.bearing_distribution_type, changes, categoryName, "Bearing Distribution Type"), Equals, False)
        Equals = If(Me.groundwater_depth.CheckChange(otherToCompare.groundwater_depth, changes, categoryName, "Groundwater Depth"), Equals, False)
        Equals = If(Me.top_and_bottom_rebar_different.CheckChange(otherToCompare.top_and_bottom_rebar_different, changes, categoryName, "Top And Bottom Rebar Different"), Equals, False)
        Equals = If(Me.block_foundation.CheckChange(otherToCompare.block_foundation, changes, categoryName, "Block Foundation"), Equals, False)
        Equals = If(Me.rectangular_foundation.CheckChange(otherToCompare.rectangular_foundation, changes, categoryName, "Rectangular Foundation"), Equals, False)
        Equals = If(Me.base_plate_distance_above_foundation.CheckChange(otherToCompare.base_plate_distance_above_foundation, changes, categoryName, "Base Plate Distance Above Foundation"), Equals, False)
        Equals = If(Me.bolt_circle_bearing_plate_width.CheckChange(otherToCompare.bolt_circle_bearing_plate_width, changes, categoryName, "Bolt Circle Bearing Plate Width"), Equals, False)
        Equals = If(Me.pier_rebar_quantity.CheckChange(otherToCompare.pier_rebar_quantity, changes, categoryName, "Pier Rebar Quantity"), Equals, False)
        Equals = If(Me.basic_soil_check.CheckChange(otherToCompare.basic_soil_check, changes, categoryName, "Basic Soil Check"), Equals, False)
        Equals = If(Me.structural_check.CheckChange(otherToCompare.structural_check, changes, categoryName, "Structural Check"), Equals, False)
        Equals = If(Me.Version.CheckChange(otherToCompare.Version, changes, categoryName, "Tool Version"), Equals, False)

        Return Equals

    End Function
#End Region

End Class
