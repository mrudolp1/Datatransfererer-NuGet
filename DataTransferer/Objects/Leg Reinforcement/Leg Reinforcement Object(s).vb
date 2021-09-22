Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
Public Class LegReinforcement

#Region "Define"
    Private prop_leg_rein_id As Integer
    Private prop_local_tool_id As Integer?
    Private prop_leg_section_el As String
    Private prop_data_stored As Boolean
    'Private prop_wind_load As Double?
    'Private prop_dead_load As Double?
    Private prop_rein_type As String
    Private prop_leg_load_time_of_mod As Boolean
    Private prop_end_connections As String 'string leaves potential for additional options in the future
    Private prop_leg_crushing As Boolean
    Private prop_applied_load_methodology As String 'string leaves potential for additional options in the future
    Private prop_slenderness_ratio_type As String 'string leaves potential for additional options in the future
    Private prop_intermediate_conn_type As String 'string leaves potential for additional options in the future
    Private prop_intermediate_conn_spacing As Double?
    Private prop_ki_override As Double?
    Private prop_leg_dia As Double?
    Private prop_leg_thickness As Double?
    Private prop_leg_yield_strength As Double?
    Private prop_leg_unbraced_length As Double?
    Private prop_rein_dia As Double?
    Private prop_rein_thickness As Double?
    Private prop_rein_yield_strength As Double?
    Private prop_print_bolton_conn_info As Boolean
    Public Property BoltOnConnections As New List(Of BoltOnConnection)
    Public Property ArbitraryShapes As New List(Of ArbitraryShape)
    'Public Property tnxSectionDatabaseInfos As New List(Of tnxSectionDatabaseInfo)
    'Public Property tnxMaterialDatabaseInfos As New List(Of tnxMaterialDatabaseInfo)

    'assign inputs for leg location and location, or row,  in tool

    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement ID")>
    Public Property leg_rein_id() As Integer
        Get
            Return Me.prop_leg_rein_id
        End Get
        Set
            Me.prop_leg_rein_id = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement Local Tool ID")>
    Public Property local_tool_id() As Integer?
        Get
            Return Me.prop_local_tool_id
        End Get
        Set
            Me.prop_local_tool_id = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Section Elevations")>
    Public Property leg_section_el() As String
        Get
            Return Me.prop_leg_section_el
        End Get
        Set
            Me.prop_leg_section_el = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement Data Stored to Tool")>
    Public Property data_stored() As Boolean
        Get
            Return Me.prop_data_stored
        End Get
        Set
            Me.prop_data_stored = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement Type")>
    Public Property rein_type() As String
        Get
            Return Me.prop_rein_type
        End Get
        Set
            Me.prop_rein_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Load at Time of Mod")>
    Public Property leg_load_time_of_mod() As Boolean
        Get
            Return Me.prop_leg_load_time_of_mod
        End Get
        Set
            Me.prop_leg_load_time_of_mod = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("End Connections")>
    Public Property end_connections() As String
        Get
            Return Me.prop_end_connections
        End Get
        Set
            Me.prop_end_connections = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Crushing")>
    Public Property leg_crushing() As Boolean
        Get
            Return Me.prop_leg_crushing
        End Get
        Set
            Me.prop_leg_crushing = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Applied Load Methodology")>
    Public Property applied_load_methodology() As String
        Get
            Return Me.prop_applied_load_methodology
        End Get
        Set
            Me.prop_applied_load_methodology = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Slenderness Ratio Type")>
    Public Property slenderness_ratio_type() As String
        Get
            Return Me.prop_slenderness_ratio_type
        End Get
        Set
            Me.prop_slenderness_ratio_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Intermediate Connection Type")>
    Public Property intermediate_conn_type() As String
        Get
            Return Me.prop_intermediate_conn_type
        End Get
        Set
            Me.prop_intermediate_conn_type = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Intermediate Connection Spacing")>
    Public Property intermediate_conn_spacing() As Double?
        Get
            Return Me.prop_intermediate_conn_spacing
        End Get
        Set
            Me.prop_intermediate_conn_spacing = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Ki, Existing Reinforcement K Factor, Override")>
    Public Property ki_override() As Double?
        Get
            Return Me.prop_ki_override
        End Get
        Set
            Me.prop_ki_override = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Diameter")>
    Public Property leg_dia() As Double?
        Get
            Return Me.prop_leg_dia
        End Get
        Set
            Me.prop_leg_dia = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Thickness")>
    Public Property leg_thickness() As Double?
        Get
            Return Me.prop_leg_thickness
        End Get
        Set
            Me.prop_leg_thickness = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Yield Strength, Fy")>
    Public Property leg_yield_strength() As Double?
        Get
            Return Me.prop_leg_yield_strength
        End Get
        Set
            Me.prop_leg_yield_strength = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Unbraced Length")>
    Public Property leg_unbraced_length() As Double?
        Get
            Return Me.prop_leg_unbraced_length
        End Get
        Set
            Me.prop_leg_unbraced_length = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement Diameter")>
    Public Property rein_dia() As Double?
        Get
            Return Me.prop_rein_dia
        End Get
        Set
            Me.prop_rein_dia = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement Thickness")>
    Public Property rein_thickness() As Double?
        Get
            Return Me.prop_rein_thickness
        End Get
        Set
            Me.prop_rein_thickness = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Leg Reinforcement Yield Strength, Fy")>
    Public Property rein_yield_strength() As Double?
        Get
            Return Me.prop_rein_yield_strength
        End Get
        Set
            Me.prop_rein_yield_strength = Value
        End Set
    End Property
    <Category("Leg Reinforcement Details"), Description(""), DisplayName("Print Bolt-On Connection Info")>
    Public Property print_bolton_conn_info() As Boolean
        Get
            Return Me.prop_print_bolton_conn_info
        End Get
        Set
            Me.prop_print_bolton_conn_info = Value
        End Set
    End Property
#End Region

#Region "Constructors"
    Public Sub New()
        'Leave method empty
    End Sub

    Public Sub New(ByVal LegReinforcementDataRow As DataRow, refID As Integer)
        'General Leg Reinforcement Details
        Try
            Me.leg_rein_id = CType(LegReinforcementDataRow.Item("leg_rein_id"), Integer)
        Catch
            Me.leg_rein_id = 0
        End Try 'Leg Reinforcement ID
        Try
            If Not IsDBNull(Me.local_tool_id = CType(LegReinforcementDataRow.Item("local_tool_id"), Integer)) Then
                Me.local_tool_id = CType(LegReinforcementDataRow.Item("local_tool_id"), Integer)
            Else
                Me.local_tool_id = Nothing
            End If
        Catch
            Me.local_tool_id = Nothing
        End Try 'Leg Reinforcement Local Tool ID
        Try
            If Not IsDBNull(Me.leg_section_el = CType(LegReinforcementDataRow.Item("leg_section_el"), String)) Then
                Me.leg_section_el = CType(LegReinforcementDataRow.Item("leg_section_el"), String)
            Else
                Me.leg_section_el = Nothing
            End If
        Catch
            Me.leg_section_el = Nothing
        End Try 'Leg Section Elevations
        Try
            If Not IsDBNull(Me.data_stored = CType(LegReinforcementDataRow.Item("data_stored"), Boolean)) Then
                Me.data_stored = CType(LegReinforcementDataRow.Item("data_stored"), Boolean)
            Else
                Me.data_stored = Nothing
            End If
        Catch
            Me.data_stored = Nothing
        End Try 'Leg Reinforcement Data Stored
        Try
            If Not IsDBNull(Me.rein_type = CType(LegReinforcementDataRow.Item("rein_type"), String)) Then
                Me.rein_type = CType(LegReinforcementDataRow.Item("rein_type"), String)
            Else
                Me.rein_type = Nothing
            End If
        Catch
            Me.rein_type = Nothing
        End Try 'Leg Reinforcement Type
        Try
            If Not IsDBNull(Me.leg_load_time_of_mod = CType(LegReinforcementDataRow.Item("leg_load_time_of_mod"), Boolean)) Then
                Me.leg_load_time_of_mod = CType(LegReinforcementDataRow.Item("leg_load_time_of_mod"), Boolean)
            Else
                Me.leg_load_time_of_mod = Nothing
            End If
        Catch
            Me.leg_load_time_of_mod = Nothing
        End Try 'Leg Load at Time of Modification
        Try
            If Not IsDBNull(Me.end_connections = CType(LegReinforcementDataRow.Item("end_connections"), String)) Then
                Me.end_connections = CType(LegReinforcementDataRow.Item("rein_type"), String)
            Else
                Me.end_connections = Nothing
            End If
        Catch
            Me.end_connections = Nothing
        End Try 'Leg Reinforcement End Connection Type
        Try
            If Not IsDBNull(Me.leg_crushing = CType(LegReinforcementDataRow.Item("leg_crushing"), Boolean)) Then
                Me.leg_crushing = CType(LegReinforcementDataRow.Item("leg_crushing"), Boolean)
            Else
                Me.leg_crushing = Nothing
            End If
        Catch
            Me.leg_crushing = Nothing
        End Try 'Leg Crushing Applied
        Try
            If Not IsDBNull(Me.applied_load_methodology = CType(LegReinforcementDataRow.Item("applied_load_methodology"), String)) Then
                Me.applied_load_methodology = CType(LegReinforcementDataRow.Item("applied_load_methodology"), String)
            Else
                Me.applied_load_methodology = Nothing
            End If
        Catch
            Me.applied_load_methodology = Nothing
        End Try 'Applied Load Methodology
        Try
            If Not IsDBNull(Me.slenderness_ratio_type = CType(LegReinforcementDataRow.Item("slenderness_ratio_type"), String)) Then
                Me.slenderness_ratio_type = CType(LegReinforcementDataRow.Item("slenderness_ratio_type"), String)
            Else
                Me.slenderness_ratio_type = Nothing
            End If
        Catch
            Me.slenderness_ratio_type = Nothing
        End Try 'Slenderness Ratio Type
        Try
            If Not IsDBNull(Me.intermediate_conn_type = CType(LegReinforcementDataRow.Item("intermediate_conn_type"), String)) Then
                Me.intermediate_conn_type = CType(LegReinforcementDataRow.Item("intermediate_conn_type"), String)
            Else
                Me.intermediate_conn_type = Nothing
            End If
        Catch
            Me.intermediate_conn_type = Nothing
        End Try 'Intermediate Connection Type
        Try
            If Not IsDBNull(Me.intermediate_conn_spacing = CType(LegReinforcementDataRow.Item("intermediate_conn_spacing"), Double)) Then
                Me.intermediate_conn_spacing = CType(LegReinforcementDataRow.Item("intermediate_conn_spacing"), Double)
            Else
                Me.intermediate_conn_spacing = Nothing
            End If
        Catch
            Me.intermediate_conn_spacing = Nothing
        End Try 'Intermediate Connection Spacing
        Try
            If Not IsDBNull(Me.ki_override = CType(LegReinforcementDataRow.Item("ki_override"), Double)) Then
                Me.ki_override = CType(LegReinforcementDataRow.Item("ki_override"), Double)
            Else
                Me.ki_override = Nothing
            End If
        Catch
            Me.ki_override = Nothing
        End Try 'Ki Override
        Try
            If Not IsDBNull(Me.leg_dia = CType(LegReinforcementDataRow.Item("leg_dia"), Double)) Then
                Me.leg_dia = CType(LegReinforcementDataRow.Item("leg_dia"), Double)
            Else
                Me.leg_dia = Nothing
            End If
        Catch
            Me.leg_dia = Nothing
        End Try 'Leg Diameter
        Try
            If Not IsDBNull(Me.leg_thickness = CType(LegReinforcementDataRow.Item("leg_thickness"), Double)) Then
                Me.leg_thickness = CType(LegReinforcementDataRow.Item("leg_thickness"), Double)
            Else
                Me.leg_thickness = Nothing
            End If
        Catch
            Me.leg_thickness = Nothing
        End Try 'Leg Thickness
        Try
            If Not IsDBNull(Me.leg_yield_strength = CType(LegReinforcementDataRow.Item("leg_yield_strength"), Double)) Then
                Me.leg_yield_strength = CType(LegReinforcementDataRow.Item("leg_yield_strength"), Double)
            Else
                Me.leg_yield_strength = Nothing
            End If
        Catch
            Me.leg_yield_strength = Nothing
        End Try 'Leg Yield Strength
        Try
            If Not IsDBNull(Me.leg_unbraced_length = CType(LegReinforcementDataRow.Item("leg_unbraced_length"), Double)) Then
                Me.leg_unbraced_length = CType(LegReinforcementDataRow.Item("leg_unbraced_length"), Double)
            Else
                Me.leg_unbraced_length = Nothing
            End If
        Catch
            Me.leg_unbraced_length = Nothing
        End Try 'Leg Unbraced Length
        Try
            If Not IsDBNull(Me.rein_dia = CType(LegReinforcementDataRow.Item("rein_dia"), Double)) Then
                Me.rein_dia = CType(LegReinforcementDataRow.Item("rein_dia"), Double)
            Else
                Me.rein_dia = Nothing
            End If
        Catch
            Me.rein_dia = Nothing
        End Try 'Leg Reinforcement Diameter
        Try
            If Not IsDBNull(Me.rein_thickness = CType(LegReinforcementDataRow.Item("rein_thickness"), Double)) Then
                Me.rein_thickness = CType(LegReinforcementDataRow.Item("rein_thickness"), Double)
            Else
                Me.rein_thickness = Nothing
            End If
        Catch
            Me.rein_thickness = Nothing
        End Try 'Leg Reinforcement Thickness
        Try
            If Not IsDBNull(Me.rein_yield_strength = CType(LegReinforcementDataRow.Item("rein_yield_strength"), Double)) Then
                Me.rein_yield_strength = CType(LegReinforcementDataRow.Item("rein_yield_strength"), Double)
            Else
                Me.rein_yield_strength = Nothing
            End If
        Catch
            Me.rein_yield_strength = Nothing
        End Try 'Leg Reinforcement Yield Strength
        Try
            If Not IsDBNull(Me.print_bolton_conn_info = CType(LegReinforcementDataRow.Item("print_bolton_conn_info"), Boolean)) Then
                Me.print_bolton_conn_info = CType(LegReinforcementDataRow.Item("print_bolton_conn_info"), Boolean)
            Else
                Me.print_bolton_conn_info = Nothing
            End If
        Catch
            Me.print_bolton_conn_info = Nothing
        End Try 'Print Bolt-On Connection Info

        For Each BoltOnDataRow As DataRow In ds.Tables("Leg Reinforcement Bolt-On Connections SQL").Rows
            Dim boltOnRefID As Integer = CType(BoltOnDataRow.Item("leg_rein_id"), Integer)

            If boltOnRefID = refID Then
                Me.BoltOnConnections.Add(New BoltOnConnection(BoltOnDataRow))
            End If
        Next 'Add Bolt-On Connection Info to Leg Reinforcement Object

        For Each ArbitraryShapeDataRow As DataRow In ds.Tables("Arbitrary Shape SQL").Rows
            Dim arbShapeRefID As Integer = CType(ArbitraryShapeDataRow.Item("arbitrary_shape_id"), Integer)

            If arbShapeRefID = refID Then
                Me.ArbitraryShapes.Add(New ArbitraryShape(ArbitraryShapeDataRow))
            End If
        Next 'Add Arbitrary Shape to Leg Reinforcement Object

        'For Each tnxSectionDataRow As DataRow In ds.Tables("TNX Section SQL").Rows
        '    Dim tnxSecRefID As Integer = CType(tnxSectionDataRow.Item("tnx_section_id"), Integer)

        '    If tnxSecRefID = refID Then
        '        Me.tnxSectionDatabaseInfos.Add(New tnxSectionDatabaseInfo(tnxSectionDataRow))
        '    End If
        'Next 'Add tnx Section to Leg Reinforcement Object

        'For Each tnxMaterialDataRow As DataRow In ds.Tables("TNX Material SQL").Rows
        '    Dim tnxMatRefID As Integer = CType(tnxMaterialDataRow.Item("tnx_mat_id"), Integer)

        '    If tnxMatRefID = refID Then
        '        Me.tnxMaterialDatabaseInfos.Add(New tnxMaterialDatabaseInfo(tnxMaterialDataRow))
        '    End If
        'Next 'Add tnx Material to Leg Reinforcement Object


    End Sub 'Generate Leg Reinforcement from EDS

    Public Sub New(ByVal LegReinforcementDataRow As DataRow, ByVal refID As Integer, ByVal refcol As String)
        Try
            Me.leg_rein_id = CType(LegReinforcementDataRow.Item("leg_rein_id"), Integer)
        Catch
            Me.leg_rein_id = 0
        End Try 'Leg Reinforcement ID
        Try
            Me.local_tool_id = CType(LegReinforcementDataRow.Item("local_tool_id"), Integer)
        Catch
            Me.local_tool_id = Nothing
        End Try 'Leg Reinforcement Local Tool ID
        Try
            Me.leg_section_el = CType(LegReinforcementDataRow.Item("leg_section_el"), String)
        Catch
            Me.leg_section_el = Nothing
        End Try 'Leg Section Elevations
        Try
            Me.data_stored = CType(LegReinforcementDataRow.Item("data_stored"), Boolean)
        Catch
            Me.data_stored = False
        End Try 'Leg Reinforcement Data Stored
        Try
            Me.rein_type = CType(LegReinforcementDataRow.Item("rein_type"), String)
        Catch
            Me.rein_type = Nothing
        End Try 'Leg Reinforcement Type
        Try
            Me.leg_load_time_of_mod = CType(LegReinforcementDataRow.Item("leg_load_time_of_mod"), Boolean)
        Catch
            Me.leg_load_time_of_mod = True
        End Try 'Leg Load at Time of Modification
        Try
            Me.end_connections = CType(LegReinforcementDataRow.Item("end_connections"), String)
        Catch
            Me.end_connections = Nothing
        End Try 'Leg Reinforcement End Connection Type
        Try
            If CType(LegReinforcementDataRow.Item("leg_crushing"), String) = "No" Then
                Me.leg_crushing = False
            Else
                Me.leg_crushing = True
            End If
            'Me.leg_crushing = CType(LegReinforcementDataRow.Item("leg_crushing"), Boolean)
        Catch
            Me.leg_crushing = True
        End Try 'Leg Crushing Applied
        Try
            Me.applied_load_methodology = CType(LegReinforcementDataRow.Item("applied_load_methodology"), String)
        Catch
            Me.applied_load_methodology = Nothing
        End Try 'Applied Load Methodology
        Try
            Me.slenderness_ratio_type = CType(LegReinforcementDataRow.Item("slenderness_ratio_type"), String)
        Catch
            Me.slenderness_ratio_type = Nothing
        End Try 'Slenderness Ratio Type
        Try
            Me.intermediate_conn_type = CType(LegReinforcementDataRow.Item("intermediate_conn_type"), String)
        Catch
            Me.intermediate_conn_type = Nothing
        End Try 'Intermediate Connection Type
        Try
            Me.intermediate_conn_spacing = CType(LegReinforcementDataRow.Item("intermediate_conn_spacing"), Double)
        Catch
            Me.intermediate_conn_spacing = Nothing
        End Try 'Intermediate Connection Spacing
        Try
            Me.ki_override = CType(LegReinforcementDataRow.Item("ki_override"), Double)
        Catch
            Me.ki_override = Nothing
        End Try 'Ki Override
        Try
            Me.leg_dia = CType(LegReinforcementDataRow.Item("leg_dia"), Double)
        Catch
            Me.leg_dia = Nothing
        End Try 'Leg Diameter
        Try
            Me.leg_thickness = CType(LegReinforcementDataRow.Item("leg_thickness"), Double)
        Catch
            Me.leg_thickness = Nothing
        End Try 'Leg Thickness
        Try
            Me.leg_yield_strength = CType(LegReinforcementDataRow.Item("leg_yield_strength"), Double)
        Catch
            Me.leg_yield_strength = Nothing
        End Try 'Leg Yield Strength
        Try
            Me.leg_unbraced_length = CType(LegReinforcementDataRow.Item("leg_unbraced_length"), Double)
        Catch
            Me.leg_unbraced_length = Nothing
        End Try 'Leg Unbraced Length
        Try
            Me.rein_dia = CType(LegReinforcementDataRow.Item("rein_dia"), Double)
        Catch
            Me.rein_dia = Nothing
        End Try 'Leg Reinforcement Diameter
        Try
            Me.rein_thickness = CType(LegReinforcementDataRow.Item("rein_thickness"), Double)
        Catch
            Me.rein_thickness = Nothing
        End Try 'Leg Reinforcement Thickness
        Try
            Me.rein_yield_strength = CType(LegReinforcementDataRow.Item("rein_yield_strength"), Double)
        Catch
            Me.rein_yield_strength = Nothing
        End Try 'Leg Reinforcement Yield Strength
        Try
            Me.print_bolton_conn_info = CType(LegReinforcementDataRow.Item("print_bolton_conn_info"), Boolean)
        Catch
            Me.print_bolton_conn_info = False
        End Try 'Print Bolt-On Connection Info

        'For Each BoltOnDataRow As DataRow In ds.Tables("Leg Reinforcement Bolt-On Connections SQL").Rows
        '    Dim boltOnRefID As Integer = CType(BoltOnDataRow.Item("leg_rein_id"), Integer)

        '    If boltOnRefID = refID Then
        '        Me.BoltOnConnections.Add(New BoltOnConnection(BoltOnDataRow))
        '    End If
        'Next 'Add Bolt-On Connection Info to Leg Reinforcement Object

        'For Each ArbitraryShapeDataRow As DataRow In ds.Tables("Arbitrary Shape SQL").Rows
        '    Dim arbShapeRefID As Integer = CType(ArbitraryShapeDataRow.Item("arbitrary_shape_id"), Integer)

        '    If arbShapeRefID = refID Then
        '        Me.ArbitraryShapes.Add(New ArbitraryShape(ArbitraryShapeDataRow))
        '    End If
        'Next 'Add Arbitrary Shape to Leg Reinforcement Object

        'For Each tnxSectionDataRow As DataRow In ds.Tables("TNX Section SQL").Rows
        '    Dim tnxSecRefID As Integer = CType(tnxSectionDataRow.Item("tnx_section_id"), Integer)

        '    If tnxSecRefID = refID Then
        '        Me.tnxSectionDatabaseInfos.Add(New tnxSectionDatabaseInfo(tnxSectionDataRow))
        '    End If
        'Next 'Add tnx Section to Leg Reinforcement Object

        'For Each tnxMaterialDataRow As DataRow In ds.Tables("TNX Material SQL").Rows
        '    Dim tnxMatRefID As Integer = CType(tnxMaterialDataRow.Item("tnx_mat_id"), Integer)

        '    If tnxMatRefID = refID Then
        '        Me.tnxMaterialDatabaseInfos.Add(New tnxMaterialDatabaseInfo(tnxMaterialDataRow))
        '    End If
        'Next 'Add tnx Material to Leg Reinforcement Object


    End Sub 'Generate Leg Reinforcement from Excel

#End Region


End Class


#Region "Leg Reinforcement Extras"
Partial Public Class BoltOnConnection
    Private prop_bolton_id As Integer
    Private prop_leg_length_of_tower_section As Double?
    Private prop_split_pipe_length As Double?
    Private prop_set_top_to_bottom As Boolean
    Private prop_qty_flange_bolt_bot As Integer?
    Private prop_bolt_circle_bot As Double?
    Private prop_bolt_orientation_bot As Integer?
    Private prop_qty_flange_bolt_top As Integer?
    Private prop_bolt_circle_top As Double?
    Private prop_bolt_orientation_top As Integer?
    Private prop_threaded_rod_dia_bot As Double?
    Private prop_threaded_rod_mat_bot As String
    Private prop_threaded_rod_qty_bot As Integer?
    Private prop_threaded_rod_unbraced_length_bot As Double?
    Private prop_threaded_rod_dia_top As Double?
    Private prop_threaded_rod_mat_top As String
    Private prop_threaded_rod_qty_top As Integer?
    Private prop_threaded_rod_unbraced_length_top As Double?
    Private prop_stiffener_height_bot As Double?
    Private prop_stiffener_length_bot As Double?
    Private prop_fillet_weld_size_bot As Double?
    Private prop_exx_bot As Double?
    Private prop_flange_thickness_bot As Double?
    Private prop_stiffener_height_top As Double?
    Private prop_stiffener_length_top As Double?
    Private prop_fillet_weld_size_top As Double?
    Private prop_exx_top As Double?
    Private prop_flange_thickness_top As Double?

    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Bolt-On Connection ID")>
    Public Property bolton_id() As Integer
        Get
            Return Me.prop_bolton_id
        End Get
        Set
            Me.prop_bolton_id = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Leg Length of Tower Section")>
    Public Property leg_length_of_tower_section() As Double?
        Get
            Return Me.prop_leg_length_of_tower_section
        End Get
        Set
            Me.prop_leg_length_of_tower_section = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Physical Split Pipe Length")>
    Public Property split_pipe_length() As Double?
        Get
            Return Me.prop_split_pipe_length
        End Get
        Set
            Me.prop_split_pipe_length = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Set Top Design Info Equal to Bottom Design Info")>
    Public Property set_top_to_bottom() As Boolean
        Get
            Return Me.prop_set_top_to_bottom
        End Get
        Set
            Me.prop_set_top_to_bottom = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Quantity of Bottom Flange Bolts")>
    Public Property qty_flange_bolt_bot() As Integer?
        Get
            Return Me.prop_qty_flange_bolt_bot
        End Get
        Set
            Me.prop_qty_flange_bolt_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Bolt Circle of Bottom Flange Bolts")>
    Public Property bolt_circle_bot() As Double?
        Get
            Return Me.prop_bolt_circle_bot
        End Get
        Set
            Me.prop_bolt_circle_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Bolt Orientation of Bottom Flange Bolts")>
    Public Property bolt_orientation_bot() As Integer?
        Get
            Return Me.prop_bolt_orientation_bot
        End Get
        Set
            Me.prop_bolt_orientation_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Quantity of Top Flange Bolts")>
    Public Property qty_flange_bolt_top() As Integer?
        Get
            Return Me.prop_qty_flange_bolt_top
        End Get
        Set
            Me.prop_qty_flange_bolt_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Bolt Circle of Top Flange Bolts")>
    Public Property bolt_circle_top() As Double?
        Get
            Return Me.prop_bolt_circle_top
        End Get
        Set
            Me.prop_bolt_circle_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Bolt Orientation of Top Flange Bolts")>
    Public Property bolt_orientation_top() As Integer?
        Get
            Return Me.prop_bolt_orientation_top
        End Get
        Set
            Me.prop_bolt_orientation_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Threaded Rod Diameter, Bottom Flange Connection")>
    Public Property threaded_rod_dia_bot() As Double?
        Get
            Return Me.prop_threaded_rod_dia_bot
        End Get
        Set
            Me.prop_threaded_rod_dia_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Threaded Rod Material, Bottom Flange Connection")>
    Public Property threaded_rod_mat_bot() As String
        Get
            Return Me.prop_threaded_rod_mat_bot
        End Get
        Set
            Me.prop_threaded_rod_mat_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Threaded Rod Quantity, Bottom Flange Connection")>
    Public Property threaded_rod_qty_bot() As Integer?
        Get
            Return Me.prop_threaded_rod_qty_bot
        End Get
        Set
            Me.prop_threaded_rod_qty_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Threaded Rod Unbraced Length, Bottom Flange Connection")>
    Public Property threaded_rod_unbraced_length_bot() As Double?
        Get
            Return Me.prop_threaded_rod_unbraced_length_bot
        End Get
        Set
            Me.prop_threaded_rod_unbraced_length_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Threaded Rod Diameter, Bottom Flange Connection")>
    Public Property threaded_rod_dia_top() As Double?
        Get
            Return Me.prop_threaded_rod_dia_top
        End Get
        Set
            Me.prop_threaded_rod_dia_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Threaded Rod Material, Top Flange Connection")>
    Public Property threaded_rod_mat_top() As String
        Get
            Return Me.prop_threaded_rod_mat_top
        End Get
        Set
            Me.prop_threaded_rod_mat_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Threaded Rod Quantity, Top Flange Connection")>
    Public Property threaded_rod_qty_top() As Integer?
        Get
            Return Me.prop_threaded_rod_qty_top
        End Get
        Set
            Me.prop_threaded_rod_qty_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Threaded Rod Unbraced Length, Top Flange Connection")>
    Public Property threaded_rod_unbraced_length_top() As Double?
        Get
            Return Me.prop_threaded_rod_unbraced_length_top
        End Get
        Set
            Me.prop_threaded_rod_unbraced_length_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Stiffener Height, Bottom Flange Connection")>
    Public Property stiffener_height_bot() As Double?
        Get
            Return Me.prop_stiffener_height_bot
        End Get
        Set
            Me.prop_stiffener_height_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Stiffener Length, Bottom Flange Connection")>
    Public Property stiffener_length_bot() As Double?
        Get
            Return Me.prop_stiffener_length_bot
        End Get
        Set
            Me.prop_stiffener_length_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Stiffener Fillet Weld Size, Bottom Flange Connection")>
    Public Property fillet_weld_size_bot() As Double?
        Get
            Return Me.prop_fillet_weld_size_bot
        End Get
        Set
            Me.prop_fillet_weld_size_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Stiffener Fillet Weld Strength, Bottom Flange Connection")>
    Public Property exx_bot() As Double?
        Get
            Return Me.prop_exx_bot
        End Get
        Set
            Me.prop_exx_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Flange Thickness, Bottom Flange Connection")>
    Public Property flange_thickness_bot() As Double?
        Get
            Return Me.prop_flange_thickness_bot
        End Get
        Set
            Me.prop_flange_thickness_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Stiffener Height, Top Flange Connection")>
    Public Property stiffener_height_top() As Double?
        Get
            Return Me.prop_stiffener_height_top
        End Get
        Set
            Me.prop_stiffener_height_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Stiffener Length, Top Flange Connection")>
    Public Property stiffener_length_top() As Double?
        Get
            Return Me.prop_stiffener_length_top
        End Get
        Set
            Me.prop_stiffener_length_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Stiffener Fillet Weld Size, Top Flange Connection")>
    Public Property fillet_weld_size_top() As Double?
        Get
            Return Me.prop_fillet_weld_size_top
        End Get
        Set
            Me.prop_fillet_weld_size_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Stiffener Fillet Weld Strength, Top Flange Connection")>
    Public Property exx_top() As Double?
        Get
            Return Me.prop_exx_top
        End Get
        Set
            Me.prop_exx_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Bolt-On Connections"), Description(""), DisplayName("Flange Thickness, Top Flange Connection")>
    Public Property flange_thickness_top() As Double?
        Get
            Return Me.prop_flange_thickness_top
        End Get
        Set
            Me.prop_flange_thickness_top = Value
        End Set
    End Property

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal BoltOnDataRow As DataRow)
        Try
            Me.bolton_id = CType(BoltOnDataRow.Item("bolton_id"), Integer)
        Catch
            Me.bolton_id = 0
        End Try 'Bolt-On Connections ID
        Try
            Me.leg_length_of_tower_section = CType(BoltOnDataRow.Item("leg_length_of_tower_section"), Double)
        Catch
            Me.leg_length_of_tower_section = Nothing
        End Try 'Leg Length of Tower Section
        Try
            Me.split_pipe_length = CType(BoltOnDataRow.Item("split_pipe_length"), Double)
        Catch
            Me.split_pipe_length = Nothing
        End Try 'Physical Split Pipe Length
        Try
            Me.set_top_to_bottom = CType(BoltOnDataRow.Item("set_top_to_bottom"), Boolean)
        Catch
            Me.set_top_to_bottom = Nothing
        End Try 'Set Top Design Info Equal to Bottom Design Info
        Try
            Me.qty_flange_bolt_bot = CType(BoltOnDataRow.Item("qty_flange_bolt_bot"), Integer)
        Catch
            Me.qty_flange_bolt_bot = Nothing
        End Try 'Quantity of Bottom Flange Bolts
        Try
            Me.bolt_circle_bot = CType(BoltOnDataRow.Item("bolt_circle_bot"), Double)
        Catch
            Me.bolt_circle_bot = Nothing
        End Try 'Bolt Circle of Bottom Flange Bolts
        Try
            Me.bolt_orientation_bot = CType(BoltOnDataRow.Item("bolt_orientation_bot"), Integer)
        Catch
            Me.bolt_orientation_bot = Nothing
        End Try 'Bolt Orientation of Bottom Flange Bolts
        Try
            Me.qty_flange_bolt_top = CType(BoltOnDataRow.Item("qty_flange_bolt_top"), Integer)
        Catch
            Me.qty_flange_bolt_top = Nothing
        End Try 'Quantity of Top Flange Bolts
        Try
            Me.bolt_circle_top = CType(BoltOnDataRow.Item("bolt_circle_top"), Double)
        Catch
            Me.bolt_circle_top = Nothing
        End Try 'Bolt Circle of Top Flange Bolts
        Try
            Me.bolt_orientation_top = CType(BoltOnDataRow.Item("bolt_orientation_top"), Integer)
        Catch
            Me.bolt_orientation_top = Nothing
        End Try 'Bolt Orientation of Top Flange Bolts
        Try
            Me.threaded_rod_dia_bot = CType(BoltOnDataRow.Item("threaded_rod_dia_bot"), Double)
        Catch
            Me.threaded_rod_dia_bot = Nothing
        End Try 'Threaded Rod Diameter, Bottom Flange Connection
        Try
            Me.threaded_rod_mat_bot = CType(BoltOnDataRow.Item("threaded_rod_mat_bot"), String)
        Catch
            Me.threaded_rod_mat_bot = Nothing
        End Try 'Threaded Rod Material, Bottom Flange Connection
        Try
            Me.threaded_rod_qty_bot = CType(BoltOnDataRow.Item("threaded_rod_qty_bot"), Integer)
        Catch
            Me.threaded_rod_qty_bot = Nothing
        End Try 'Threaded Rod Quantity, Bottom Flange Connection
        Try
            Me.threaded_rod_unbraced_length_bot = CType(BoltOnDataRow.Item("threaded_rod_unbraced_length_bot"), Double)
        Catch
            Me.threaded_rod_unbraced_length_bot = Nothing
        End Try 'Threaded Rod Unbraced Length, Bottom Flange Connection
        Try
            Me.threaded_rod_dia_top = CType(BoltOnDataRow.Item("threaded_rod_dia_top"), Double)
        Catch
            Me.threaded_rod_dia_top = Nothing
        End Try 'Threaded Rod Diameter, Bottom Flange Connection
        Try
            Me.threaded_rod_mat_top = CType(BoltOnDataRow.Item("threaded_rod_mat_top"), String)
        Catch
            Me.threaded_rod_mat_top = Nothing
        End Try 'Threaded Rod Material, Top Flange Connection
        Try
            Me.threaded_rod_qty_top = CType(BoltOnDataRow.Item("threaded_rod_qty_top"), Integer)
        Catch
            Me.threaded_rod_qty_top = Nothing
        End Try 'Threaded Rod Quantity, Top Flange Connection
        Try
            Me.threaded_rod_unbraced_length_top = CType(BoltOnDataRow.Item("threaded_rod_unbraced_length_top"), Double)
        Catch
            Me.threaded_rod_unbraced_length_top = Nothing
        End Try 'Threaded Rod Unbraced Length, Top Flange Connection
        Try
            Me.stiffener_height_bot = CType(BoltOnDataRow.Item("stiffener_height_bot"), Double)
        Catch
            Me.stiffener_height_bot = Nothing
        End Try 'Stiffener Height, Bottom Flange Connection
        Try
            Me.stiffener_length_bot = CType(BoltOnDataRow.Item("stiffener_length_bot"), Double)
        Catch
            Me.stiffener_length_bot = Nothing
        End Try 'Stiffener Length, Bottom Flange Connection
        Try
            Me.fillet_weld_size_bot = CType(BoltOnDataRow.Item("fillet_weld_size_bot"), Double)
        Catch
            Me.fillet_weld_size_bot = Nothing
        End Try 'Stiffener Fillet Weld Size, Bottom Flange Connection
        Try
            Me.exx_bot = CType(BoltOnDataRow.Item("exx_bot"), Double)
        Catch
            Me.exx_bot = Nothing
        End Try 'Stiffener Fillet Weld Strength, Bottom Flange Connection
        Try
            Me.flange_thickness_bot = CType(BoltOnDataRow.Item("flange_thickness_bot"), Double)
        Catch
            Me.flange_thickness_bot = Nothing
        End Try 'Flange Thickness, Bottom Flange Connection
        Try
            Me.stiffener_height_top = CType(BoltOnDataRow.Item("stiffener_height_top"), Double)
        Catch
            Me.stiffener_height_top = Nothing
        End Try 'Stiffener Height, Top Flange Connection
        Try
            Me.stiffener_length_top = CType(BoltOnDataRow.Item("stiffener_length_top"), Double)
        Catch
            Me.stiffener_length_top = Nothing
        End Try 'Stiffener Length, Top Flange Connection
        Try
            Me.fillet_weld_size_top = CType(BoltOnDataRow.Item("fillet_weld_size_top"), Double)
        Catch
            Me.fillet_weld_size_top = Nothing
        End Try 'Stiffener Fillet Weld Size, Top Flange Connection
        Try
            Me.exx_top = CType(BoltOnDataRow.Item("exx_bot"), Double)
        Catch
            Me.exx_top = Nothing
        End Try 'Stiffener Fillet Weld Strength, Top Flange Connection
        Try
            Me.flange_thickness_top = CType(BoltOnDataRow.Item("flange_thickness_top"), Double)
        Catch
            Me.flange_thickness_top = Nothing
        End Try 'Flange Thickness, Top Flange Connection

    End Sub

End Class

Partial Public Class ArbitraryShape
    Private prop_arb_shape_id As Integer
    Private prop_us_name As String
    Private prop_si_name As String
    Private prop_height As Double?
    Private prop_width As Double?
    Private prop_wind_projected_width As Double?
    Private prop_perimeter As Double?
    Private prop_modulus_of_elasticity As Double?
    Private prop_density As Double?
    Private prop_area As Double?
    Private prop_stress_reduction_factor As Double?
    Private prop_warp_constant As Double?
    Private prop_moment_of_inertia_x As Double?
    Private prop_moment_of_inertia_y As Double?
    Private prop_tors_constant As Double?
    Private prop_elastic_sec_mod_x_top As Double?
    Private prop_elastic_sec_mod_y_left As Double?
    Private prop_elastic_sec_mod_x_bot As Double?
    Private prop_elastic_sec_mod_y_right As Double?
    Private prop_radius_of_gyration_x As Double?
    Private prop_radius_of_gyration_y As Double?
    Private prop_shear_deflection_form_factor_x As Double?
    Private prop_shear_deflection_form_factor_y As Double?
    Private prop_K_factor_adj As Double?
    Private prop_sec_file As String
    Private prop_tnx_us_name As String
    Private prop_tnx_si_name As String
    Private prop_sec_values As String
    Private prop_member_mat_file As String
    Private prop_mat_name As String
    Private prop_mat_values As String

    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Arbitrary Shape ID")>
    Public Property arb_shape_id() As Integer
        Get
            Return Me.prop_arb_shape_id
        End Get
        Set
            Me.prop_arb_shape_id = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("US Name")>
    Public Property us_name() As String
        Get
            Return Me.prop_us_name
        End Get
        Set
            Me.prop_us_name = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("SI Name")>
    Public Property si_name() As String
        Get
            Return Me.prop_si_name
        End Get
        Set
            Me.prop_si_name = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Height")>
    Public Property height() As Double?
        Get
            Return Me.prop_height
        End Get
        Set
            Me.prop_height = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Width")>
    Public Property width() As Double?
        Get
            Return Me.prop_width
        End Get
        Set
            Me.prop_width = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Wind Projected Width")>
    Public Property wind_projected_width() As Double?
        Get
            Return Me.prop_wind_projected_width
        End Get
        Set
            Me.prop_wind_projected_width = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Perimeter")>
    Public Property perimeter() As Double?
        Get
            Return Me.prop_perimeter
        End Get
        Set
            Me.prop_perimeter = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Modulus of Elasticity")>
    Public Property modulus_of_elasticity() As Double?
        Get
            Return Me.prop_modulus_of_elasticity
        End Get
        Set
            Me.prop_modulus_of_elasticity = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Density")>
    Public Property density() As Double?
        Get
            Return Me.prop_density
        End Get
        Set
            Me.prop_density = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Cross-Sectional Area")>
    Public Property area() As Double?
        Get
            Return Me.prop_area
        End Get
        Set
            Me.prop_area = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("QaQs, Local Buckling Interaction Stress Reduction Factors")>
    Public Property stress_reduction_factor() As Double?
        Get
            Return Me.prop_stress_reduction_factor
        End Get
        Set
            Me.prop_stress_reduction_factor = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Cw, Warping Constant")>
    Public Property warp_constant() As Double?
        Get
            Return Me.prop_warp_constant
        End Get
        Set
            Me.prop_warp_constant = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Ix, Moment of Inertia about X-X Axis")>
    Public Property moment_of_inertia_x() As Double?
        Get
            Return Me.prop_moment_of_inertia_x
        End Get
        Set
            Me.prop_moment_of_inertia_x = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Iy, Moment of Inertia about Y-Y Axis")>
    Public Property moment_of_inertia_y() As Double?
        Get
            Return Me.prop_moment_of_inertia_y
        End Get
        Set
            Me.prop_moment_of_inertia_y = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("J, Torsional Constant")>
    Public Property tors_constant() As Double?
        Get
            Return Me.prop_tors_constant
        End Get
        Set
            Me.prop_tors_constant = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Sx_top, Elastic Section Modulus about X-X Axis, Top Direction")>
    Public Property elastic_sec_mod_x_top() As Double?
        Get
            Return Me.prop_elastic_sec_mod_x_top
        End Get
        Set
            Me.prop_elastic_sec_mod_x_top = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Sy_left, Elastic Section Modulus about Y-Y Axis, Left Direction? (TNX specifies this as top)")>
    Public Property elastic_sec_mod_y_left() As Double?
        Get
            Return Me.prop_elastic_sec_mod_y_left
        End Get
        Set
            Me.prop_elastic_sec_mod_y_left = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Sx_bot, Elastic Section Modulus about X-X Axis, Bottom Direction")>
    Public Property elastic_sec_mod_x_bot() As Double?
        Get
            Return Me.prop_elastic_sec_mod_x_bot
        End Get
        Set
            Me.prop_elastic_sec_mod_x_bot = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("Sy_right, Elastic Section Modulus about Y-Y Axis, Right Direction? (TNX specifies this as bottom)")>
    Public Property elastic_sec_mod_y_right() As Double?
        Get
            Return Me.prop_elastic_sec_mod_y_right
        End Get
        Set
            Me.prop_elastic_sec_mod_y_right = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("rx, Radius of Gyration about X-X Axis")>
    Public Property radius_of_gyration_x() As Double?
        Get
            Return Me.prop_radius_of_gyration_x
        End Get
        Set
            Me.prop_radius_of_gyration_x = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("ry, Radius of Gyration about Y-Y Axis")>
    Public Property radius_of_gyration_y() As Double?
        Get
            Return Me.prop_radius_of_gyration_y
        End Get
        Set
            Me.prop_radius_of_gyration_y = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("SFx, Shear Deflection Form Factor about X-X Axis")>
    Public Property shear_deflection_form_factor_x() As Double?
        Get
            Return Me.prop_shear_deflection_form_factor_x
        End Get
        Set
            Me.prop_shear_deflection_form_factor_x = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("SFy, Shear Deflection Form Factor about Y-Y Axis")>
    Public Property shear_deflection_form_factor_y() As Double?
        Get
            Return Me.prop_shear_deflection_form_factor_y
        End Get
        Set
            Me.prop_shear_deflection_form_factor_y = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("K Factor Adjustment. Allows TNX to match tool result.")>
    Public Property K_factor_adj() As Double?
        Get
            Return Me.prop_K_factor_adj
        End Get
        Set
            Me.prop_K_factor_adj = Value
        End Set
    End Property
    '<Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("TNX Section Database Info ID")>
    'Public Property tnx_section_db_info_id() As Integer
    '    Get
    '        Return Me.prop_tnx_section_db_info_id
    '    End Get
    '    Set
    '        Me.prop_tnx_section_db_info_id = Value
    '    End Set
    'End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("TNX Section File= (code to insert into .eri file)")>
    Public Property sec_file() As String
        Get
            Return Me.prop_sec_file
        End Get
        Set
            Me.prop_sec_file = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("TNX Section USName= (code to insert into .eri file)")>
    Public Property tnx_us_name() As String
        Get
            Return Me.prop_tnx_us_name
        End Get
        Set
            Me.prop_tnx_us_name = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("TNX Section SIName= (code to insert into .eri file)")>
    Public Property tnx_si_name() As String
        Get
            Return Me.prop_tnx_si_name
        End Get
        Set
            Me.prop_tnx_si_name = Value
        End Set
    End Property
    <Category("Leg Reinforcement Arbitrary Shape"), Description(""), DisplayName("TNX Section Values= (code to insert into .eri file)")>
    Public Property sec_values() As String
        Get
            Return Me.prop_sec_values
        End Get
        Set
            Me.prop_sec_values = Value
        End Set
    End Property
    <Category("Leg Reinforcement TNX Material Database Info"), Description(""), DisplayName("TNX Material Database File= (code to insert into .eri file)")>
    Public Property member_mat_file() As String
        Get
            Return Me.prop_member_mat_file
        End Get
        Set
            Me.prop_member_mat_file = Value
        End Set
    End Property
    <Category("Leg Reinforcement TNX Material Database Info"), Description(""), DisplayName("TNX Material Database Name= (code to insert into .eri file)")>
    Public Property mat_name() As String
        Get
            Return Me.prop_mat_name
        End Get
        Set
            Me.prop_mat_name = Value
        End Set
    End Property
    <Category("Leg Reinforcement TNX Material Database Info"), Description(""), DisplayName("TNX Material Database Values= (code to insert into .eri file)")>
    Public Property mat_values() As String
        Get
            Return Me.prop_mat_values
        End Get
        Set
            Me.prop_mat_values = Value
        End Set
    End Property

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal ArbitraryShapeDataRow As DataRow)
        Try
            Me.arb_shape_id = CType(ArbitraryShapeDataRow.Item("arb_shape_id"), Integer)
        Catch
            Me.arb_shape_id = 0
        End Try 'Arbitrary Shape ID
        Try
            Me.us_name = CType(ArbitraryShapeDataRow.Item("us_name"), String)
        Catch
            Me.us_name = Nothing
        End Try 'US Name
        Try
            Me.si_name = CType(ArbitraryShapeDataRow.Item("si_name"), String)
        Catch
            Me.si_name = Nothing
        End Try 'SI Name
        Try
            Me.height = CType(ArbitraryShapeDataRow.Item("height"), Double)
        Catch
            Me.height = Nothing
        End Try 'Height
        Try
            Me.width = CType(ArbitraryShapeDataRow.Item("width"), Double)
        Catch
            Me.width = Nothing
        End Try 'Width
        Try
            Me.wind_projected_width = CType(ArbitraryShapeDataRow.Item("wind_projected_width"), Double)
        Catch
            Me.wind_projected_width = Nothing
        End Try 'Wind Projected Width
        Try
            Me.perimeter = CType(ArbitraryShapeDataRow.Item("perimeter"), Double)
        Catch
            Me.perimeter = Nothing
        End Try 'Perimeter
        Try
            Me.modulus_of_elasticity = CType(ArbitraryShapeDataRow.Item("modulus_of_elasticity"), Double)
        Catch
            Me.modulus_of_elasticity = Nothing
        End Try 'Modulus of Elasticity
        Try
            Me.density = CType(ArbitraryShapeDataRow.Item("density"), Double)
        Catch
            Me.density = Nothing
        End Try 'Density
        Try
            Me.area = CType(ArbitraryShapeDataRow.Item("area"), Double)
        Catch
            Me.area = Nothing
        End Try 'Cross-Sectional Area
        Try
            Me.stress_reduction_factor = CType(ArbitraryShapeDataRow.Item("stress_reduction_factor"), Double)
        Catch
            Me.stress_reduction_factor = Nothing
        End Try 'Local Buckling Interaction Stress Reduction Factors
        Try
            Me.warp_constant = CType(ArbitraryShapeDataRow.Item("warp_constant"), Double)
        Catch
            Me.warp_constant = Nothing
        End Try 'Warping Constant
        Try
            Me.moment_of_inertia_x = CType(ArbitraryShapeDataRow.Item("moment_of_inertia_x"), Double)
        Catch
            Me.moment_of_inertia_x = Nothing
        End Try 'Moment of Inertia about X-X Axis
        Try
            Me.moment_of_inertia_y = CType(ArbitraryShapeDataRow.Item("moment_of_inertia_y"), Double)
        Catch
            Me.moment_of_inertia_y = Nothing
        End Try 'Moment of Inertia about Y-Y Axis
        Try
            Me.tors_constant = CType(ArbitraryShapeDataRow.Item("tors_constant"), Double)
        Catch
            Me.tors_constant = Nothing
        End Try 'Torsional Constant
        Try
            Me.elastic_sec_mod_x_top = CType(ArbitraryShapeDataRow.Item("elastic_sec_mod_x_top"), Double)
        Catch
            Me.elastic_sec_mod_x_top = Nothing
        End Try 'Sx_top, Elastic Section Modulus about X-X Axis, Top Direction
        Try
            Me.elastic_sec_mod_y_left = CType(ArbitraryShapeDataRow.Item("elastic_sec_mod_y_left"), Double)
        Catch
            Me.elastic_sec_mod_y_left = Nothing
        End Try 'Sy_left, Elastic Section Modulus about Y-Y Axis, Left Direction? (TNX specifies this as top)
        Try
            Me.elastic_sec_mod_x_bot = CType(ArbitraryShapeDataRow.Item("elastic_sec_mod_x_bot"), Double)
        Catch
            Me.elastic_sec_mod_x_bot = Nothing
        End Try 'Sx_bot, Elastic Section Modulus about X-X Axis, Bottom Direction
        Try
            Me.elastic_sec_mod_y_right = CType(ArbitraryShapeDataRow.Item("elastic_sec_mod_y_right"), Double)
        Catch
            Me.elastic_sec_mod_y_right = Nothing
        End Try 'Sy_right, Elastic Section Modulus about Y-Y Axis, Right Direction? (TNX specifies this as top)
        Try
            Me.radius_of_gyration_x = CType(ArbitraryShapeDataRow.Item("radius_of_gyration_x"), Double)
        Catch
            Me.radius_of_gyration_x = Nothing
        End Try 'rx, Radius of Gyration about X-X Axis
        Try
            Me.radius_of_gyration_y = CType(ArbitraryShapeDataRow.Item("radius_of_gyration_y"), Double)
        Catch
            Me.radius_of_gyration_y = Nothing
        End Try 'ry, Radius of Gyration about Y-Y Axis
        Try
            Me.shear_deflection_form_factor_x = CType(ArbitraryShapeDataRow.Item("shear_deflection_form_factor_x"), Double)
        Catch
            Me.shear_deflection_form_factor_x = Nothing
        End Try 'SFx, Shear Deflection Form Factor about X-X Axis
        Try
            Me.shear_deflection_form_factor_y = CType(ArbitraryShapeDataRow.Item("shear_deflection_form_factor_y"), Double)
        Catch
            Me.shear_deflection_form_factor_y = Nothing
        End Try 'SFy, Shear Deflection Form Factor about Y-Y Axis
        Try
            Me.K_factor_adj = CType(ArbitraryShapeDataRow.Item("K_factor_adj"), Double)
        Catch
            Me.K_factor_adj = Nothing
        End Try 'K Factor Adjustment. Allows TNX to match tool result.
        Try
            Me.sec_file = CType(ArbitraryShapeDataRow.Item("sec_file"), String)
        Catch
            Me.sec_file = Nothing
        End Try 'TNX Section File= (code to insert into .eri file)
        Try
            Me.tnx_us_name = CType(ArbitraryShapeDataRow.Item("tnx_us_name"), String)
        Catch
            Me.tnx_us_name = Nothing
        End Try 'TNX Section USName= (code to insert into .eri file)
        Try
            Me.tnx_si_name = CType(ArbitraryShapeDataRow.Item("tnx_si_name"), String)
        Catch
            Me.tnx_si_name = Nothing
        End Try 'TNX Section SIName= (code to insert into .eri file)
        Try
            Me.sec_values = CType(ArbitraryShapeDataRow.Item("sec_values"), String)
        Catch
            Me.sec_values = Nothing
        End Try 'TNX Section Values= (code to insert into .eri file)
        Try
            Me.member_mat_file = CType(ArbitraryShapeDataRow.Item("member_mat_file"), String)
        Catch
            Me.member_mat_file = Nothing
        End Try 'TNX Material Database File= (code to insert into .eri file)
        Try
            Me.mat_name = CType(ArbitraryShapeDataRow.Item("mat_name"), String)
        Catch
            Me.mat_name = Nothing
        End Try 'TNX Material Database Name= (code to insert into .eri file)
        Try
            Me.mat_values = CType(ArbitraryShapeDataRow.Item("mat_values"), String)
        Catch
            Me.mat_values = Nothing
        End Try 'TNX Material Database Values= (code to insert into .eri file)
    End Sub

End Class

'Partial Public Class tnxSectionDatabaseInfo
'    Private prop_tnx_section_db_info_id As Integer
'    Private prop_file As String
'    Private prop_us_name As String
'    Private prop_si_name As String
'    Private prop_sec_values As String

'    <Category("Leg Reinforcement TNX Section Database Info"), Description(""), DisplayName("TNX Section Database Info ID")>
'    Public Property tnx_section_db_info_id() As Integer
'        Get
'            Return Me.prop_tnx_section_db_info_id
'        End Get
'        Set
'            Me.prop_tnx_section_db_info_id = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement TNX Section Database Info"), Description(""), DisplayName("TNX Section File= (code to insert into .eri file)")>
'    Public Property file() As String
'        Get
'            Return Me.prop_file
'        End Get
'        Set
'            Me.prop_file = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement TNX Section Database Info"), Description(""), DisplayName("TNX Section USName= (code to insert into .eri file)")>
'    Public Property us_name() As String
'        Get
'            Return Me.prop_us_name
'        End Get
'        Set
'            Me.prop_us_name = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement TNX Section Database Info"), Description(""), DisplayName("TNX Section SIName= (code to insert into .eri file)")>
'    Public Property si_name() As String
'        Get
'            Return Me.prop_si_name
'        End Get
'        Set
'            Me.prop_si_name = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement TNX Section Database Info"), Description(""), DisplayName("TNX Section Values= (code to insert into .eri file)")>
'    Public Property sec_values() As String
'        Get
'            Return Me.prop_sec_values
'        End Get
'        Set
'            Me.prop_sec_values = Value
'        End Set
'    End Property

'    Sub New()
'        'Leave method empty
'    End Sub

'    Sub New(ByVal tnxSectionDataRow As DataRow)
'        Try
'            Me.tnx_section_db_info_id = CType(tnxSectionDataRow.Item("tnx_section_db_info_id"), Integer)
'        Catch
'            Me.tnx_section_db_info_id = 0
'        End Try 'TNX Section Database Info ID
'        Try
'            Me.file = CType(tnxSectionDataRow.Item("file"), String)
'        Catch
'            Me.file = Nothing
'        End Try 'TNX Section File= (code to insert into .eri file)
'        Try
'            Me.us_name = CType(tnxSectionDataRow.Item("us_name"), String)
'        Catch
'            Me.us_name = Nothing
'        End Try 'TNX Section USName= (code to insert into .eri file)
'        Try
'            Me.si_name = CType(tnxSectionDataRow.Item("si_name"), String)
'        Catch
'            Me.si_name = Nothing
'        End Try 'TNX Section SIName= (code to insert into .eri file)
'        Try
'            Me.sec_values = CType(tnxSectionDataRow.Item("sec_values"), String)
'        Catch
'            Me.sec_values = Nothing
'        End Try 'TNX Section Values= (code to insert into .eri file)
'    End Sub

'End Class

'Partial Public Class tnxMaterialDatabaseInfo
'    Private prop_tnx_mat_db_info_id As Integer
'    Private prop_member_mat_file As String
'    Private prop_mat_name As String
'    Private prop_mat_values As String

'    <Category("Leg Reinforcement TNX Material Database Info"), Description(""), DisplayName("TNX Material Database Info ID")>
'    Public Property tnx_mat_db_info_id() As Integer
'        Get
'            Return Me.prop_tnx_mat_db_info_id
'        End Get
'        Set
'            Me.prop_tnx_mat_db_info_id = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement TNX Material Database Info"), Description(""), DisplayName("TNX Material Database File= (code to insert into .eri file)")>
'    Public Property member_mat_file() As String
'        Get
'            Return Me.prop_member_mat_file
'        End Get
'        Set
'            Me.prop_member_mat_file = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement TNX Material Database Info"), Description(""), DisplayName("TNX Material Database Name= (code to insert into .eri file)")>
'    Public Property mat_name() As String
'        Get
'            Return Me.prop_mat_name
'        End Get
'        Set
'            Me.prop_mat_name = Value
'        End Set
'    End Property
'    <Category("Leg Reinforcement TNX Material Database Info"), Description(""), DisplayName("TNX Material Database Values= (code to insert into .eri file)")>
'    Public Property mat_values() As String
'        Get
'            Return Me.prop_mat_values
'        End Get
'        Set
'            Me.prop_mat_values = Value
'        End Set
'    End Property

'    Sub New()
'        'Leave method empty
'    End Sub

'    Public Sub New(ByVal tnxMaterialDataRow As DataRow)
'        Try
'            Me.tnx_mat_db_info_id = CType(tnxMaterialDataRow.Item("tnx_mat_db_info_id"), Integer)
'        Catch
'            Me.tnx_mat_db_info_id = 0
'        End Try 'TNX Material Database Info ID
'        Try
'            Me.member_mat_file = CType(tnxMaterialDataRow.Item("member_mat_file"), String)
'        Catch
'            Me.member_mat_file = Nothing
'        End Try 'TNX Material Database File= (code to insert into .eri file)
'        Try
'            Me.mat_name = CType(tnxMaterialDataRow.Item("mat_name"), String)
'        Catch
'            Me.mat_name = Nothing
'        End Try 'TNX Material Database Name= (code to insert into .eri file)
'        Try
'            Me.mat_values = CType(tnxMaterialDataRow.Item("mat_values"), String)
'        Catch
'            Me.mat_values = Nothing
'        End Try 'TNX Material Database Values= (code to insert into .eri file)
'    End Sub

'End Class
#End Region