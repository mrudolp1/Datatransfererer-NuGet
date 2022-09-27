'Option Strict On

'Imports System.ComponentModel
'Imports System.Data
'Imports DevExpress.Spreadsheet

'Partial Public Class CCIplate
'#Region "Define"
'    Private prop_connection_group_id As Integer?
'    Public Property plate_connections As New List(Of CCIplatePlateConnections)
'    Public Property plate_details As New List(Of CCIplatePlateDetails)
'    Public Property base_plate_options As New List(Of CCIplateBasePlateOptions)

'    <Category("Connection Group"), Description(""), DisplayName("Connection Group ID")>
'    Public Property connection_group_id() As Integer?
'        Get
'            Return Me.prop_connection_group_id
'        End Get
'        Set
'            Me.prop_connection_group_id = Value
'        End Set
'    End Property
'#End Region
'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal ConnectionGroupDataRow As DataRow, refID As Integer)

'        Try
'            Me.connection_group_id = refID
'        Catch
'            Me.connection_group_id = 0
'        End Try 'Connection_Group_ID


'        For Each PlateConnectionsDataRow As DataRow In ds.Tables("CCIplate Plate Connections SQL").Rows
'            Dim connRefID As Integer = CType(PlateConnectionsDataRow.Item("connection_group_id"), Integer)
'            If connRefID = refID Then
'                Me.plate_connections.Add(New CCIplatePlateConnections(PlateConnectionsDataRow))
'            End If
'        Next 'Add plate connections

'        For Each PlateDetailsDataRow As DataRow In ds.Tables("CCIplate Plate Details SQL").Rows
'            Dim detRefID As Integer = CType(PlateDetailsDataRow.Item("connection_id"), Integer)
'            If detRefID = refID Then
'                Me.plate_details.Add(New CCIplatePlateDetails(PlateDetailsDataRow))
'            End If
'        Next 'Add plate connections

'        For Each BasePlateOptionsDataRow As DataRow In ds.Tables("CCIplate Base Plate Options SQL").Rows
'            Dim detRefID As Integer = CType(BasePlateOptionsDataRow.Item("base_plate_option_id"), Integer)
'            If detRefID = refID Then
'                Me.base_plate_options.Add(New CCIplateBasePlateOptions(BasePlateOptionsDataRow))
'            End If
'        Next 'Add base plate options

'    End Sub 'Generate from EDS
'    Public Sub New(ByVal path As String)

'        Try
'            Me.connection_group_id = CType(GetOneExcelRange(path, "ID"), Integer)
'        Catch
'            Me.connection_group_id = 0
'        End Try 'Connection_Group_Id

'        For Each PlateConnectionsDataRow As DataRow In ds.Tables("CCIplate Plate Connections EXCEL").Rows
'            Me.plate_connections.Add(New CCIplatePlateConnections(PlateConnectionsDataRow))
'        Next 'Add plate connections

'        For Each PlateDetailsDataRow As DataRow In ds.Tables("CCIplate Plate Details EXCEL").Rows
'            Me.plate_details.Add(New CCIplatePlateDetails(PlateDetailsDataRow))
'        Next 'Add plate connections

'        'For Each BasePlateOptionsDataRow As DataRow In ds.Tables("CCIplate Base Plate Options EXCEL").Rows
'        Me.base_plate_options.Add(New CCIplateBasePlateOptions(path))
'        'Next 'Add plate connections

'    End Sub
'#End Region 'Generate from Excel

'End Class

'Partial Public Class CCIplatePlateConnections
'#Region "Define"
'    Private prop_connection_id As Integer
'    Private prop_connection_elevation As Double?
'    Private prop_connection_type As String
'    Private prop_bolt_configuration As String
'    Private prop_local_id As Integer?

'    <Category("Plate Connections"), Description(""), DisplayName("Connection_Id")>
'    Public Property connection_id() As Integer
'        Get
'            Return Me.prop_connection_id
'        End Get
'        Set
'            Me.prop_connection_id = Value
'        End Set
'    End Property
'    <Category("Plate Connections"), Description(""), DisplayName("Connection_Elevation")>
'    Public Property connection_elevation() As Double?
'        Get
'            Return Me.prop_connection_elevation
'        End Get
'        Set
'            Me.prop_connection_elevation = Value
'        End Set
'    End Property
'    <Category("Plate Connections"), Description(""), DisplayName("Connection_Type")>
'    Public Property connection_type() As String
'        Get
'            Return Me.prop_connection_type
'        End Get
'        Set
'            Me.prop_connection_type = Value
'        End Set
'    End Property
'    <Category("Plate Connections"), Description(""), DisplayName("Bolt_Configuration")>
'    Public Property bolt_configuration() As String
'        Get
'            Return Me.prop_bolt_configuration
'        End Get
'        Set
'            Me.prop_bolt_configuration = Value
'        End Set
'    End Property
'    <Category("Plate Connections"), Description(""), DisplayName("Local_Id")>
'    Public Property local_id() As Integer?
'        Get
'            Return Me.prop_local_id
'        End Get
'        Set
'            Me.prop_local_id = Value
'        End Set
'    End Property

'#End Region
'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PlateConnectionsDataRow As DataRow)
'        Try
'            Me.connection_id = CType(PlateConnectionsDataRow.Item("connection_id"), Integer)
'        Catch
'            Me.connection_id = 0
'        End Try 'Connection_Id
'        Try
'            If Not IsDBNull(CType(PlateConnectionsDataRow.Item("connection_elevation"), Double)) Then
'                Me.connection_elevation = CType(PlateConnectionsDataRow.Item("connection_elevation"), Double)
'            Else
'                Me.connection_elevation = Nothing
'            End If
'        Catch
'            Me.connection_elevation = Nothing
'        End Try 'Connection_Elevation
'        Try
'            Me.connection_type = CType(PlateConnectionsDataRow.Item("connection_type"), String)
'        Catch
'            Me.connection_type = ""
'        End Try 'Connection_Type
'        Try
'            Me.bolt_configuration = CType(PlateConnectionsDataRow.Item("bolt_configuration"), String)
'        Catch
'            Me.bolt_configuration = ""
'        End Try 'Bolt_Configuration
'        Try
'            If Not IsDBNull(CType(PlateConnectionsDataRow.Item("local_id"), Integer)) Then
'                Me.local_id = CType(PlateConnectionsDataRow.Item("local_id"), Integer)
'            Else
'                Me.local_id = Nothing
'            End If
'        Catch
'            Me.local_id = Nothing
'        End Try 'Local_Id

'    End Sub 'ENTER DESCRIPTOR
'#End Region

'End Class

'Partial Public Class CCIplatePlateDetails
'#Region "Define"
'    Private prop_plate_id As Integer
'    Private prop_plate_location As String
'    Private prop_plate_type As String
'    Private prop_plate_diameter As Double?
'    Private prop_plate_thickness As Double?
'    Private prop_plate_material As Integer?
'    Private prop_stiffener_configuration As Integer?
'    Private prop_stiffener_clear_space As Double?
'    Private prop_plate_check As Boolean
'    Private prop_local_id As Integer?
'    <Category("Plate Details"), Description(""), DisplayName("Plate_Id")>
'    Public Property plate_id() As Integer
'        Get
'            Return Me.prop_plate_id
'        End Get
'        Set
'            Me.prop_plate_id = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate_Location")>
'    Public Property plate_location() As String
'        Get
'            Return Me.prop_plate_location
'        End Get
'        Set
'            Me.prop_plate_location = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate_Type")>
'    Public Property plate_type() As String
'        Get
'            Return Me.prop_plate_type
'        End Get
'        Set
'            Me.prop_plate_type = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate_Diameter")>
'    Public Property plate_diameter() As Double?
'        Get
'            Return Me.prop_plate_diameter
'        End Get
'        Set
'            Me.prop_plate_diameter = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate_Thickness")>
'    Public Property plate_thickness() As Double?
'        Get
'            Return Me.prop_plate_thickness
'        End Get
'        Set
'            Me.prop_plate_thickness = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate_Material")>
'    Public Property plate_material() As Integer?
'        Get
'            Return Me.prop_plate_material
'        End Get
'        Set
'            Me.prop_plate_material = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Stiffener_Configuration")>
'    Public Property stiffener_configuration() As Integer?
'        Get
'            Return Me.prop_stiffener_configuration
'        End Get
'        Set
'            Me.prop_stiffener_configuration = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Stiffener_Clear_Space")>
'    Public Property stiffener_clear_space() As Double?
'        Get
'            Return Me.prop_stiffener_clear_space
'        End Get
'        Set
'            Me.prop_stiffener_clear_space = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Plate_Check")>
'    Public Property plate_check() As Boolean
'        Get
'            Return Me.prop_plate_check
'        End Get
'        Set
'            Me.prop_plate_check = Value
'        End Set
'    End Property
'    <Category("Plate Details"), Description(""), DisplayName("Local_Id")>
'    Public Property local_id() As Integer?
'        Get
'            Return Me.prop_local_id
'        End Get
'        Set
'            Me.prop_local_id = Value
'        End Set
'    End Property
'#End Region
'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PlateDetailsDataRow As DataRow)
'        Try
'            Me.plate_id = CType(PlateDetailsDataRow.Item("plate_id"), Integer)
'        Catch
'            Me.plate_id = 0
'        End Try 'Plate_Id
'        Try
'            Me.plate_location = CType(PlateDetailsDataRow.Item("plate_location"), String)
'        Catch
'            Me.plate_location = ""
'        End Try 'Plate_Location
'        Try
'            Me.plate_type = CType(PlateDetailsDataRow.Item("plate_type"), String)
'        Catch
'            Me.plate_type = ""
'        End Try 'Plate_Type
'        Try
'            If Not IsDBNull(CType(PlateDetailsDataRow.Item("plate_diameter"), Double)) Then
'                Me.plate_diameter = CType(PlateDetailsDataRow.Item("plate_diameter"), Double)
'            Else
'                Me.plate_diameter = Nothing
'            End If
'        Catch
'            Me.plate_diameter = Nothing
'        End Try 'Plate_Diameter
'        Try
'            If Not IsDBNull(CType(PlateDetailsDataRow.Item("plate_thickness"), Double)) Then
'                Me.plate_thickness = CType(PlateDetailsDataRow.Item("plate_thickness"), Double)
'            Else
'                Me.plate_thickness = Nothing
'            End If
'        Catch
'            Me.plate_thickness = Nothing
'        End Try 'Plate_Thickness
'        Try
'            If Not IsDBNull(CType(PlateDetailsDataRow.Item("plate_material"), Integer)) Then
'                Me.plate_material = CType(PlateDetailsDataRow.Item("plate_material"), Integer)
'            Else
'                Me.plate_material = Nothing
'            End If
'        Catch
'            Me.plate_material = Nothing
'        End Try 'Plate_Material
'        Try
'            If Not IsDBNull(CType(PlateDetailsDataRow.Item("stiffener_configuration"), Integer)) Then
'                Me.stiffener_configuration = CType(PlateDetailsDataRow.Item("stiffener_configuration"), Integer)
'            Else
'                Me.stiffener_configuration = Nothing
'            End If
'        Catch
'            Me.stiffener_configuration = Nothing
'        End Try 'Stiffener_Configuration
'        Try
'            If Not IsDBNull(CType(PlateDetailsDataRow.Item("stiffener_clear_space"), Double)) Then
'                Me.stiffener_clear_space = CType(PlateDetailsDataRow.Item("stiffener_clear_space"), Double)
'            Else
'                Me.stiffener_clear_space = Nothing
'            End If
'        Catch
'            Me.stiffener_clear_space = Nothing
'        End Try 'Stiffener_Clear_Space
'        Try
'            If CType(PlateDetailsDataRow.Item("plate_check"), String) = "Yes" Then
'                Me.plate_check = True
'            Else
'                Me.plate_check = False
'            End If
'            'Me.plate_check = CType(PlateDetailsDataRow.Item("plate_check"), Boolean)
'        Catch
'            Me.plate_check = True
'        End Try 'Plate_Check
'        Try
'            If Not IsDBNull(CType(PlateDetailsDataRow.Item("local_id"), Integer)) Then
'                Me.local_id = CType(PlateDetailsDataRow.Item("local_id"), Integer)
'            Else
'                Me.local_id = Nothing
'            End If
'        Catch
'            Me.local_id = Nothing
'        End Try 'Local_Id
'    End Sub 'ENTER DESCRIPTOR
'#End Region

'End Class

'Partial Public Class CCIplateBasePlateOptions
'#Region "Define"
'    Private prop_base_plate_option_id As Integer
'    Private prop_anchor_rod_spacing As Double?
'    Private prop_clip_distance As Double?
'    Private prop_barb_cl_elevation As Double?
'    Private prop_include_pole_reactions As Boolean
'    Private prop_consider_ar_eccentricity As Boolean
'    Private prop_leg_mod_eccentricity As Double?
'    <Category("Base Plate Options"), Description(""), DisplayName("Base_Plate_Option_Id")>
'    Public Property base_plate_option_id() As Integer
'        Get
'            Return Me.prop_base_plate_option_id
'        End Get
'        Set
'            Me.prop_base_plate_option_id = Value
'        End Set
'    End Property
'    <Category("Base Plate Options"), Description(""), DisplayName("Anchor_Rod_Spacing")>
'    Public Property anchor_rod_spacing() As Double?
'        Get
'            Return Me.prop_anchor_rod_spacing
'        End Get
'        Set
'            Me.prop_anchor_rod_spacing = Value
'        End Set
'    End Property
'    <Category("Base Plate Options"), Description(""), DisplayName("Clip_Distance")>
'    Public Property clip_distance() As Double?
'        Get
'            Return Me.prop_clip_distance
'        End Get
'        Set
'            Me.prop_clip_distance = Value
'        End Set
'    End Property
'    <Category("Base Plate Options"), Description(""), DisplayName("Barb_Cl_Elevation")>
'    Public Property barb_cl_elevation() As Double?
'        Get
'            Return Me.prop_barb_cl_elevation
'        End Get
'        Set
'            Me.prop_barb_cl_elevation = Value
'        End Set
'    End Property
'    <Category("Base Plate Options"), Description(""), DisplayName("Include_Pole_Reactions")>
'    Public Property include_pole_reactions() As Boolean
'        Get
'            Return Me.prop_include_pole_reactions
'        End Get
'        Set
'            Me.prop_include_pole_reactions = Value
'        End Set
'    End Property
'    <Category("Base Plate Options"), Description(""), DisplayName("Consider_Ar_Eccentricity")>
'    Public Property consider_ar_eccentricity() As Boolean
'        Get
'            Return Me.prop_consider_ar_eccentricity
'        End Get
'        Set
'            Me.prop_consider_ar_eccentricity = Value
'        End Set
'    End Property
'    <Category("Base Plate Options"), Description(""), DisplayName("Leg_Mod_Eccentricity")>
'    Public Property leg_mod_eccentricity() As Double?
'        Get
'            Return Me.prop_leg_mod_eccentricity
'        End Get
'        Set
'            Me.prop_leg_mod_eccentricity = Value
'        End Set
'    End Property
'#End Region
'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal BasePlateOptionsDataRow As DataRow)
'        Try
'            Me.base_plate_option_id = CType(BasePlateOptionsDataRow.Item("base_plate_option_id"), Integer)
'        Catch
'            Me.base_plate_option_id = 0
'        End Try 'Base_Plate_Option_Id
'        Try
'            If Not IsDBNull(CType(BasePlateOptionsDataRow.Item("anchor_rod_spacing"), Double)) Then
'                Me.anchor_rod_spacing = CType(BasePlateOptionsDataRow.Item("anchor_rod_spacing"), Double)
'            Else
'                Me.anchor_rod_spacing = Nothing
'            End If
'        Catch
'            Me.anchor_rod_spacing = Nothing
'        End Try 'Anchor_Rod_Spacing
'        Try
'            If Not IsDBNull(CType(BasePlateOptionsDataRow.Item("clip_distance"), Double)) Then
'                Me.clip_distance = CType(BasePlateOptionsDataRow.Item("clip_distance"), Double)
'            Else
'                Me.clip_distance = Nothing
'            End If
'        Catch
'            Me.clip_distance = Nothing
'        End Try 'Clip_Distance
'        Try
'            If Not IsDBNull(CType(BasePlateOptionsDataRow.Item("barb_cl_elevation"), Double)) Then
'                Me.barb_cl_elevation = CType(BasePlateOptionsDataRow.Item("barb_cl_elevation"), Double)
'            Else
'                Me.barb_cl_elevation = Nothing
'            End If
'        Catch
'            Me.barb_cl_elevation = Nothing
'        End Try 'Barb_Cl_Elevation
'        Try
'            Me.include_pole_reactions = CType(BasePlateOptionsDataRow.Item("include_pole_reactions"), Boolean)
'        Catch
'            Me.include_pole_reactions = False
'        End Try 'Include_Pole_Reactions
'        Try
'            Me.consider_ar_eccentricity = CType(BasePlateOptionsDataRow.Item("consider_ar_eccentricity"), Boolean)
'        Catch
'            Me.consider_ar_eccentricity = True
'        End Try 'Consider_Ar_Eccentricity
'        Try
'            If Not IsDBNull(CType(BasePlateOptionsDataRow.Item("leg_mod_eccentricity"), Double)) Then
'                Me.leg_mod_eccentricity = CType(BasePlateOptionsDataRow.Item("leg_mod_eccentricity"), Double)
'            Else
'                Me.leg_mod_eccentricity = Nothing
'            End If
'        Catch
'            Me.leg_mod_eccentricity = Nothing
'        End Try 'Leg_Mod_Eccentricity
'    End Sub 'Generate from EDS

'    Public Sub New(ByVal path As String)
'        Try
'            Me.base_plate_option_id = CType(GetOneExcelRange(path, "", ""), Integer)
'        Catch
'            Me.base_plate_option_id = 0
'        End Try 'Base_Plate_Option_Id
'        Try
'            If Not IsNothing(CType(GetOneExcelRange(path, "rod_spacing"), Double)) Then
'                Me.anchor_rod_spacing = CType(GetOneExcelRange(path, "rod_spacing"), Double)
'            Else
'                Me.anchor_rod_spacing = Nothing
'            End If
'        Catch
'            Me.anchor_rod_spacing = Nothing
'        End Try 'Anchor_Rod_Spacing
'        Try
'            If Not IsNothing(CType(GetOneExcelRange(path, "clip"), Double)) Then
'                Me.clip_distance = CType(GetOneExcelRange(path, "clip"), Double)
'            Else
'                Me.clip_distance = Nothing
'            End If
'        Catch
'            Me.clip_distance = Nothing
'        End Try 'Clip_Distance
'        Try
'            If Not IsNothing(CType(GetOneExcelRange(path, "H8", "Custom Connection"), Double)) Then
'                Me.barb_cl_elevation = CType(GetOneExcelRange(path, "H8", "Custom Connection"), Double)
'            Else
'                Me.barb_cl_elevation = Nothing
'            End If
'        Catch
'            Me.barb_cl_elevation = Nothing
'        End Try 'Barb_Cl_Elevation
'        Try
'            Me.include_pole_reactions = CType(GetOneExcelRange(path, "X2", "BARB"), Boolean)
'        Catch
'            Me.include_pole_reactions = False
'        End Try 'Include_Pole_Reactions
'        Try
'            Me.consider_ar_eccentricity = CType(GetOneExcelRange(path, "D17", "Engine"), Boolean)
'        Catch
'            Me.consider_ar_eccentricity = True
'        End Try 'Consider_Ar_Eccentricity
'        Try
'            If Not IsNothing(CType(GetOneExcelRange(path, "J8", "Custom Connection"), Double)) Then
'                Me.leg_mod_eccentricity = CType(GetOneExcelRange(path, "J8", "Custom Connection"), Double)
'            Else
'                Me.leg_mod_eccentricity = Nothing
'            End If
'        Catch
'            Me.leg_mod_eccentricity = Nothing
'        End Try 'Leg_Mod_Eccentricity
'    End Sub 'Generate from Excel
'#End Region

'End Class