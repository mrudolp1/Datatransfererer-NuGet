'Option Strict On

'Imports System.ComponentModel
'Imports System.Data
'Imports DevExpress.Spreadsheet

'Public Class CCIpole_old

'#Region "Define"
'    Private prop_pole_structure_id As Integer

'    Public Property criteria As New List(Of PoleCriteria)
'    Public Property unreinf_sections As New List(Of PoleSection)
'    Public Property reinf_sections As New List(Of PoleReinfSection)
'    Public Property reinf_groups As New List(Of PoleReinfGroup)
'    'Public Property reinf_ids As New List(Of PoleReinfDetail)
'    Public Property int_groups As New List(Of PoleIntGroup)
'    'Public Property int_ids As New List(Of PoleIntDetail)
'    Public Property reinf_section_results As New List(Of PoleReinfResults)
'    Public Property reinfs As New List(Of PropReinf)
'    Public Property bolts As New List(Of PropBolt)
'    Public Property matls As New List(Of PropMatl)


'    <Category("Pole Structure"), Description(""), DisplayName("Pole Structure ID")>
'    Public Property pole_structure_id() As Integer
'        Get
'            Return Me.prop_pole_structure_id
'        End Get
'        Set
'            Me.prop_pole_structure_id = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave method empty
'    End Sub

'    Public Sub New(ByVal PoleStructureDataRow As DataRow, refID As Integer)
'        Try
'            Me.pole_structure_id = refID
'        Catch
'            Me.pole_structure_id = 0
'        End Try 'Pole Structure ID


'        For Each PoleAnalysisCriteriaDataRow As DataRow In ds.Tables("CCIpole Criteria SQL").Rows
'            Dim PoleCriteriaRefID As Integer? = CType(PoleAnalysisCriteriaDataRow.Item("pole_structure_id"), Integer)
'            If PoleCriteriaRefID = refID Then
'                Me.criteria.Add(New PoleCriteria(PoleAnalysisCriteriaDataRow))
'            End If
'        Next 'Add Analysis Criteria to CCIpole Object

'        For Each PoleSectionDataRow As DataRow In ds.Tables("CCIpole Pole Sections SQL").Rows
'            Dim PoleSectionRefID As Integer? = CType(PoleSectionDataRow.Item("pole_structure_id"), Integer)
'            If PoleSectionRefID = refID Then
'                Me.unreinf_sections.Add(New PoleSection(PoleSectionDataRow))
'            End If
'        Next 'Add Unreinf Sections to CCIpole Object

'        For Each PoleReinfSectionDataRow As DataRow In ds.Tables("CCIpole Pole Reinf Sections SQL").Rows
'            Dim PoleReinfSectionRefID As Integer? = CType(PoleReinfSectionDataRow.Item("pole_structure_id"), Integer)
'            If PoleReinfSectionRefID = refID Then
'                Me.reinf_sections.Add(New PoleReinfSection(PoleReinfSectionDataRow))
'            End If
'        Next 'Add Reinf Sections to CCIpole Object

'        For Each PoleReinfGroupDataRow As DataRow In ds.Tables("CCIpole Reinf Groups SQL").Rows
'            Dim PoleReinfGroupRefID As Integer? = CType(PoleReinfGroupDataRow.Item("pole_structure_id"), Integer)
'            Dim PoleReinfGroupID As Integer? = CType(PoleReinfGroupDataRow.Item("group_id"), Integer)

'            If PoleReinfGroupRefID = refID Then
'                Dim NewReinfGroup As PoleReinfGroup
'                NewReinfGroup = New PoleReinfGroup(PoleReinfGroupDataRow)

'                For Each PoleReinfDetailDataRow As DataRow In ds.Tables("CCIpole Reinf Details SQL").Rows
'                    Dim PoleReinfDetailRefID As Integer? = CType(PoleReinfDetailDataRow.Item("group_id"), Integer)
'                    If PoleReinfDetailRefID = PoleReinfGroupID Then
'                        NewReinfGroup.reinf_ids.Add(New PoleReinfDetail(PoleReinfDetailDataRow))
'                    End If
'                Next 'Add Reinf Details to Group Object

'                Me.reinf_groups.Add(NewReinfGroup)
'            End If
'        Next 'Add Reinf Groups to CCIpole Object

'        For Each PoleIntGroupDataRow As DataRow In ds.Tables("CCIpole Int Groups SQL").Rows
'            Dim PoleIntGroupRefID As Integer? = CType(PoleIntGroupDataRow.Item("pole_structure_id"), Integer)
'            Dim PoleIntGroupID As Integer? = CType(PoleIntGroupDataRow.Item("group_id"), Integer)

'            If PoleIntGroupRefID = refID Then
'                Dim NewIntGroup As PoleIntGroup
'                NewIntGroup = New PoleIntGroup(PoleIntGroupDataRow)

'                For Each PoleIntDetailDataRow As DataRow In ds.Tables("CCIpole Int Details SQL").Rows
'                    Dim PoleIntDetailsRefID As Integer? = CType(PoleIntDetailDataRow.Item("group_id"), Integer)
'                    If PoleIntDetailsRefID = PoleIntGroupID Then
'                        NewIntGroup.int_ids.Add(New PoleIntDetail(PoleIntDetailDataRow))
'                    End If
'                Next 'Add Interference Details to Group Object

'                Me.int_groups.Add(NewIntGroup)
'            End If
'        Next 'Add Interference Groups to CCIpole Object

'        For Each PoleReinfResultsDataRow As DataRow In ds.Tables("CCIpole Pole Reinf Results SQL").Rows
'            Dim PoleReinfResultsRefID As Integer? = CType(PoleReinfResultsDataRow.Item("pole_structure_id"), Integer)
'            If PoleReinfResultsRefID = refID Then
'                Me.reinf_section_results.Add(New PoleReinfResults(PoleReinfResultsDataRow))
'            End If
'        Next 'Add Reinf Section Results to CCIpole Object

'        For Each PropReinfDataRow As DataRow In ds.Tables("CCIpole Reinf Property Details SQL").Rows
'            Dim PropReinfRefID As Integer? = CType(PropReinfDataRow.Item("pole_structure_id"), Integer)
'            If PropReinfRefID = refID Then
'                Me.reinfs.Add(New PropReinf(PropReinfDataRow))
'            End If
'        Next 'Add Custom Reinf Properties to CCIpole Object

'        For Each PropBoltDataRow As DataRow In ds.Tables("CCIpole Bolt Property Details SQL").Rows
'            Dim PropBoltRefID As Integer? = CType(PropBoltDataRow.Item("pole_structure_id"), Integer)
'            If PropBoltRefID = refID Then
'                Me.bolts.Add(New PropBolt(PropBoltDataRow))
'            End If
'        Next 'Add Custom Bolt Properties to CCIpole Object

'        For Each PropMatlDataRow As DataRow In ds.Tables("CCIpole Matl Property Details SQL").Rows
'            Dim PropMatlRefID As Integer? = CType(PropMatlDataRow.Item("pole_structure_id"), Integer)
'            If PropMatlRefID = refID Then
'                Me.matls.Add(New PropMatl(PropMatlDataRow))
'            End If
'        Next 'Add Custom Matl Properties to CCIpole Object

'    End Sub 'Generate a CCIpole object from EDS or Excel Datarow

'    Public Sub New(ByVal path As String)
'        Try
'            Me.pole_structure_id = CType(GetOneExcelRange(path, "ID_pole"), Integer)
'        Catch
'            Me.pole_structure_id = 0
'        End Try 'Pole Structure ID

'        For Each PoleAnalysisCriteriaDataRow As DataRow In ds.Tables("CCIPole Criteria EXCEL").Rows
'            Me.criteria.Add(New PoleCriteria(PoleAnalysisCriteriaDataRow))
'        Next 'Add Analysis Criteria to CCIpole Object

'        For Each PoleSectionDataRow As DataRow In ds.Tables("CCIpole Pole Sections EXCEL").Rows
'            Me.unreinf_sections.Add(New PoleSection(PoleSectionDataRow))
'        Next 'Add Unreinf Sections to CCIpole Object

'        For Each PoleReinfSectionDataRow As DataRow In ds.Tables("CCIpole Pole Reinf Sections EXCEL").Rows
'            Me.reinf_sections.Add(New PoleReinfSection(PoleReinfSectionDataRow))
'        Next 'Add Reinf Sections to CCIpole Object

'        For Each PoleReinfGroupDataRow As DataRow In ds.Tables("CCIpole Reinf Groups EXCEL").Rows
'            Dim NewReinfGroup As PoleReinfGroup
'            NewReinfGroup = New PoleReinfGroup(PoleReinfGroupDataRow)

'            Dim ReinfGroupID, DetailsGroupID As Integer?

'            Try
'                If Not IsDBNull(CType(PoleReinfGroupDataRow.Item("local_group_id"), Integer)) Then
'                    ReinfGroupID = CType(PoleReinfGroupDataRow.Item("local_group_id"), Integer)
'                Else
'                    ReinfGroupID = Nothing
'                End If
'            Catch
'                ReinfGroupID = Nothing
'            End Try 'Group local_group_id

'            For Each PoleReinfDetailDataRow As DataRow In ds.Tables("CCIpole Reinf Details EXCEL").Rows

'                Try
'                    If Not IsDBNull(CType(PoleReinfDetailDataRow.Item("local_group_id"), Integer)) Then
'                        DetailsGroupID = CType(PoleReinfDetailDataRow.Item("local_group_id"), Integer)
'                    Else
'                        DetailsGroupID = Nothing
'                    End If
'                Catch
'                    DetailsGroupID = Nothing
'                End Try 'Details local_group_id

'                If ReinfGroupID = DetailsGroupID Then
'                    NewReinfGroup.reinf_ids.Add(New PoleReinfDetail(PoleReinfDetailDataRow))
'                End If

'            Next 'Add Reinf Details to CCIpole Object

'            Me.reinf_groups.Add(NewReinfGroup)

'        Next 'Add Reinf Groups to CCIpole Object

'        'For Each PoleReinfDetailDataRow As DataRow In ds.Tables("CCIpole Reinf Details EXCEL").Rows
'        '    Me.reinf_ids.Add(New PoleReinfDetail(PoleReinfDetailDataRow))
'        'Next 'Add Reinf Details to CCIpole Object

'        For Each PoleIntGroupDataRow As DataRow In ds.Tables("CCIpole Int Groups EXCEL").Rows
'            Dim NewIntGroup As PoleIntGroup
'            NewIntGroup = New PoleIntGroup(PoleIntGroupDataRow)

'            Dim IntGroupID, IntDetailsGroupID As Integer?

'            Try
'                If Not IsDBNull(CType(PoleIntGroupDataRow.Item("local_group_id"), Integer)) Then
'                    IntGroupID = CType(PoleIntGroupDataRow.Item("local_group_id"), Integer)
'                Else
'                    IntGroupID = Nothing
'                End If
'            Catch
'                IntGroupID = Nothing
'            End Try 'Group local_group_id

'            For Each PoleIntDetailDataRow As DataRow In ds.Tables("CCIpole Int Details EXCEL").Rows

'                Try
'                    If Not IsDBNull(CType(PoleIntDetailDataRow.Item("local_group_id"), Integer)) Then
'                        IntDetailsGroupID = CType(PoleIntDetailDataRow.Item("local_group_id"), Integer)
'                    Else
'                        IntDetailsGroupID = Nothing
'                    End If
'                Catch
'                    IntDetailsGroupID = Nothing
'                End Try 'Details local_group_id

'                If IntGroupID = IntDetailsGroupID Then
'                    NewIntGroup.int_ids.Add(New PoleIntDetail(PoleIntDetailDataRow))
'                End If

'            Next

'            Me.int_groups.Add(NewIntGroup)

'        Next 'Add Interference Groups to CCIpole Object

'        'For Each PoleIntDetailDataRow As DataRow In ds.Tables("CCIpole Int Details EXCEL").Rows
'        '    Me.int_ids.Add(New PoleIntDetail(PoleIntDetailDataRow))
'        'Next 'Add Interference Details to CCIpole Object

'        For Each PoleReinfResultsDataRow As DataRow In ds.Tables("CCIpole Pole Reinf Results EXCEL").Rows
'            Me.reinf_section_results.Add(New PoleReinfResults(PoleReinfResultsDataRow))
'        Next 'Add Reinf Section Results to CCIpole Object

'        For Each PropReinfDataRow As DataRow In ds.Tables("CCIpole Reinf Property Details EXCEL").Rows
'            Me.reinfs.Add(New PropReinf(PropReinfDataRow))
'        Next 'Add Custom Reinf Properties to CCIpole Object

'        For Each PropBoltDataRow As DataRow In ds.Tables("CCIpole Bolt Property Details EXCEL").Rows
'            Me.bolts.Add(New PropBolt(PropBoltDataRow))
'        Next 'Add Custom Bolt Properties to CCIpole Object

'        For Each PropMatlDataRow As DataRow In ds.Tables("CCIpole Matl Property Details EXCEL").Rows
'            Me.matls.Add(New PropMatl(PropMatlDataRow))
'        Next 'Add Custom Matl Properties to CCIpole Object

'    End Sub 'Generate a CCIpole object from Excel

'#End Region

'End Class 'Add main Pole Structure ID for CCIpole Object


'Partial Public Class PoleCriteria_old

'#Region "Define"
'    Private prop_criteria_id As Integer
'    Private prop_upper_structure_type As String
'    Private prop_analysis_deg As Double?
'    Private prop_geom_increment_length As Double?
'    Private prop_vnum As String
'    Private prop_check_connections As Boolean
'    Private prop_hole_deformation As Boolean
'    Private prop_ineff_mod_check As Boolean
'    Private prop_modified As Boolean

'    <Category("Pole Analysis Criteria"), Description(""), DisplayName("Criteria ID")>
'    Public Property criteria_id() As Integer
'        Get
'            Return Me.prop_criteria_id
'        End Get
'        Set
'            Me.prop_criteria_id = Value
'        End Set
'    End Property
'    <Category("Pole Analysis Criteria"), Description(""), DisplayName("Upper Structure Type")>
'    Public Property upper_structure_type() As String
'        Get
'            Return Me.prop_upper_structure_type
'        End Get
'        Set
'            Me.prop_upper_structure_type = Value
'        End Set
'    End Property
'    <Category("Pole Analysis Criteria"), Description(""), DisplayName("Analysis Degrees")>
'    Public Property analysis_deg() As Double?
'        Get
'            Return Me.prop_analysis_deg
'        End Get
'        Set
'            Me.prop_analysis_deg = Value
'        End Set
'    End Property
'    <Category("Pole Analysis Criteria"), Description(""), DisplayName("Geometry Increment Length")>
'    Public Property geom_increment_length() As Double?
'        Get
'            Return Me.prop_geom_increment_length
'        End Get
'        Set
'            Me.prop_geom_increment_length = Value
'        End Set
'    End Property
'    <Category("Pole Analysis Criteria"), Description(""), DisplayName("Version Number")>
'    Public Property vnum() As String
'        Get
'            Return Me.prop_vnum
'        End Get
'        Set
'            Me.prop_vnum = Value
'        End Set
'    End Property
'    <Category("Pole Analysis Criteria"), Description(""), DisplayName("Check Connections")>
'    Public Property check_connections() As Boolean
'        Get
'            Return Me.prop_check_connections
'        End Get
'        Set
'            Me.prop_check_connections = Value
'        End Set
'    End Property
'    <Category("Pole Analysis Criteria"), Description(""), DisplayName("Allow Hole Deformation")>
'    Public Property hole_deformation() As Boolean
'        Get
'            Return Me.prop_hole_deformation
'        End Get
'        Set
'            Me.prop_hole_deformation = Value
'        End Set
'    End Property
'    <Category("Pole Analysis Criteria"), Description(""), DisplayName("Ineffective Mod Check")>
'    Public Property ineff_mod_check() As Boolean
'        Get
'            Return Me.prop_ineff_mod_check
'        End Get
'        Set
'            Me.prop_ineff_mod_check = Value
'        End Set
'    End Property
'    <Category("Pole Analysis Criteria"), Description(""), DisplayName("Modified")>
'    Public Property modified() As Boolean
'        Get
'            Return Me.prop_modified
'        End Get
'        Set
'            Me.prop_modified = Value
'        End Set
'    End Property
'#End Region

'#Region "Constructors"
'    Sub New()
'        'Leave method empty
'    End Sub

'    Sub New(ByVal PoleAnalysisCriteriaDataRow As DataRow)
'        Try
'            Me.criteria_id = CType(PoleAnalysisCriteriaDataRow.Item("criteria_id"), Integer)
'        Catch
'            Me.criteria_id = 0
'        End Try 'Criteria ID
'        Try
'            Me.upper_structure_type = CType(PoleAnalysisCriteriaDataRow.Item("upper_structure_type"), String)
'        Catch
'            Me.upper_structure_type = Nothing
'        End Try 'Upper Structure Type
'        Try
'            Me.analysis_deg = CType(PoleAnalysisCriteriaDataRow.Item("analysis_deg"), Double)
'        Catch
'            Me.analysis_deg = Nothing
'        End Try 'Analysis Degrees
'        Try
'            Me.geom_increment_length = CType(PoleAnalysisCriteriaDataRow.Item("geom_increment_length"), Double)
'        Catch
'            Me.geom_increment_length = Nothing
'        End Try 'Geometry Increment Length
'        Try
'            Me.vnum = CType(PoleAnalysisCriteriaDataRow.Item("vnum"), String)
'        Catch
'            Me.vnum = Nothing
'        End Try 'Version Number
'        Try
'            Me.check_connections = CType(PoleAnalysisCriteriaDataRow.Item("check_connections"), Boolean)
'        Catch
'            Me.check_connections = Nothing
'        End Try 'Check Connections
'        Try
'            Me.hole_deformation = CType(PoleAnalysisCriteriaDataRow.Item("hole_deformation"), Boolean)
'        Catch
'            Me.hole_deformation = Nothing
'        End Try 'Allow Hole Deformation
'        Try
'            Me.ineff_mod_check = CType(PoleAnalysisCriteriaDataRow.Item("ineff_mod_check"), Boolean)
'        Catch
'            Me.ineff_mod_check = Nothing
'        End Try 'Ineffective Mod Check
'        Try
'            Me.modified = CType(PoleAnalysisCriteriaDataRow.Item("modified"), Boolean)
'        Catch
'            Me.modified = Nothing
'        End Try 'Modified

'    End Sub
'#End Region

'End Class 'Add an Analysis Criteria


'Partial Public Class PoleSection_old

'#Region "Define"
'    Private prop_section_id As Integer?
'    Private prop_local_section_id As Integer?
'    Private prop_elev_bot As Double?
'    Private prop_elev_top As Double?
'    Private prop_length_section As Double?
'    Private prop_length_splice As Double?
'    Private prop_num_sides As Integer?
'    Private prop_diam_bot As Double?
'    Private prop_diam_top As Double?
'    Private prop_wall_thickness As Double?
'    Private prop_bend_radius As Double?
'    Private prop_steel_grade_id As Integer?
'    Private prop_local_matl_id As Integer?
'    Private prop_pole_type As String
'    Private prop_section_name As String
'    Private prop_socket_length As Double?
'    Private prop_weight_mult As Double?
'    Private prop_wp_mult As Double?
'    Private prop_af_factor As Double?
'    Private prop_ar_factor As Double?
'    Private prop_round_area_ratio As Double?
'    Private prop_flat_area_ratio As Double?

'    <Category("PoleSection"), Description(""), DisplayName("Global_Section_ID")>
'    Public Property section_id() As Integer?
'        Get
'            Return Me.prop_section_id
'        End Get
'        Set
'            Me.prop_section_id = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Local_Section_ID")>
'    Public Property local_section_id() As Integer?
'        Get
'            Return Me.prop_local_section_id
'        End Get
'        Set
'            Me.prop_local_section_id = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Elev_Bot")>
'    Public Property elev_bot() As Double?
'        Get
'            Return Me.prop_elev_bot
'        End Get
'        Set
'            Me.prop_elev_bot = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Elev_Top")>
'    Public Property elev_top() As Double?
'        Get
'            Return Me.prop_elev_top
'        End Get
'        Set
'            Me.prop_elev_top = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Length_Section")>
'    Public Property length_section() As Double?
'        Get
'            Return Me.prop_length_section
'        End Get
'        Set
'            Me.prop_length_section = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Length_Splice")>
'    Public Property length_splice() As Double?
'        Get
'            Return Me.prop_length_splice
'        End Get
'        Set
'            Me.prop_length_splice = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Num_Sides")>
'    Public Property num_sides() As Integer?
'        Get
'            Return Me.prop_num_sides
'        End Get
'        Set
'            Me.prop_num_sides = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Diam_Bot")>
'    Public Property diam_bot() As Double?
'        Get
'            Return Me.prop_diam_bot
'        End Get
'        Set
'            Me.prop_diam_bot = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Diam_Top")>
'    Public Property diam_top() As Double?
'        Get
'            Return Me.prop_diam_top
'        End Get
'        Set
'            Me.prop_diam_top = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Wall_Thickness")>
'    Public Property wall_thickness() As Double?
'        Get
'            Return Me.prop_wall_thickness
'        End Get
'        Set
'            Me.prop_wall_thickness = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Bend_Radius")>
'    Public Property bend_radius() As Double?
'        Get
'            Return Me.prop_bend_radius
'        End Get
'        Set
'            Me.prop_bend_radius = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Steel_Grade_ID")>
'    Public Property steel_grade_id() As Integer?
'        Get
'            Return Me.prop_steel_grade_id
'        End Get
'        Set
'            Me.prop_steel_grade_id = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Local_Steel_Grade_ID")>
'    Public Property local_matl_id() As Integer?
'        Get
'            Return Me.prop_local_matl_id
'        End Get
'        Set
'            Me.prop_local_matl_id = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Pole_Type")>
'    Public Property pole_type() As String
'        Get
'            Return Me.prop_pole_type
'        End Get
'        Set
'            Me.prop_pole_type = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Section_Name")>
'    Public Property section_name() As String
'        Get
'            Return Me.prop_section_name
'        End Get
'        Set
'            Me.prop_section_name = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Socket_Length")>
'    Public Property socket_length() As Double?
'        Get
'            Return Me.prop_socket_length
'        End Get
'        Set
'            Me.prop_socket_length = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Weight_Mult")>
'    Public Property weight_mult() As Double?
'        Get
'            Return Me.prop_weight_mult
'        End Get
'        Set
'            Me.prop_weight_mult = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Wp_Mult")>
'    Public Property wp_mult() As Double?
'        Get
'            Return Me.prop_wp_mult
'        End Get
'        Set
'            Me.prop_wp_mult = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Af_Factor")>
'    Public Property af_factor() As Double?
'        Get
'            Return Me.prop_af_factor
'        End Get
'        Set
'            Me.prop_af_factor = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Ar_Factor")>
'    Public Property ar_factor() As Double?
'        Get
'            Return Me.prop_ar_factor
'        End Get
'        Set
'            Me.prop_ar_factor = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Round_Area_Ratio")>
'    Public Property round_area_ratio() As Double?
'        Get
'            Return Me.prop_round_area_ratio
'        End Get
'        Set
'            Me.prop_round_area_ratio = Value
'        End Set
'    End Property
'    <Category("PoleSection"), Description(""), DisplayName("Flat_Area_Ratio")>
'    Public Property flat_area_ratio() As Double?
'        Get
'            Return Me.prop_flat_area_ratio
'        End Get
'        Set
'            Me.prop_flat_area_ratio = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Sub New()
'        'Leave method empty
'    End Sub

'    Public Sub New(ByVal PoleSectionDataRow As DataRow)
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("section_id"), Integer)) Then
'                Me.section_id = CType(PoleSectionDataRow.Item("section_id"), Integer)
'            Else
'                Me.section_id = 0
'            End If
'        Catch
'            Me.section_id = 0
'        End Try 'Section_Id
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("local_section_id"), Integer)) Then
'                Me.local_section_id = CType(PoleSectionDataRow.Item("local_section_id"), Integer)
'            Else
'                Me.local_section_id = Nothing
'            End If
'        Catch
'            Me.local_section_id = Nothing
'        End Try 'Analysis_Section_Id
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("elev_bot"), Double)) Then
'                Me.elev_bot = CType(PoleSectionDataRow.Item("elev_bot"), Double)
'            Else
'                Me.elev_bot = Nothing
'            End If
'        Catch
'            Me.elev_bot = Nothing
'        End Try 'Elev_Bot
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("elev_top"), Double)) Then
'                Me.elev_top = CType(PoleSectionDataRow.Item("elev_top"), Double)
'            Else
'                Me.elev_top = Nothing
'            End If
'        Catch
'            Me.elev_top = Nothing
'        End Try 'Elev_Top
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("length_section"), Double)) Then
'                Me.length_section = CType(PoleSectionDataRow.Item("length_section"), Double)
'            Else
'                Me.length_section = Nothing
'            End If
'        Catch
'            Me.length_section = Nothing
'        End Try 'Length_Section
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("length_splice"), Double)) Then
'                Me.length_splice = CType(PoleSectionDataRow.Item("length_splice"), Double)
'            Else
'                Me.length_splice = Nothing
'            End If
'        Catch
'            Me.length_splice = Nothing
'        End Try 'Length_Splice
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("num_sides"), Integer)) Then
'                Me.num_sides = CType(PoleSectionDataRow.Item("num_sides"), Integer)
'            Else
'                Me.num_sides = Nothing
'            End If
'        Catch
'            Me.num_sides = Nothing
'        End Try 'Num_Sides
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("diam_bot"), Double)) Then
'                Me.diam_bot = CType(PoleSectionDataRow.Item("diam_bot"), Double)
'            Else
'                Me.diam_bot = Nothing
'            End If
'        Catch
'            Me.diam_bot = Nothing
'        End Try 'Diam_Bot
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("diam_top"), Double)) Then
'                Me.diam_top = CType(PoleSectionDataRow.Item("diam_top"), Double)
'            Else
'                Me.diam_top = Nothing
'            End If
'        Catch
'            Me.diam_top = Nothing
'        End Try 'Diam_Top
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("wall_thickness"), Double)) Then
'                Me.wall_thickness = CType(PoleSectionDataRow.Item("wall_thickness"), Double)
'            Else
'                Me.wall_thickness = Nothing
'            End If
'        Catch
'            Me.wall_thickness = Nothing
'        End Try 'Wall_Thickness
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("bend_radius"), Double)) Then
'                Me.bend_radius = CType(PoleSectionDataRow.Item("bend_radius"), Double)
'            Else
'                Me.bend_radius = Nothing
'            End If
'        Catch
'            Me.bend_radius = Nothing
'        End Try 'Bend_Radius
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("steel_grade_id"), Integer)) Then
'                Me.steel_grade_id = CType(PoleSectionDataRow.Item("steel_grade_id"), Integer)
'            Else
'                Me.steel_grade_id = 0
'            End If
'        Catch
'            Me.steel_grade_id = 0
'        End Try 'Steel_Grade_ID
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("local_matl_id"), Integer)) Then
'                Me.local_matl_id = CType(PoleSectionDataRow.Item("local_matl_id"), Integer)
'            Else
'                Me.local_matl_id = Nothing
'            End If
'        Catch
'            Me.local_matl_id = Nothing
'        End Try 'Local_Steel_Grade_ID
'        Try
'            Me.pole_type = CType(PoleSectionDataRow.Item("pole_type"), String)
'        Catch
'            Me.pole_type = ""
'        End Try 'Pole_Type
'        Try
'            Me.section_name = CType(PoleSectionDataRow.Item("section_name"), String)
'        Catch
'            Me.section_name = ""
'        End Try 'Section_Name
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("socket_length"), Double)) Then
'                Me.socket_length = CType(PoleSectionDataRow.Item("socket_length"), Double)
'            Else
'                Me.socket_length = Nothing
'            End If
'        Catch
'            Me.socket_length = Nothing
'        End Try 'Socket_Length
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("weight_mult"), Double)) Then
'                Me.weight_mult = CType(PoleSectionDataRow.Item("weight_mult"), Double)
'            Else
'                Me.weight_mult = Nothing
'            End If
'        Catch
'            Me.weight_mult = Nothing
'        End Try 'Weight_Mult
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("wp_mult"), Double)) Then
'                Me.wp_mult = CType(PoleSectionDataRow.Item("wp_mult"), Double)
'            Else
'                Me.wp_mult = Nothing
'            End If
'        Catch
'            Me.wp_mult = Nothing
'        End Try 'Wp_Mult
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("af_factor"), Double)) Then
'                Me.af_factor = CType(PoleSectionDataRow.Item("af_factor"), Double)
'            Else
'                Me.af_factor = Nothing
'            End If
'        Catch
'            Me.af_factor = Nothing
'        End Try 'Af_Factor
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("ar_factor"), Double)) Then
'                Me.ar_factor = CType(PoleSectionDataRow.Item("ar_factor"), Double)
'            Else
'                Me.ar_factor = Nothing
'            End If
'        Catch
'            Me.ar_factor = Nothing
'        End Try 'Ar_Factor
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("round_area_ratio"), Double)) Then
'                Me.round_area_ratio = CType(PoleSectionDataRow.Item("round_area_ratio"), Double)
'            Else
'                Me.round_area_ratio = Nothing
'            End If
'        Catch
'            Me.round_area_ratio = Nothing
'        End Try 'Round_Area_Ratio
'        Try
'            If Not IsDBNull(CType(PoleSectionDataRow.Item("flat_area_ratio"), Double)) Then
'                Me.flat_area_ratio = CType(PoleSectionDataRow.Item("flat_area_ratio"), Double)
'            Else
'                Me.flat_area_ratio = Nothing
'            End If
'        Catch
'            Me.flat_area_ratio = Nothing
'        End Try 'Flat_Area_Ratio
'    End Sub

'#End Region

'End Class 'Add Unreinf Geom to CCIpole Object

'Partial Public Class PoleReinfSection_old

'#Region "Define"
'    Private prop_section_id As Integer
'    Private prop_local_section_id As Integer?
'    Private prop_elev_bot As Double?
'    Private prop_elev_top As Double?
'    Private prop_length_section As Double?
'    Private prop_length_splice As Double?
'    Private prop_num_sides As Integer?
'    Private prop_diam_bot As Double?
'    Private prop_diam_top As Double?
'    Private prop_wall_thickness As Double?
'    Private prop_bend_radius As Double?
'    Private prop_steel_grade_id As Integer?
'    Private prop_local_matl_id As Integer?
'    Private prop_pole_type As String
'    Private prop_weight_mult As Double?
'    Private prop_section_name As String
'    Private prop_socket_length As Double?
'    Private prop_wp_mult As Double?
'    Private prop_af_factor As Double?
'    Private prop_ar_factor As Double?
'    Private prop_round_area_ratio As Double?
'    Private prop_flat_area_ratio As Double?

'    <Category("PoleReinfSection"), Description(""), DisplayName("Section_ID")>
'    Public Property section_id() As Integer
'        Get
'            Return Me.prop_section_id
'        End Get
'        Set
'            Me.prop_section_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Local_Section_ID")>
'    Public Property local_section_id() As Integer?
'        Get
'            Return Me.prop_local_section_id
'        End Get
'        Set
'            Me.prop_local_section_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Elev_Bot")>
'    Public Property elev_bot() As Double?
'        Get
'            Return Me.prop_elev_bot
'        End Get
'        Set
'            Me.prop_elev_bot = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Elev_Top")>
'    Public Property elev_top() As Double?
'        Get
'            Return Me.prop_elev_top
'        End Get
'        Set
'            Me.prop_elev_top = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Length_Section")>
'    Public Property length_section() As Double?
'        Get
'            Return Me.prop_length_section
'        End Get
'        Set
'            Me.prop_length_section = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Length_Splice")>
'    Public Property length_splice() As Double?
'        Get
'            Return Me.prop_length_splice
'        End Get
'        Set
'            Me.prop_length_splice = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Num_Sides")>
'    Public Property num_sides() As Integer?
'        Get
'            Return Me.prop_num_sides
'        End Get
'        Set
'            Me.prop_num_sides = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Diam_Bot")>
'    Public Property diam_bot() As Double?
'        Get
'            Return Me.prop_diam_bot
'        End Get
'        Set
'            Me.prop_diam_bot = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Diam_Top")>
'    Public Property diam_top() As Double?
'        Get
'            Return Me.prop_diam_top
'        End Get
'        Set
'            Me.prop_diam_top = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Wall_Thickness")>
'    Public Property wall_thickness() As Double?
'        Get
'            Return Me.prop_wall_thickness
'        End Get
'        Set
'            Me.prop_wall_thickness = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Bend_Radius")>
'    Public Property bend_radius() As Double?
'        Get
'            Return Me.prop_bend_radius
'        End Get
'        Set
'            Me.prop_bend_radius = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Steel_Grade_Id")>
'    Public Property steel_grade_id() As Integer?
'        Get
'            Return Me.prop_steel_grade_id
'        End Get
'        Set
'            Me.prop_steel_grade_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Local_Steel_Grade_Id")>
'    Public Property local_matl_id() As Integer?
'        Get
'            Return Me.prop_local_matl_id
'        End Get
'        Set
'            Me.prop_local_matl_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Pole_Type")>
'    Public Property pole_type() As String
'        Get
'            Return Me.prop_pole_type
'        End Get
'        Set
'            Me.prop_pole_type = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Weight_Mult")>
'    Public Property weight_mult() As Double?
'        Get
'            Return Me.prop_weight_mult
'        End Get
'        Set
'            Me.prop_weight_mult = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Section_Name")>
'    Public Property section_name() As String
'        Get
'            Return Me.prop_section_name
'        End Get
'        Set
'            Me.prop_section_name = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Socket_Length")>
'    Public Property socket_length() As Double?
'        Get
'            Return Me.prop_socket_length
'        End Get
'        Set
'            Me.prop_socket_length = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Wp_Mult")>
'    Public Property wp_mult() As Double?
'        Get
'            Return Me.prop_wp_mult
'        End Get
'        Set
'            Me.prop_wp_mult = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Af_Factor")>
'    Public Property af_factor() As Double?
'        Get
'            Return Me.prop_af_factor
'        End Get
'        Set
'            Me.prop_af_factor = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Ar_Factor")>
'    Public Property ar_factor() As Double?
'        Get
'            Return Me.prop_ar_factor
'        End Get
'        Set
'            Me.prop_ar_factor = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Round_Area_Ratio")>
'    Public Property round_area_ratio() As Double?
'        Get
'            Return Me.prop_round_area_ratio
'        End Get
'        Set
'            Me.prop_round_area_ratio = Value
'        End Set
'    End Property
'    <Category("PoleReinfSection"), Description(""), DisplayName("Flat_Area_Ratio")>
'    Public Property flat_area_ratio() As Double?
'        Get
'            Return Me.prop_flat_area_ratio
'        End Get
'        Set
'            Me.prop_flat_area_ratio = Value
'        End Set
'    End Property
'#End Region

'#Region "Constructors"

'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PoleReinfSectionDataRow As DataRow)
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("section_ID"), Integer)) Then
'                Me.section_id = CType(PoleReinfSectionDataRow.Item("section_ID"), Integer)
'            Else
'                Me.section_id = 0
'            End If
'        Catch
'            Me.section_id = 0
'        End Try 'Section_Id
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("local_section_id"), Integer)) Then
'                Me.local_section_id = CType(PoleReinfSectionDataRow.Item("local_section_id"), Integer)
'            Else
'                Me.local_section_id = Nothing
'            End If
'        Catch
'            Me.local_section_id = Nothing
'        End Try 'Analysis_Section_Id
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("elev_bot"), Double)) Then
'                Me.elev_bot = CType(PoleReinfSectionDataRow.Item("elev_bot"), Double)
'            Else
'                Me.elev_bot = Nothing
'            End If
'        Catch
'            Me.elev_bot = Nothing
'        End Try 'Elev_Bot
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("elev_top"), Double)) Then
'                Me.elev_top = CType(PoleReinfSectionDataRow.Item("elev_top"), Double)
'            Else
'                Me.elev_top = Nothing
'            End If
'        Catch
'            Me.elev_top = Nothing
'        End Try 'Elev_Top
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("length_section"), Double)) Then
'                Me.length_section = CType(PoleReinfSectionDataRow.Item("length_section"), Double)
'            Else
'                Me.length_section = Nothing
'            End If
'        Catch
'            Me.length_section = Nothing
'        End Try 'Length_Section
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("length_splice"), Double)) Then
'                Me.length_splice = CType(PoleReinfSectionDataRow.Item("length_splice"), Double)
'            Else
'                Me.length_splice = Nothing
'            End If
'        Catch
'            Me.length_splice = Nothing
'        End Try 'Length_Splice
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("num_sides"), Integer)) Then
'                Me.num_sides = CType(PoleReinfSectionDataRow.Item("num_sides"), Integer)
'            Else
'                Me.num_sides = Nothing
'            End If
'        Catch
'            Me.num_sides = Nothing
'        End Try 'Num_Sides
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("diam_bot"), Double)) Then
'                Me.diam_bot = CType(PoleReinfSectionDataRow.Item("diam_bot"), Double)
'            Else
'                Me.diam_bot = Nothing
'            End If
'        Catch
'            Me.diam_bot = Nothing
'        End Try 'Diam_Bot
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("diam_top"), Double)) Then
'                Me.diam_top = CType(PoleReinfSectionDataRow.Item("diam_top"), Double)
'            Else
'                Me.diam_top = Nothing
'            End If
'        Catch
'            Me.diam_top = Nothing
'        End Try 'Diam_Top
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("wall_thickness"), Double)) Then
'                Me.wall_thickness = CType(PoleReinfSectionDataRow.Item("wall_thickness"), Double)
'            Else
'                Me.wall_thickness = Nothing
'            End If
'        Catch
'            Me.wall_thickness = Nothing
'        End Try 'Wall_Thickness
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("bend_radius"), Double)) Then
'                Me.bend_radius = CType(PoleReinfSectionDataRow.Item("bend_radius"), Double)
'            Else
'                Me.bend_radius = Nothing
'            End If
'        Catch
'            Me.bend_radius = Nothing
'        End Try 'Bend_Radius
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("steel_grade_id"), Integer)) Then
'                Me.steel_grade_id = CType(PoleReinfSectionDataRow.Item("steel_grade_id"), Integer)
'            Else
'                Me.steel_grade_id = 0
'            End If
'        Catch
'            Me.steel_grade_id = 0
'        End Try 'Steel_Grade_Id
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("local_matl_id"), Integer)) Then
'                Me.local_matl_id = CType(PoleReinfSectionDataRow.Item("local_matl_id"), Integer)
'            Else
'                Me.local_matl_id = Nothing
'            End If
'        Catch
'            Me.local_matl_id = Nothing
'        End Try 'local_matl_id
'        Try
'            Me.pole_type = CType(PoleReinfSectionDataRow.Item("pole_type"), String)
'        Catch
'            Me.pole_type = ""
'        End Try 'Pole_Type
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("weight_mult"), Double)) Then
'                Me.weight_mult = CType(PoleReinfSectionDataRow.Item("weight_mult"), Double)
'            Else
'                Me.weight_mult = Nothing
'            End If
'        Catch
'            Me.weight_mult = Nothing
'        End Try 'Weight_Mult
'        Try
'            Me.section_name = CType(PoleReinfSectionDataRow.Item("section_name"), String)
'        Catch
'            Me.section_name = ""
'        End Try 'Section_Name
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("socket_length"), Double)) Then
'                Me.socket_length = CType(PoleReinfSectionDataRow.Item("socket_length"), Double)
'            Else
'                Me.socket_length = Nothing
'            End If
'        Catch
'            Me.socket_length = Nothing
'        End Try 'Socket_Length
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("wp_mult"), Double)) Then
'                Me.wp_mult = CType(PoleReinfSectionDataRow.Item("wp_mult"), Double)
'            Else
'                Me.wp_mult = Nothing
'            End If
'        Catch
'            Me.wp_mult = Nothing
'        End Try 'Wp_Mult
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("af_factor"), Double)) Then
'                Me.af_factor = CType(PoleReinfSectionDataRow.Item("af_factor"), Double)
'            Else
'                Me.af_factor = Nothing
'            End If
'        Catch
'            Me.af_factor = Nothing
'        End Try 'Af_Factor
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("ar_factor"), Double)) Then
'                Me.ar_factor = CType(PoleReinfSectionDataRow.Item("ar_factor"), Double)
'            Else
'                Me.ar_factor = Nothing
'            End If
'        Catch
'            Me.ar_factor = Nothing
'        End Try 'Ar_Factor
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("round_area_ratio"), Double)) Then
'                Me.round_area_ratio = CType(PoleReinfSectionDataRow.Item("round_area_ratio"), Double)
'            Else
'                Me.round_area_ratio = Nothing
'            End If
'        Catch
'            Me.round_area_ratio = Nothing
'        End Try 'Round_Area_Ratio
'        Try
'            If Not IsDBNull(CType(PoleReinfSectionDataRow.Item("flat_area_ratio"), Double)) Then
'                Me.flat_area_ratio = CType(PoleReinfSectionDataRow.Item("flat_area_ratio"), Double)
'            Else
'                Me.flat_area_ratio = Nothing
'            End If
'        Catch
'            Me.flat_area_ratio = Nothing
'        End Try 'Flat_Area_Ratio
'    End Sub 'Generate from EDS
'#End Region

'End Class 'Add Reinforced Geom to CCIpole Object

'Partial Public Class PoleReinfGroup_old

'#Region "Define"
'    Private prop_group_id As Integer?
'    Private prop_local_group_id As Integer?
'    Private prop_elev_bot_actual As Double?
'    Private prop_elev_bot_eff As Double?
'    Private prop_elev_top_actual As Double?
'    Private prop_elev_top_eff As Double?
'    Private prop_reinf_db_id As Integer?
'    Private prop_local_reinf_id As Integer?
'    Private prop_qty As Integer?
'    Public Property reinf_ids As New List(Of PoleReinfDetail)

'    <Category("PoleReinfGroup"), Description(""), DisplayName("Reinf_Group_Id")>
'    Public Property group_id() As Integer?
'        Get
'            Return Me.prop_group_id
'        End Get
'        Set
'            Me.prop_group_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfGroup"), Description(""), DisplayName("Local_Reinf_Group_Id")>
'    Public Property local_group_id() As Integer?
'        Get
'            Return Me.prop_local_group_id
'        End Get
'        Set
'            Me.prop_local_group_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfGroup"), Description(""), DisplayName("Elev_Bot_Actual")>
'    Public Property elev_bot_actual() As Double?
'        Get
'            Return Me.prop_elev_bot_actual
'        End Get
'        Set
'            Me.prop_elev_bot_actual = Value
'        End Set
'    End Property
'    <Category("PoleReinfGroup"), Description(""), DisplayName("Elev_Bot_Eff")>
'    Public Property elev_bot_eff() As Double?
'        Get
'            Return Me.prop_elev_bot_eff
'        End Get
'        Set
'            Me.prop_elev_bot_eff = Value
'        End Set
'    End Property
'    <Category("PoleReinfGroup"), Description(""), DisplayName("Elev_Top_Actual")>
'    Public Property elev_top_actual() As Double?
'        Get
'            Return Me.prop_elev_top_actual
'        End Get
'        Set
'            Me.prop_elev_top_actual = Value
'        End Set
'    End Property
'    <Category("PoleReinfGroup"), Description(""), DisplayName("Elev_Top_Eff")>
'    Public Property elev_top_eff() As Double?
'        Get
'            Return Me.prop_elev_top_eff
'        End Get
'        Set
'            Me.prop_elev_top_eff = Value
'        End Set
'    End Property
'    <Category("PoleReinfGroup"), Description(""), DisplayName("Reinf_Db_Id")>
'    Public Property reinf_db_id() As Integer?
'        Get
'            Return Me.prop_reinf_db_id
'        End Get
'        Set
'            Me.prop_reinf_db_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfGroup"), Description(""), DisplayName("Local_Reinf_Id")>
'    Public Property local_reinf_id() As Integer?
'        Get
'            Return Me.prop_local_reinf_id
'        End Get
'        Set
'            Me.prop_local_reinf_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfGroup"), Description(""), DisplayName("QTY")>
'    Public Property qty() As Integer?
'        Get
'            Return Me.prop_qty
'        End Get
'        Set
'            Me.prop_qty = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"

'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PoleReinfGroupDataRow As DataRow)
'        Try
'            If Not IsDBNull(CType(PoleReinfGroupDataRow.Item("group_id"), Integer)) Then
'                Me.group_id = CType(PoleReinfGroupDataRow.Item("group_id"), Integer)
'            Else
'                Me.group_id = 0
'            End If
'        Catch
'            Me.group_id = 0
'        End Try 'group_id
'        Try
'            If Not IsDBNull(CType(PoleReinfGroupDataRow.Item("local_group_id"), Integer)) Then
'                Me.local_group_id = CType(PoleReinfGroupDataRow.Item("local_group_id"), Integer)
'            Else
'                Me.local_group_id = Nothing
'            End If
'        Catch
'            Me.local_group_id = Nothing
'        End Try 'Local_Reinf_Group_Id
'        Try
'            If Not IsDBNull(CType(PoleReinfGroupDataRow.Item("elev_bot_actual"), Double)) Then
'                Me.elev_bot_actual = CType(PoleReinfGroupDataRow.Item("elev_bot_actual"), Double)
'            Else
'                Me.elev_bot_actual = Nothing
'            End If
'        Catch
'            Me.elev_bot_actual = Nothing
'        End Try 'Elev_Bot_Actual
'        Try
'            If Not IsDBNull(CType(PoleReinfGroupDataRow.Item("elev_bot_eff"), Double)) Then
'                Me.elev_bot_eff = CType(PoleReinfGroupDataRow.Item("elev_bot_eff"), Double)
'            Else
'                Me.elev_bot_eff = Nothing
'            End If
'        Catch
'            Me.elev_bot_eff = Nothing
'        End Try 'Elev_Bot_Eff
'        Try
'            If Not IsDBNull(CType(PoleReinfGroupDataRow.Item("elev_top_actual"), Double)) Then
'                Me.elev_top_actual = CType(PoleReinfGroupDataRow.Item("elev_top_actual"), Double)
'            Else
'                Me.elev_top_actual = Nothing
'            End If
'        Catch
'            Me.elev_top_actual = Nothing
'        End Try 'Elev_Top_Actual
'        Try
'            If Not IsDBNull(CType(PoleReinfGroupDataRow.Item("elev_top_eff"), Double)) Then
'                Me.elev_top_eff = CType(PoleReinfGroupDataRow.Item("elev_top_eff"), Double)
'            Else
'                Me.elev_top_eff = Nothing
'            End If
'        Catch
'            Me.elev_top_eff = Nothing
'        End Try 'Elev_Top_Eff
'        Try
'            If Not IsDBNull(CType(PoleReinfGroupDataRow.Item("reinf_db_id"), Integer)) Then
'                Me.reinf_db_id = CType(PoleReinfGroupDataRow.Item("reinf_db_id"), Integer)
'            Else
'                Me.reinf_db_id = 0
'            End If
'        Catch
'            Me.reinf_db_id = 0
'        End Try 'Reinf_Db_Id
'        Try
'            If Not IsDBNull(CType(PoleReinfGroupDataRow.Item("local_reinf_id"), Integer)) Then
'                Me.local_reinf_id = CType(PoleReinfGroupDataRow.Item("local_reinf_id"), Integer)
'            Else
'                Me.local_reinf_id = Nothing
'            End If
'        Catch
'            Me.local_reinf_id = Nothing
'        End Try 'Local_Reinf_Id
'        Try
'            If Not IsDBNull(CType(PoleReinfGroupDataRow.Item("qty"), Integer)) Then
'                Me.qty = CType(PoleReinfGroupDataRow.Item("qty"), Integer)
'            Else
'                Me.qty = Nothing
'            End If
'        Catch
'            Me.qty = Nothing
'        End Try 'QTY


'        'For Each PoleReinfDetailDataRow As DataRow In ds.Tables("CCIpole Reinf Details SQL").Rows
'        '    Dim PoleReinfDetailRefID As Integer = CType(PoleReinfDetailDataRow.Item("pole_structure_id"), Integer)
'        '    If PoleReinfDetailRefID = refID Then
'        '        Me.reinf_ids.Add(New PoleReinfDetail(PoleReinfDetailDataRow))
'        '    End If
'        'Next 'Add Reinf Details to CCIpole Object

'    End Sub

'#End Region

'End Class 'Add Reinf Group to CCIpole Object

'Partial Public Class PoleReinfDetail_old

'#Region "Define"
'    Private prop_reinf_id As Integer?
'    Private prop_local_group_id As Integer?
'    Private prop_local_reinf_id As Integer?
'    Private prop_pole_flat As Integer?
'    Private prop_horizontal_offset As Double?
'    Private prop_rotation As Double?
'    Private prop_note As String

'    <Category("PoleReinfDetails"), Description(""), DisplayName("Reinf_Id")>
'    Public Property reinf_id() As Integer?
'        Get
'            Return Me.prop_reinf_id
'        End Get
'        Set
'            Me.prop_reinf_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfDetails"), Description(""), DisplayName("Local_Group_ID")>
'    Public Property local_group_id() As Integer?
'        Get
'            Return Me.prop_local_group_id
'        End Get
'        Set
'            Me.prop_local_group_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfDetails"), Description(""), DisplayName("Local_Reinf_ID")>
'    Public Property local_reinf_id() As Integer?
'        Get
'            Return Me.prop_local_reinf_id
'        End Get
'        Set
'            Me.prop_local_reinf_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfDetails"), Description(""), DisplayName("Pole_Flat")>
'    Public Property pole_flat() As Integer?
'        Get
'            Return Me.prop_pole_flat
'        End Get
'        Set
'            Me.prop_pole_flat = Value
'        End Set
'    End Property
'    <Category("PoleReinfDetails"), Description(""), DisplayName("Horizontal_Offset")>
'    Public Property horizontal_offset() As Double?
'        Get
'            Return Me.prop_horizontal_offset
'        End Get
'        Set
'            Me.prop_horizontal_offset = Value
'        End Set
'    End Property
'    <Category("PoleReinfDetails"), Description(""), DisplayName("Rotation")>
'    Public Property rotation() As Double?
'        Get
'            Return Me.prop_rotation
'        End Get
'        Set
'            Me.prop_rotation = Value
'        End Set
'    End Property
'    <Category("PoleReinfDetails"), Description(""), DisplayName("Note")>
'    Public Property note() As String
'        Get
'            Return Me.prop_note
'        End Get
'        Set
'            Me.prop_note = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PoleReinfDetailDataRow As DataRow)
'        Try
'            If Not IsDBNull(CType(PoleReinfDetailDataRow.Item("reinf_id"), Integer)) Then
'                Me.reinf_id = CType(PoleReinfDetailDataRow.Item("reinf_id"), Integer)
'            Else
'                Me.reinf_id = 0
'            End If
'        Catch
'            Me.reinf_id = 0
'        End Try 'Reinf_Id
'        Try
'            If Not IsDBNull(CType(PoleReinfDetailDataRow.Item("local_group_id"), Integer)) Then
'                Me.local_group_id = CType(PoleReinfDetailDataRow.Item("local_group_id"), Integer)
'            Else
'                Me.local_group_id = Nothing
'            End If
'        Catch
'            Me.local_group_id = Nothing
'        End Try 'local_group_id
'        Try
'            If Not IsDBNull(CType(PoleReinfDetailDataRow.Item("local_reinf_id"), Integer)) Then
'                Me.local_reinf_id = CType(PoleReinfDetailDataRow.Item("local_reinf_id"), Integer)
'            Else
'                Me.local_reinf_id = Nothing
'            End If
'        Catch
'            Me.local_reinf_id = Nothing
'        End Try 'local_reinf_id
'        Try
'            If Not IsDBNull(CType(PoleReinfDetailDataRow.Item("pole_flat"), Integer)) Then
'                Me.pole_flat = CType(PoleReinfDetailDataRow.Item("pole_flat"), Integer)
'            Else
'                Me.pole_flat = Nothing
'            End If
'        Catch
'            Me.pole_flat = Nothing
'        End Try 'Pole_Flat
'        Try
'            If Not IsDBNull(CType(PoleReinfDetailDataRow.Item("horizontal_offset"), Double)) Then
'                Me.horizontal_offset = CType(PoleReinfDetailDataRow.Item("horizontal_offset"), Double)
'            Else
'                Me.horizontal_offset = Nothing
'            End If
'        Catch
'            Me.horizontal_offset = Nothing
'        End Try 'Horizontal_Offset
'        Try
'            If Not IsDBNull(CType(PoleReinfDetailDataRow.Item("rotation"), Double)) Then
'                Me.rotation = CType(PoleReinfDetailDataRow.Item("rotation"), Double)
'            Else
'                Me.rotation = Nothing
'            End If
'        Catch
'            Me.rotation = Nothing
'        End Try 'Rotation
'        Try
'            Me.note = CType(PoleReinfDetailDataRow.Item("note"), String)
'        Catch
'            Me.note = ""
'        End Try 'Note
'    End Sub

'#End Region

'End Class 'Add Reinf Detail to CCIpole Object

'Partial Public Class PoleIntGroup_old

'#Region "Define"
'    Private prop_group_id As Integer?
'    Private prop_local_group_id As Integer?
'    Private prop_elev_bot As Double?
'    Private prop_elev_top As Double?
'    Private prop_width As Double?
'    Private prop_description As String
'    Private prop_qty As Integer?
'    Public Property int_ids As New List(Of PoleIntDetail)

'    <Category("PoleIntGroup"), Description(""), DisplayName("Interference_Group_Id")>
'    Public Property group_id() As Integer?
'        Get
'            Return Me.prop_group_id
'        End Get
'        Set
'            Me.prop_group_id = Value
'        End Set
'    End Property
'    <Category("PoleIntGroup"), Description(""), DisplayName("Local_Interference_Group_Id")>
'    Public Property local_group_id() As Integer?
'        Get
'            Return Me.prop_local_group_id
'        End Get
'        Set
'            Me.prop_local_group_id = Value
'        End Set
'    End Property
'    <Category("PoleIntGroup"), Description(""), DisplayName("Elev_Bot")>
'    Public Property elev_bot() As Double?
'        Get
'            Return Me.prop_elev_bot
'        End Get
'        Set
'            Me.prop_elev_bot = Value
'        End Set
'    End Property
'    <Category("PoleIntGroup"), Description(""), DisplayName("Elev_Top")>
'    Public Property elev_top() As Double?
'        Get
'            Return Me.prop_elev_top
'        End Get
'        Set
'            Me.prop_elev_top = Value
'        End Set
'    End Property
'    <Category("PoleIntGroup"), Description(""), DisplayName("Width")>
'    Public Property width() As Double?
'        Get
'            Return Me.prop_width
'        End Get
'        Set
'            Me.prop_width = Value
'        End Set
'    End Property
'    <Category("PoleIntGroup"), Description(""), DisplayName("Description")>
'    Public Property description() As String
'        Get
'            Return Me.prop_description
'        End Get
'        Set
'            Me.prop_description = Value
'        End Set
'    End Property
'    <Category("PoleIntGroup"), Description(""), DisplayName("QTY")>
'    Public Property qty() As Integer?
'        Get
'            Return Me.prop_qty
'        End Get
'        Set
'            Me.prop_qty = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PoleIntGroupDataRow As DataRow)
'        Try
'            If Not IsDBNull(CType(PoleIntGroupDataRow.Item("group_id"), Integer)) Then
'                Me.group_id = CType(PoleIntGroupDataRow.Item("group_id"), Integer)
'            Else
'                Me.group_id = 0
'            End If
'        Catch
'            Me.group_id = 0
'        End Try 'group_id
'        Try
'            If Not IsDBNull(CType(PoleIntGroupDataRow.Item("local_group_id"), Integer)) Then
'                Me.local_group_id = CType(PoleIntGroupDataRow.Item("local_group_id"), Integer)
'            Else
'                Me.local_group_id = Nothing
'            End If
'        Catch
'            Me.local_group_id = Nothing
'        End Try 'local_group_id
'        Try
'            If Not IsDBNull(CType(PoleIntGroupDataRow.Item("elev_bot"), Double)) Then
'                Me.elev_bot = CType(PoleIntGroupDataRow.Item("elev_bot"), Double)
'            Else
'                Me.elev_bot = Nothing
'            End If
'        Catch
'            Me.elev_bot = Nothing
'        End Try 'Elev_Bot
'        Try
'            If Not IsDBNull(CType(PoleIntGroupDataRow.Item("elev_top"), Double)) Then
'                Me.elev_top = CType(PoleIntGroupDataRow.Item("elev_top"), Double)
'            Else
'                Me.elev_top = Nothing
'            End If
'        Catch
'            Me.elev_top = Nothing
'        End Try 'Elev_Top
'        Try
'            If Not IsDBNull(CType(PoleIntGroupDataRow.Item("width"), Double)) Then
'                Me.width = CType(PoleIntGroupDataRow.Item("width"), Double)
'            Else
'                Me.width = Nothing
'            End If
'        Catch
'            Me.width = Nothing
'        End Try 'Width
'        Try
'            Me.description = CType(PoleIntGroupDataRow.Item("description"), String)
'        Catch
'            Me.description = ""
'        End Try 'Description
'        Try
'            If Not IsDBNull(CType(PoleIntGroupDataRow.Item("qty"), Integer)) Then
'                Me.qty = CType(PoleIntGroupDataRow.Item("qty"), Integer)
'            Else
'                Me.qty = Nothing
'            End If
'        Catch
'            Me.qty = Nothing
'        End Try 'QTY


'        'For Each PoleIntDetailDataRow As DataRow In ds.Tables("CCIpole Int Details SQL").Rows
'        '    Dim PoleIntDetailsRefID As Integer = CType(PoleIntDetailDataRow.Item("pole_structure_id"), Integer)
'        '    If PoleIntDetailsRefID = refID Then
'        '        Me.int_ids.Add(New PoleIntDetail(PoleIntDetailDataRow, refID))
'        '    End If
'        'Next 'Add Interference Details to CCIpole Object

'    End Sub

'#End Region

'End Class 'Add Interference Group to CCIpole Object

'Partial Public Class PoleIntDetail_old

'#Region "Define"
'    Private prop_int_id As Integer?
'    Private prop_local_group_id As Integer?
'    Private prop_local_int_id As Integer?
'    Private prop_pole_flat As Integer?
'    Private prop_horizontal_offset As Double?
'    Private prop_rotation As Double?
'    Private prop_note As String

'    <Category("PoleIntDetail"), Description(""), DisplayName("Interference_Id")>
'    Public Property int_id() As Integer?
'        Get
'            Return Me.prop_int_id
'        End Get
'        Set
'            Me.prop_int_id = Value
'        End Set
'    End Property
'    <Category("PoleIntDetail"), Description(""), DisplayName("Local_Int_Group_Id")>
'    Public Property local_group_id() As Integer?
'        Get
'            Return Me.prop_local_group_id
'        End Get
'        Set
'            Me.prop_local_group_id = Value
'        End Set
'    End Property
'    <Category("PoleIntDetail"), Description(""), DisplayName("Local_Int_Detail_Id")>
'    Public Property local_int_id() As Integer?
'        Get
'            Return Me.prop_local_int_id
'        End Get
'        Set
'            Me.prop_local_int_id = Value
'        End Set
'    End Property
'    <Category("PoleIntDetail"), Description(""), DisplayName("Pole_Flat")>
'    Public Property pole_flat() As Integer?
'        Get
'            Return Me.prop_pole_flat
'        End Get
'        Set
'            Me.prop_pole_flat = Value
'        End Set
'    End Property
'    <Category("PoleIntDetail"), Description(""), DisplayName("Horizontal_Offset")>
'    Public Property horizontal_offset() As Double?
'        Get
'            Return Me.prop_horizontal_offset
'        End Get
'        Set
'            Me.prop_horizontal_offset = Value
'        End Set
'    End Property
'    <Category("PoleIntDetail"), Description(""), DisplayName("Rotation")>
'    Public Property rotation() As Double?
'        Get
'            Return Me.prop_rotation
'        End Get
'        Set
'            Me.prop_rotation = Value
'        End Set
'    End Property
'    <Category("PoleIntDetail"), Description(""), DisplayName("Note")>
'    Public Property note() As String
'        Get
'            Return Me.prop_note
'        End Get
'        Set
'            Me.prop_note = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"

'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PoleIntDetailDataRow As DataRow)
'        Try
'            If Not IsDBNull(CType(PoleIntDetailDataRow.Item("int_id"), Integer)) Then
'                Me.int_id = CType(PoleIntDetailDataRow.Item("int_id"), Integer)
'            Else
'                Me.int_id = 0
'            End If
'        Catch
'            Me.int_id = 0
'        End Try 'Interference_Id
'        Try
'            If Not IsDBNull(CType(PoleIntDetailDataRow.Item("local_group_id"), Integer)) Then
'                Me.local_group_id = CType(PoleIntDetailDataRow.Item("local_group_id"), Integer)
'            Else
'                Me.local_group_id = Nothing
'            End If
'        Catch
'            Me.local_group_id = Nothing
'        End Try 'local_group_id
'        Try
'            If Not IsDBNull(CType(PoleIntDetailDataRow.Item("local_int_id"), Integer)) Then
'                Me.local_int_id = CType(PoleIntDetailDataRow.Item("local_int_id"), Integer)
'            Else
'                Me.local_int_id = Nothing
'            End If
'        Catch
'            Me.local_int_id = Nothing
'        End Try 'local_int_id
'        Try
'            If Not IsDBNull(CType(PoleIntDetailDataRow.Item("pole_flat"), Integer)) Then
'                Me.pole_flat = CType(PoleIntDetailDataRow.Item("pole_flat"), Integer)
'            Else
'                Me.pole_flat = Nothing
'            End If
'        Catch
'            Me.pole_flat = Nothing
'        End Try 'Pole_Flat
'        Try
'            If Not IsDBNull(CType(PoleIntDetailDataRow.Item("horizontal_offset"), Double)) Then
'                Me.horizontal_offset = CType(PoleIntDetailDataRow.Item("horizontal_offset"), Double)
'            Else
'                Me.horizontal_offset = Nothing
'            End If
'        Catch
'            Me.horizontal_offset = Nothing
'        End Try 'Horizontal_Offset
'        Try
'            If Not IsDBNull(CType(PoleIntDetailDataRow.Item("rotation"), Double)) Then
'                Me.rotation = CType(PoleIntDetailDataRow.Item("rotation"), Double)
'            Else
'                Me.rotation = Nothing
'            End If
'        Catch
'            Me.rotation = Nothing
'        End Try 'Rotation
'        Try
'            Me.note = CType(PoleIntDetailDataRow.Item("note"), String)
'        Catch
'            Me.note = ""
'        End Try 'Note
'    End Sub

'#End Region

'End Class 'Add Interference Detail to CCIpole Object

'Partial Public Class PoleReinfResults_old

'#Region "Define"
'    Private prop_work_order_seq_num As Double?
'    Private prop_section_id As Integer?
'    Private prop_reinf_group_id As Integer?
'    Private prop_local_section_id As Integer?
'    Private prop_local_reinf_group_id As Integer?
'    Private prop_result_lkup_value As Integer?
'    Private prop_rating As Double?

'    <Category("PoleReinfResults"), Description(""), DisplayName("Work_Order_Seq_Num")>
'    Public Property work_order_seq_num() As Double?
'        Get
'            Return Me.prop_work_order_seq_num
'        End Get
'        Set
'            Me.prop_work_order_seq_num = Value
'        End Set
'    End Property
'    <Category("PoleReinfResults"), Description(""), DisplayName("Section_Id")>
'    Public Property section_id() As Integer?
'        Get
'            Return Me.prop_section_id
'        End Get
'        Set
'            Me.prop_section_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfResults"), Description(""), DisplayName("Reinf_Group_Id")>
'    Public Property reinf_group_id() As Integer?
'        Get
'            Return Me.prop_reinf_group_id
'        End Get
'        Set
'            Me.prop_reinf_group_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfResults"), Description(""), DisplayName("Local_Section_ID")>
'    Public Property local_section_id() As Integer?
'        Get
'            Return Me.prop_local_section_id
'        End Get
'        Set
'            Me.prop_local_section_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfResults"), Description(""), DisplayName("local_reinf_group_id")>
'    Public Property local_reinf_group_id() As Integer?
'        Get
'            Return Me.prop_local_reinf_group_id
'        End Get
'        Set
'            Me.prop_local_reinf_group_id = Value
'        End Set
'    End Property
'    <Category("PoleReinfResults"), Description(""), DisplayName("Result_Lkup_Value")>
'    Public Property result_lkup_value() As Integer?
'        Get
'            Return Me.prop_result_lkup_value
'        End Get
'        Set
'            Me.prop_result_lkup_value = Value
'        End Set
'    End Property
'    <Category("PoleReinfResults"), Description(""), DisplayName("Rating")>
'    Public Property rating() As Double?
'        Get
'            Return Me.prop_rating
'        End Get
'        Set
'            Me.prop_rating = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PoleReinfResultsDataRow As DataRow)
'        Try
'            If Not IsDBNull(CType(PoleReinfResultsDataRow.Item("work_order_seq_num"), Double)) Then
'                Me.work_order_seq_num = CType(PoleReinfResultsDataRow.Item("work_order_seq_num"), Double)
'            Else
'                Me.work_order_seq_num = 0
'            End If
'        Catch
'            Me.work_order_seq_num = 0
'        End Try 'Work_Order_Seq_Num
'        Try
'            If Not IsDBNull(CType(PoleReinfResultsDataRow.Item("section_id"), Integer)) Then
'                Me.section_id = CType(PoleReinfResultsDataRow.Item("section_id"), Integer)
'            Else
'                Me.section_id = 0
'            End If
'        Catch
'            Me.section_id = 0
'        End Try 'Section_Id
'        Try
'            If Not IsDBNull(CType(PoleReinfResultsDataRow.Item("reinf_group_id"), Integer)) Then
'                Me.reinf_group_id = CType(PoleReinfResultsDataRow.Item("reinf_group_id"), Integer)
'            Else
'                Me.reinf_group_id = 0
'            End If
'        Catch
'            Me.reinf_group_id = 0
'        End Try 'Reinf_Group_Id
'        Try
'            If Not IsDBNull(CType(PoleReinfResultsDataRow.Item("local_section_id"), Integer)) Then
'                Me.local_section_id = CType(PoleReinfResultsDataRow.Item("local_section_id"), Integer)
'            Else
'                Me.local_section_id = Nothing
'            End If
'        Catch
'            Me.local_section_id = Nothing
'        End Try 'local_section_id
'        Try
'            If Not IsDBNull(CType(PoleReinfResultsDataRow.Item("local_reinf_group_id"), Integer)) Then
'                Me.local_reinf_group_id = CType(PoleReinfResultsDataRow.Item("local_reinf_group_id"), Integer)
'            Else
'                Me.local_reinf_group_id = Nothing
'            End If
'        Catch
'            Me.local_reinf_group_id = Nothing
'        End Try 'local_reinf_group_id
'        Try
'            If Not IsDBNull(CType(PoleReinfResultsDataRow.Item("result_lkup_value"), Integer)) Then
'                Me.result_lkup_value = CType(PoleReinfResultsDataRow.Item("result_lkup_value"), Integer)
'            Else
'                Me.result_lkup_value = Nothing
'            End If
'        Catch
'            Me.result_lkup_value = Nothing
'        End Try 'Result_Lkup_Value
'        Try
'            If Not IsDBNull(CType(PoleReinfResultsDataRow.Item("rating"), Double)) Then
'                Me.rating = CType(PoleReinfResultsDataRow.Item("rating"), Double)
'            Else
'                Me.rating = Nothing
'            End If
'        Catch
'            Me.rating = Nothing
'        End Try 'Rating
'    End Sub

'#End Region

'End Class 'Add CCIpole Results to CCIpole Object

'Partial Public Class PropReinf_old

'#Region "Define"
'    Private prop_reinf_db_id As Integer?
'    Private prop_local_id As Integer?                           'Local ID
'    Private prop_name As String
'    Private prop_type As String
'    Private prop_b As Double?
'    Private prop_h As Double?
'    Private prop_sr_diam As Double?
'    Private prop_channel_thkns_web As Double?
'    Private prop_channel_thkns_flange As Double?
'    Private prop_channel_eo As Double?
'    Private prop_channel_J As Double?
'    Private prop_channel_Cw As Double?
'    Private prop_area_gross As Double?
'    Private prop_centroid As Double?
'    Private prop_istension As Boolean
'    Private prop_matl_id As Integer?
'    Private prop_local_matl_id As Integer?                       'Local ID
'    Private prop_Ix As Double?
'    Private prop_Iy As Double?
'    Private prop_Lu As Double?
'    Private prop_Kx As Double?
'    Private prop_Ky As Double?
'    Private prop_bolt_hole_size As Double?
'    Private prop_area_net As Double?
'    Private prop_shear_lag As Double?
'    Private prop_connection_type_bot As String
'    Private prop_connection_cap_revF_bot As Double?
'    Private prop_connection_cap_revG_bot As Double?
'    Private prop_connection_cap_revH_bot As Double?
'    Private prop_bolt_type_id_bot As Integer?
'    Private prop_local_bolt_id_bot As Integer?                  'Local ID
'    Private prop_bolt_N_or_X_bot As String
'    Private prop_bolt_num_bot As Integer?
'    Private prop_bolt_spacing_bot As Double?
'    Private prop_bolt_edge_dist_bot As Double?
'    Private prop_FlangeOrBP_connected_bot As Boolean
'    Private prop_weld_grade_bot As Double?
'    Private prop_weld_trans_type_bot As String
'    Private prop_weld_trans_length_bot As Double?
'    Private prop_weld_groove_depth_bot As Double?
'    Private prop_weld_groove_angle_bot As Integer?
'    Private prop_weld_trans_fillet_size_bot As Double?
'    Private prop_weld_trans_eff_throat_bot As Double?
'    Private prop_weld_long_type_bot As String
'    Private prop_weld_long_length_bot As Double?
'    Private prop_weld_long_fillet_size_bot As Double?
'    Private prop_weld_long_eff_throat_bot As Double?
'    Private prop_top_bot_connections_symmetrical As Boolean
'    Private prop_connection_type_top As String
'    Private prop_connection_cap_revF_top As Double?
'    Private prop_connection_cap_revG_top As Double?
'    Private prop_connection_cap_revH_top As Double?
'    Private prop_bolt_type_id_top As Integer?
'    Private prop_local_bolt_id_top As Integer?                  'Local ID
'    Private prop_bolt_N_or_X_top As String
'    Private prop_bolt_num_top As Integer?
'    Private prop_bolt_spacing_top As Double?
'    Private prop_bolt_edge_dist_top As Double?
'    Private prop_FlangeOrBP_connected_top As Boolean
'    Private prop_weld_grade_top As Double?
'    Private prop_weld_trans_type_top As String
'    Private prop_weld_trans_length_top As Double?
'    Private prop_weld_groove_depth_top As Double?
'    Private prop_weld_groove_angle_top As Integer?
'    Private prop_weld_trans_fillet_size_top As Double?
'    Private prop_weld_trans_eff_throat_top As Double?
'    Private prop_weld_long_type_top As String
'    Private prop_weld_long_length_top As Double?
'    Private prop_weld_long_fillet_size_top As Double?
'    Private prop_weld_long_eff_throat_top As Double?
'    Private prop_conn_length_channel As Double?
'    Private prop_conn_length_bot As Double?
'    Private prop_conn_length_top As Double?
'    Private prop_cap_comp_xx_f As Double?
'    Private prop_cap_comp_yy_f As Double?
'    Private prop_cap_tens_yield_f As Double?
'    Private prop_cap_tens_rupture_f As Double?
'    Private prop_cap_shear_f As Double?
'    Private prop_cap_bolt_shear_bot_f As Double?
'    Private prop_cap_bolt_shear_top_f As Double?
'    Private prop_cap_boltshaft_bearing_nodeform_bot_f As Double?
'    Private prop_cap_boltshaft_bearing_deform_bot_f As Double?
'    Private prop_cap_boltshaft_bearing_nodeform_top_f As Double?
'    Private prop_cap_boltshaft_bearing_deform_top_f As Double?
'    Private prop_cap_boltreinf_bearing_nodeform_bot_f As Double?
'    Private prop_cap_boltreinf_bearing_deform_bot_f As Double?
'    Private prop_cap_boltreinf_bearing_nodeform_top_f As Double?
'    Private prop_cap_boltreinf_bearing_deform_top_f As Double?
'    Private prop_cap_weld_trans_bot_f As Double?
'    Private prop_cap_weld_long_bot_f As Double?
'    Private prop_cap_weld_trans_top_f As Double?
'    Private prop_cap_weld_long_top_f As Double?
'    Private prop_cap_comp_xx_g As Double?
'    Private prop_cap_comp_yy_g As Double?
'    Private prop_cap_tens_yield_g As Double?
'    Private prop_cap_tens_rupture_g As Double?
'    Private prop_cap_shear_g As Double?
'    Private prop_cap_bolt_shear_bot_g As Double?
'    Private prop_cap_bolt_shear_top_g As Double?
'    Private prop_cap_boltshaft_bearing_nodeform_bot_g As Double?
'    Private prop_cap_boltshaft_bearing_deform_bot_g As Double?
'    Private prop_cap_boltshaft_bearing_nodeform_top_g As Double?
'    Private prop_cap_boltshaft_bearing_deform_top_g As Double?
'    Private prop_cap_boltreinf_bearing_nodeform_bot_g As Double?
'    Private prop_cap_boltreinf_bearing_deform_bot_g As Double?
'    Private prop_cap_boltreinf_bearing_nodeform_top_g As Double?
'    Private prop_cap_boltreinf_bearing_deform_top_g As Double?
'    Private prop_cap_weld_trans_bot_g As Double?
'    Private prop_cap_weld_long_bot_g As Double?
'    Private prop_cap_weld_trans_top_g As Double?
'    Private prop_cap_weld_long_top_g As Double?
'    Private prop_cap_comp_xx_h As Double?
'    Private prop_cap_comp_yy_h As Double?
'    Private prop_cap_tens_yield_h As Double?
'    Private prop_cap_tens_rupture_h As Double?
'    Private prop_cap_shear_h As Double?
'    Private prop_cap_bolt_shear_bot_h As Double?
'    Private prop_cap_bolt_shear_top_h As Double?
'    Private prop_cap_boltshaft_bearing_nodeform_bot_h As Double?
'    Private prop_cap_boltshaft_bearing_deform_bot_h As Double?
'    Private prop_cap_boltshaft_bearing_nodeform_top_h As Double?
'    Private prop_cap_boltshaft_bearing_deform_top_h As Double?
'    Private prop_cap_boltreinf_bearing_nodeform_bot_h As Double?
'    Private prop_cap_boltreinf_bearing_deform_bot_h As Double?
'    Private prop_cap_boltreinf_bearing_nodeform_top_h As Double?
'    Private prop_cap_boltreinf_bearing_deform_top_h As Double?
'    Private prop_cap_weld_trans_bot_h As Double?
'    Private prop_cap_weld_long_bot_h As Double?
'    Private prop_cap_weld_trans_top_h As Double?
'    Private prop_cap_weld_long_top_h As Double?

'    <Category("PropFlatPlate"), Description(""), DisplayName("Reinf_Db_Id")>
'    Public Property reinf_db_id() As Integer?
'        Get
'            Return Me.prop_reinf_db_id
'        End Get
'        Set
'            Me.prop_reinf_db_id = Value
'        End Set
'    End Property
'    Public Property local_id() As Integer?
'        Get
'            Return Me.prop_local_id
'        End Get
'        Set
'            Me.prop_local_id = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Name")>
'    Public Property name() As String
'        Get
'            Return Me.prop_name
'        End Get
'        Set
'            Me.prop_name = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Type")>
'    Public Property type() As String
'        Get
'            Return Me.prop_type
'        End Get
'        Set
'            Me.prop_type = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("B")>
'    Public Property b() As Double?
'        Get
'            Return Me.prop_b
'        End Get
'        Set
'            Me.prop_b = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("H")>
'    Public Property h() As Double?
'        Get
'            Return Me.prop_h
'        End Get
'        Set
'            Me.prop_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Sr_Diam")>
'    Public Property sr_diam() As Double?
'        Get
'            Return Me.prop_sr_diam
'        End Get
'        Set
'            Me.prop_sr_diam = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Channel_Thkns_Web")>
'    Public Property channel_thkns_web() As Double?
'        Get
'            Return Me.prop_channel_thkns_web
'        End Get
'        Set
'            Me.prop_channel_thkns_web = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Channel_Thkns_Flange")>
'    Public Property channel_thkns_flange() As Double?
'        Get
'            Return Me.prop_channel_thkns_flange
'        End Get
'        Set
'            Me.prop_channel_thkns_flange = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Channel_Eo")>
'    Public Property channel_eo() As Double?
'        Get
'            Return Me.prop_channel_eo
'        End Get
'        Set
'            Me.prop_channel_eo = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Channel_J")>
'    Public Property channel_J() As Double?
'        Get
'            Return Me.prop_channel_J
'        End Get
'        Set
'            Me.prop_channel_J = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Channel_Cw")>
'    Public Property channel_Cw() As Double?
'        Get
'            Return Me.prop_channel_Cw
'        End Get
'        Set
'            Me.prop_channel_Cw = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Area_Gross")>
'    Public Property area_gross() As Double?
'        Get
'            Return Me.prop_area_gross
'        End Get
'        Set
'            Me.prop_area_gross = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Centroid")>
'    Public Property centroid() As Double?
'        Get
'            Return Me.prop_centroid
'        End Get
'        Set
'            Me.prop_centroid = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Istension")>
'    Public Property istension() As Boolean
'        Get
'            Return Me.prop_istension
'        End Get
'        Set
'            Me.prop_istension = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Matl_Id")>
'    Public Property matl_id() As Integer?
'        Get
'            Return Me.prop_matl_id
'        End Get
'        Set
'            Me.prop_matl_id = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Local_Matl_Id")>
'    Public Property local_matl_id() As Integer?
'        Get
'            Return Me.prop_local_matl_id
'        End Get
'        Set
'            Me.prop_local_matl_id = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Ix")>
'    Public Property Ix() As Double?
'        Get
'            Return Me.prop_Ix
'        End Get
'        Set
'            Me.prop_Ix = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Iy")>
'    Public Property Iy() As Double?
'        Get
'            Return Me.prop_Iy
'        End Get
'        Set
'            Me.prop_Iy = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Lu")>
'    Public Property Lu() As Double?
'        Get
'            Return Me.prop_Lu
'        End Get
'        Set
'            Me.prop_Lu = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Kx")>
'    Public Property Kx() As Double?
'        Get
'            Return Me.prop_Kx
'        End Get
'        Set
'            Me.prop_Kx = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Ky")>
'    Public Property Ky() As Double?
'        Get
'            Return Me.prop_Ky
'        End Get
'        Set
'            Me.prop_Ky = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_Hole_Size")>
'    Public Property bolt_hole_size() As Double?
'        Get
'            Return Me.prop_bolt_hole_size
'        End Get
'        Set
'            Me.prop_bolt_hole_size = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Area_Net")>
'    Public Property area_net() As Double?
'        Get
'            Return Me.prop_area_net
'        End Get
'        Set
'            Me.prop_area_net = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Shear_Lag")>
'    Public Property shear_lag() As Double?
'        Get
'            Return Me.prop_shear_lag
'        End Get
'        Set
'            Me.prop_shear_lag = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Connection_Type_Bot")>
'    Public Property connection_type_bot() As String
'        Get
'            Return Me.prop_connection_type_bot
'        End Get
'        Set
'            Me.prop_connection_type_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Connection_Cap_Revf_Bot")>
'    Public Property connection_cap_revF_bot() As Double?
'        Get
'            Return Me.prop_connection_cap_revF_bot
'        End Get
'        Set
'            Me.prop_connection_cap_revF_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Connection_Cap_Revg_Bot")>
'    Public Property connection_cap_revG_bot() As Double?
'        Get
'            Return Me.prop_connection_cap_revG_bot
'        End Get
'        Set
'            Me.prop_connection_cap_revG_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Connection_Cap_Revh_Bot")>
'    Public Property connection_cap_revH_bot() As Double?
'        Get
'            Return Me.prop_connection_cap_revH_bot
'        End Get
'        Set
'            Me.prop_connection_cap_revH_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_Id_Bot")>
'    Public Property bolt_type_id_bot() As Integer?
'        Get
'            Return Me.prop_bolt_type_id_bot
'        End Get
'        Set
'            Me.prop_bolt_type_id_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Local_Bolt_Id_Bot")>
'    Public Property local_bolt_id_bot() As Integer?
'        Get
'            Return Me.prop_local_bolt_id_bot
'        End Get
'        Set
'            Me.prop_local_bolt_id_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_N_Or_X_Bot")>
'    Public Property bolt_N_or_X_bot() As String
'        Get
'            Return Me.prop_bolt_N_or_X_bot
'        End Get
'        Set
'            Me.prop_bolt_N_or_X_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_Num_Bot")>
'    Public Property bolt_num_bot() As Integer?
'        Get
'            Return Me.prop_bolt_num_bot
'        End Get
'        Set
'            Me.prop_bolt_num_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_Spacing_Bot")>
'    Public Property bolt_spacing_bot() As Double?
'        Get
'            Return Me.prop_bolt_spacing_bot
'        End Get
'        Set
'            Me.prop_bolt_spacing_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_Edge_Dist_Bot")>
'    Public Property bolt_edge_dist_bot() As Double?
'        Get
'            Return Me.prop_bolt_edge_dist_bot
'        End Get
'        Set
'            Me.prop_bolt_edge_dist_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Flangeorbp_Connected_Bot")>
'    Public Property FlangeOrBP_connected_bot() As Boolean
'        Get
'            Return Me.prop_FlangeOrBP_connected_bot
'        End Get
'        Set
'            Me.prop_FlangeOrBP_connected_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Grade_Bot")>
'    Public Property weld_grade_bot() As Double?
'        Get
'            Return Me.prop_weld_grade_bot
'        End Get
'        Set
'            Me.prop_weld_grade_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Trans_Type_Bot")>
'    Public Property weld_trans_type_bot() As String
'        Get
'            Return Me.prop_weld_trans_type_bot
'        End Get
'        Set
'            Me.prop_weld_trans_type_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Trans_Length_Bot")>
'    Public Property weld_trans_length_bot() As Double?
'        Get
'            Return Me.prop_weld_trans_length_bot
'        End Get
'        Set
'            Me.prop_weld_trans_length_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Groove_Depth_Bot")>
'    Public Property weld_groove_depth_bot() As Double?
'        Get
'            Return Me.prop_weld_groove_depth_bot
'        End Get
'        Set
'            Me.prop_weld_groove_depth_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Groove_Angle_Bot")>
'    Public Property weld_groove_angle_bot() As Integer?
'        Get
'            Return Me.prop_weld_groove_angle_bot
'        End Get
'        Set
'            Me.prop_weld_groove_angle_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Trans_Fillet_Size_Bot")>
'    Public Property weld_trans_fillet_size_bot() As Double?
'        Get
'            Return Me.prop_weld_trans_fillet_size_bot
'        End Get
'        Set
'            Me.prop_weld_trans_fillet_size_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Trans_Eff_Throat_Bot")>
'    Public Property weld_trans_eff_throat_bot() As Double?
'        Get
'            Return Me.prop_weld_trans_eff_throat_bot
'        End Get
'        Set
'            Me.prop_weld_trans_eff_throat_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Long_Type_Bot")>
'    Public Property weld_long_type_bot() As String
'        Get
'            Return Me.prop_weld_long_type_bot
'        End Get
'        Set
'            Me.prop_weld_long_type_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Long_Length_Bot")>
'    Public Property weld_long_length_bot() As Double?
'        Get
'            Return Me.prop_weld_long_length_bot
'        End Get
'        Set
'            Me.prop_weld_long_length_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Long_Fillet_Size_Bot")>
'    Public Property weld_long_fillet_size_bot() As Double?
'        Get
'            Return Me.prop_weld_long_fillet_size_bot
'        End Get
'        Set
'            Me.prop_weld_long_fillet_size_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Long_Eff_Throat_Bot")>
'    Public Property weld_long_eff_throat_bot() As Double?
'        Get
'            Return Me.prop_weld_long_eff_throat_bot
'        End Get
'        Set
'            Me.prop_weld_long_eff_throat_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Top_Bot_Connections_Symmetrical")>
'    Public Property top_bot_connections_symmetrical() As Boolean
'        Get
'            Return Me.prop_top_bot_connections_symmetrical
'        End Get
'        Set
'            Me.prop_top_bot_connections_symmetrical = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Connection_Type_Top")>
'    Public Property connection_type_top() As String
'        Get
'            Return Me.prop_connection_type_top
'        End Get
'        Set
'            Me.prop_connection_type_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Connection_Cap_Revf_Top")>
'    Public Property connection_cap_revF_top() As Double?
'        Get
'            Return Me.prop_connection_cap_revF_top
'        End Get
'        Set
'            Me.prop_connection_cap_revF_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Connection_Cap_Revg_Top")>
'    Public Property connection_cap_revG_top() As Double?
'        Get
'            Return Me.prop_connection_cap_revG_top
'        End Get
'        Set
'            Me.prop_connection_cap_revG_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Connection_Cap_Revh_Top")>
'    Public Property connection_cap_revH_top() As Double?
'        Get
'            Return Me.prop_connection_cap_revH_top
'        End Get
'        Set
'            Me.prop_connection_cap_revH_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_Id_Top")>
'    Public Property bolt_type_id_top() As Integer?
'        Get
'            Return Me.prop_bolt_type_id_top
'        End Get
'        Set
'            Me.prop_bolt_type_id_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Local_Bolt_Id_Top")>
'    Public Property local_bolt_id_top() As Integer?
'        Get
'            Return Me.prop_local_bolt_id_top
'        End Get
'        Set
'            Me.prop_local_bolt_id_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_N_Or_X_Top")>
'    Public Property bolt_N_or_X_top() As String
'        Get
'            Return Me.prop_bolt_N_or_X_top
'        End Get
'        Set
'            Me.prop_bolt_N_or_X_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_Num_Top")>
'    Public Property bolt_num_top() As Integer?
'        Get
'            Return Me.prop_bolt_num_top
'        End Get
'        Set
'            Me.prop_bolt_num_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_Spacing_Top")>
'    Public Property bolt_spacing_top() As Double?
'        Get
'            Return Me.prop_bolt_spacing_top
'        End Get
'        Set
'            Me.prop_bolt_spacing_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Bolt_Edge_Dist_Top")>
'    Public Property bolt_edge_dist_top() As Double?
'        Get
'            Return Me.prop_bolt_edge_dist_top
'        End Get
'        Set
'            Me.prop_bolt_edge_dist_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Flangeorbp_Connected_Top")>
'    Public Property FlangeOrBP_connected_top() As Boolean
'        Get
'            Return Me.prop_FlangeOrBP_connected_top
'        End Get
'        Set
'            Me.prop_FlangeOrBP_connected_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Grade_Top")>
'    Public Property weld_grade_top() As Double?
'        Get
'            Return Me.prop_weld_grade_top
'        End Get
'        Set
'            Me.prop_weld_grade_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Trans_Type_Top")>
'    Public Property weld_trans_type_top() As String
'        Get
'            Return Me.prop_weld_trans_type_top
'        End Get
'        Set
'            Me.prop_weld_trans_type_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Trans_Length_Top")>
'    Public Property weld_trans_length_top() As Double?
'        Get
'            Return Me.prop_weld_trans_length_top
'        End Get
'        Set
'            Me.prop_weld_trans_length_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Groove_Depth_Top")>
'    Public Property weld_groove_depth_top() As Double?
'        Get
'            Return Me.prop_weld_groove_depth_top
'        End Get
'        Set
'            Me.prop_weld_groove_depth_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Groove_Angle_Top")>
'    Public Property weld_groove_angle_top() As Integer?
'        Get
'            Return Me.prop_weld_groove_angle_top
'        End Get
'        Set
'            Me.prop_weld_groove_angle_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Trans_Fillet_Size_Top")>
'    Public Property weld_trans_fillet_size_top() As Double?
'        Get
'            Return Me.prop_weld_trans_fillet_size_top
'        End Get
'        Set
'            Me.prop_weld_trans_fillet_size_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Trans_Eff_Throat_Top")>
'    Public Property weld_trans_eff_throat_top() As Double?
'        Get
'            Return Me.prop_weld_trans_eff_throat_top
'        End Get
'        Set
'            Me.prop_weld_trans_eff_throat_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Long_Type_Top")>
'    Public Property weld_long_type_top() As String
'        Get
'            Return Me.prop_weld_long_type_top
'        End Get
'        Set
'            Me.prop_weld_long_type_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Long_Length_Top")>
'    Public Property weld_long_length_top() As Double?
'        Get
'            Return Me.prop_weld_long_length_top
'        End Get
'        Set
'            Me.prop_weld_long_length_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Long_Fillet_Size_Top")>
'    Public Property weld_long_fillet_size_top() As Double?
'        Get
'            Return Me.prop_weld_long_fillet_size_top
'        End Get
'        Set
'            Me.prop_weld_long_fillet_size_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Weld_Long_Eff_Throat_Top")>
'    Public Property weld_long_eff_throat_top() As Double?
'        Get
'            Return Me.prop_weld_long_eff_throat_top
'        End Get
'        Set
'            Me.prop_weld_long_eff_throat_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("conn_length_channel")>
'    Public Property conn_length_channel() As Double?
'        Get
'            Return Me.prop_conn_length_channel
'        End Get
'        Set
'            Me.prop_conn_length_channel = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Conn_Length_Bot")>
'    Public Property conn_length_bot() As Double?
'        Get
'            Return Me.prop_conn_length_bot
'        End Get
'        Set
'            Me.prop_conn_length_bot = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Conn_Length_Top")>
'    Public Property conn_length_top() As Double?
'        Get
'            Return Me.prop_conn_length_top
'        End Get
'        Set
'            Me.prop_conn_length_top = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Comp_Xx_F")>
'    Public Property cap_comp_xx_f() As Double?
'        Get
'            Return Me.prop_cap_comp_xx_f
'        End Get
'        Set
'            Me.prop_cap_comp_xx_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Comp_Yy_F")>
'    Public Property cap_comp_yy_f() As Double?
'        Get
'            Return Me.prop_cap_comp_yy_f
'        End Get
'        Set
'            Me.prop_cap_comp_yy_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Tens_Yield_F")>
'    Public Property cap_tens_yield_f() As Double?
'        Get
'            Return Me.prop_cap_tens_yield_f
'        End Get
'        Set
'            Me.prop_cap_tens_yield_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Tens_Rupture_F")>
'    Public Property cap_tens_rupture_f() As Double?
'        Get
'            Return Me.prop_cap_tens_rupture_f
'        End Get
'        Set
'            Me.prop_cap_tens_rupture_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Shear_F")>
'    Public Property cap_shear_f() As Double?
'        Get
'            Return Me.prop_cap_shear_f
'        End Get
'        Set
'            Me.prop_cap_shear_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Bolt_Shear_Bot_F")>
'    Public Property cap_bolt_shear_bot_f() As Double?
'        Get
'            Return Me.prop_cap_bolt_shear_bot_f
'        End Get
'        Set
'            Me.prop_cap_bolt_shear_bot_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Bolt_Shear_Top_F")>
'    Public Property cap_bolt_shear_top_f() As Double?
'        Get
'            Return Me.prop_cap_bolt_shear_top_f
'        End Get
'        Set
'            Me.prop_cap_bolt_shear_top_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Nodeform_Bot_F")>
'    Public Property cap_boltshaft_bearing_nodeform_bot_f() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_nodeform_bot_f
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_nodeform_bot_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Deform_Bot_F")>
'    Public Property cap_boltshaft_bearing_deform_bot_f() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_deform_bot_f
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_deform_bot_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Nodeform_Top_F")>
'    Public Property cap_boltshaft_bearing_nodeform_top_f() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_nodeform_top_f
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_nodeform_top_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Deform_Top_F")>
'    Public Property cap_boltshaft_bearing_deform_top_f() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_deform_top_f
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_deform_top_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Nodeform_Bot_F")>
'    Public Property cap_boltreinf_bearing_nodeform_bot_f() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_nodeform_bot_f
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_nodeform_bot_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Deform_Bot_F")>
'    Public Property cap_boltreinf_bearing_deform_bot_f() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_deform_bot_f
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_deform_bot_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Nodeform_Top_F")>
'    Public Property cap_boltreinf_bearing_nodeform_top_f() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_nodeform_top_f
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_nodeform_top_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Deform_Top_F")>
'    Public Property cap_boltreinf_bearing_deform_top_f() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_deform_top_f
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_deform_top_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Trans_Bot_F")>
'    Public Property cap_weld_trans_bot_f() As Double?
'        Get
'            Return Me.prop_cap_weld_trans_bot_f
'        End Get
'        Set
'            Me.prop_cap_weld_trans_bot_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Long_Bot_F")>
'    Public Property cap_weld_long_bot_f() As Double?
'        Get
'            Return Me.prop_cap_weld_long_bot_f
'        End Get
'        Set
'            Me.prop_cap_weld_long_bot_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Trans_Top_F")>
'    Public Property cap_weld_trans_top_f() As Double?
'        Get
'            Return Me.prop_cap_weld_trans_top_f
'        End Get
'        Set
'            Me.prop_cap_weld_trans_top_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Long_Top_F")>
'    Public Property cap_weld_long_top_f() As Double?
'        Get
'            Return Me.prop_cap_weld_long_top_f
'        End Get
'        Set
'            Me.prop_cap_weld_long_top_f = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Comp_Xx_G")>
'    Public Property cap_comp_xx_g() As Double?
'        Get
'            Return Me.prop_cap_comp_xx_g
'        End Get
'        Set
'            Me.prop_cap_comp_xx_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Comp_Yy_G")>
'    Public Property cap_comp_yy_g() As Double?
'        Get
'            Return Me.prop_cap_comp_yy_g
'        End Get
'        Set
'            Me.prop_cap_comp_yy_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Tens_Yield_G")>
'    Public Property cap_tens_yield_g() As Double?
'        Get
'            Return Me.prop_cap_tens_yield_g
'        End Get
'        Set
'            Me.prop_cap_tens_yield_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Tens_Rupture_G")>
'    Public Property cap_tens_rupture_g() As Double?
'        Get
'            Return Me.prop_cap_tens_rupture_g
'        End Get
'        Set
'            Me.prop_cap_tens_rupture_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Shear_G")>
'    Public Property cap_shear_g() As Double?
'        Get
'            Return Me.prop_cap_shear_g
'        End Get
'        Set
'            Me.prop_cap_shear_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Bolt_Shear_Bot_G")>
'    Public Property cap_bolt_shear_bot_g() As Double?
'        Get
'            Return Me.prop_cap_bolt_shear_bot_g
'        End Get
'        Set
'            Me.prop_cap_bolt_shear_bot_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Bolt_Shear_Top_G")>
'    Public Property cap_bolt_shear_top_g() As Double?
'        Get
'            Return Me.prop_cap_bolt_shear_top_g
'        End Get
'        Set
'            Me.prop_cap_bolt_shear_top_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Nodeform_Bot_G")>
'    Public Property cap_boltshaft_bearing_nodeform_bot_g() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_nodeform_bot_g
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_nodeform_bot_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Deform_Bot_G")>
'    Public Property cap_boltshaft_bearing_deform_bot_g() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_deform_bot_g
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_deform_bot_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Nodeform_Top_G")>
'    Public Property cap_boltshaft_bearing_nodeform_top_g() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_nodeform_top_g
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_nodeform_top_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Deform_Top_G")>
'    Public Property cap_boltshaft_bearing_deform_top_g() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_deform_top_g
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_deform_top_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Nodeform_Bot_G")>
'    Public Property cap_boltreinf_bearing_nodeform_bot_g() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_nodeform_bot_g
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_nodeform_bot_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Deform_Bot_G")>
'    Public Property cap_boltreinf_bearing_deform_bot_g() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_deform_bot_g
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_deform_bot_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Nodeform_Top_G")>
'    Public Property cap_boltreinf_bearing_nodeform_top_g() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_nodeform_top_g
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_nodeform_top_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Deform_Top_G")>
'    Public Property cap_boltreinf_bearing_deform_top_g() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_deform_top_g
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_deform_top_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Trans_Bot_G")>
'    Public Property cap_weld_trans_bot_g() As Double?
'        Get
'            Return Me.prop_cap_weld_trans_bot_g
'        End Get
'        Set
'            Me.prop_cap_weld_trans_bot_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Long_Bot_G")>
'    Public Property cap_weld_long_bot_g() As Double?
'        Get
'            Return Me.prop_cap_weld_long_bot_g
'        End Get
'        Set
'            Me.prop_cap_weld_long_bot_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Trans_Top_G")>
'    Public Property cap_weld_trans_top_g() As Double?
'        Get
'            Return Me.prop_cap_weld_trans_top_g
'        End Get
'        Set
'            Me.prop_cap_weld_trans_top_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Long_Top_G")>
'    Public Property cap_weld_long_top_g() As Double?
'        Get
'            Return Me.prop_cap_weld_long_top_g
'        End Get
'        Set
'            Me.prop_cap_weld_long_top_g = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Comp_Xx_H")>
'    Public Property cap_comp_xx_h() As Double?
'        Get
'            Return Me.prop_cap_comp_xx_h
'        End Get
'        Set
'            Me.prop_cap_comp_xx_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Comp_Yy_H")>
'    Public Property cap_comp_yy_h() As Double?
'        Get
'            Return Me.prop_cap_comp_yy_h
'        End Get
'        Set
'            Me.prop_cap_comp_yy_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Tens_Yield_H")>
'    Public Property cap_tens_yield_h() As Double?
'        Get
'            Return Me.prop_cap_tens_yield_h
'        End Get
'        Set
'            Me.prop_cap_tens_yield_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Tens_Rupture_H")>
'    Public Property cap_tens_rupture_h() As Double?
'        Get
'            Return Me.prop_cap_tens_rupture_h
'        End Get
'        Set
'            Me.prop_cap_tens_rupture_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Shear_H")>
'    Public Property cap_shear_h() As Double?
'        Get
'            Return Me.prop_cap_shear_h
'        End Get
'        Set
'            Me.prop_cap_shear_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Bolt_Shear_Bot_H")>
'    Public Property cap_bolt_shear_bot_h() As Double?
'        Get
'            Return Me.prop_cap_bolt_shear_bot_h
'        End Get
'        Set
'            Me.prop_cap_bolt_shear_bot_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Bolt_Shear_Top_H")>
'    Public Property cap_bolt_shear_top_h() As Double?
'        Get
'            Return Me.prop_cap_bolt_shear_top_h
'        End Get
'        Set
'            Me.prop_cap_bolt_shear_top_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Nodeform_Bot_H")>
'    Public Property cap_boltshaft_bearing_nodeform_bot_h() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_nodeform_bot_h
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_nodeform_bot_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Deform_Bot_H")>
'    Public Property cap_boltshaft_bearing_deform_bot_h() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_deform_bot_h
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_deform_bot_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Nodeform_Top_H")>
'    Public Property cap_boltshaft_bearing_nodeform_top_h() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_nodeform_top_h
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_nodeform_top_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltshaft_Bearing_Deform_Top_H")>
'    Public Property cap_boltshaft_bearing_deform_top_h() As Double?
'        Get
'            Return Me.prop_cap_boltshaft_bearing_deform_top_h
'        End Get
'        Set
'            Me.prop_cap_boltshaft_bearing_deform_top_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Nodeform_Bot_H")>
'    Public Property cap_boltreinf_bearing_nodeform_bot_h() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_nodeform_bot_h
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_nodeform_bot_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Deform_Bot_H")>
'    Public Property cap_boltreinf_bearing_deform_bot_h() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_deform_bot_h
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_deform_bot_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Nodeform_Top_H")>
'    Public Property cap_boltreinf_bearing_nodeform_top_h() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_nodeform_top_h
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_nodeform_top_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Boltreinf_Bearing_Deform_Top_H")>
'    Public Property cap_boltreinf_bearing_deform_top_h() As Double?
'        Get
'            Return Me.prop_cap_boltreinf_bearing_deform_top_h
'        End Get
'        Set
'            Me.prop_cap_boltreinf_bearing_deform_top_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Trans_Bot_H")>
'    Public Property cap_weld_trans_bot_h() As Double?
'        Get
'            Return Me.prop_cap_weld_trans_bot_h
'        End Get
'        Set
'            Me.prop_cap_weld_trans_bot_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Long_Bot_H")>
'    Public Property cap_weld_long_bot_h() As Double?
'        Get
'            Return Me.prop_cap_weld_long_bot_h
'        End Get
'        Set
'            Me.prop_cap_weld_long_bot_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Trans_Top_H")>
'    Public Property cap_weld_trans_top_h() As Double?
'        Get
'            Return Me.prop_cap_weld_trans_top_h
'        End Get
'        Set
'            Me.prop_cap_weld_trans_top_h = Value
'        End Set
'    End Property
'    <Category("PropFlatPlate"), Description(""), DisplayName("Cap_Weld_Long_Top_H")>
'    Public Property cap_weld_long_top_h() As Double?
'        Get
'            Return Me.prop_cap_weld_long_top_h
'        End Get
'        Set
'            Me.prop_cap_weld_long_top_h = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PropReinfDataRow As DataRow)
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("reinf_db_id"), Integer)) Then
'                Me.reinf_db_id = CType(PropReinfDataRow.Item("reinf_db_id"), Integer)
'            Else
'                Me.reinf_db_id = 0
'            End If
'        Catch
'            Me.reinf_db_id = 0
'        End Try 'Reinf_Db_Id
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("local_id"), Integer)) Then
'                Me.local_id = CType(PropReinfDataRow.Item("local_id"), Integer)
'            Else
'                Me.local_id = Nothing
'            End If
'        Catch
'            Me.local_id = Nothing
'        End Try 'local_id
'        Try
'            Me.name = CType(PropReinfDataRow.Item("name"), String)
'        Catch
'            Me.name = ""
'        End Try 'Name
'        Try
'            Me.type = CType(PropReinfDataRow.Item("type"), String)
'        Catch
'            Me.type = ""
'        End Try 'Type
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("b"), Double)) Then
'                Me.b = CType(PropReinfDataRow.Item("b"), Double)
'            Else
'                Me.b = Nothing
'            End If
'        Catch
'            Me.b = Nothing
'        End Try 'B
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("h"), Double)) Then
'                Me.h = CType(PropReinfDataRow.Item("h"), Double)
'            Else
'                Me.h = Nothing
'            End If
'        Catch
'            Me.h = Nothing
'        End Try 'H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("sr_diam"), Double)) Then
'                Me.sr_diam = CType(PropReinfDataRow.Item("sr_diam"), Double)
'            Else
'                Me.sr_diam = Nothing
'            End If
'        Catch
'            Me.sr_diam = Nothing
'        End Try 'Sr_Diam
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("channel_thkns_web"), Double)) Then
'                Me.channel_thkns_web = CType(PropReinfDataRow.Item("channel_thkns_web"), Double)
'            Else
'                Me.channel_thkns_web = Nothing
'            End If
'        Catch
'            Me.channel_thkns_web = Nothing
'        End Try 'Channel_Thkns_Web
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("channel_thkns_flange"), Double)) Then
'                Me.channel_thkns_flange = CType(PropReinfDataRow.Item("channel_thkns_flange"), Double)
'            Else
'                Me.channel_thkns_flange = Nothing
'            End If
'        Catch
'            Me.channel_thkns_flange = Nothing
'        End Try 'Channel_Thkns_Flange
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("channel_eo"), Double)) Then
'                Me.channel_eo = CType(PropReinfDataRow.Item("channel_eo"), Double)
'            Else
'                Me.channel_eo = Nothing
'            End If
'        Catch
'            Me.channel_eo = Nothing
'        End Try 'Channel_Eo
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("channel_J"), Double)) Then
'                Me.channel_J = CType(PropReinfDataRow.Item("channel_J"), Double)
'            Else
'                Me.channel_J = Nothing
'            End If
'        Catch
'            Me.channel_J = Nothing
'        End Try 'Channel_J
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("channel_Cw"), Double)) Then
'                Me.channel_Cw = CType(PropReinfDataRow.Item("channel_Cw"), Double)
'            Else
'                Me.channel_Cw = Nothing
'            End If
'        Catch
'            Me.channel_Cw = Nothing
'        End Try 'Channel_Cw
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("area_gross"), Double)) Then
'                Me.area_gross = CType(PropReinfDataRow.Item("area_gross"), Double)
'            Else
'                Me.area_gross = Nothing
'            End If
'        Catch
'            Me.area_gross = Nothing
'        End Try 'Area_Gross
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("centroid"), Double)) Then
'                Me.centroid = CType(PropReinfDataRow.Item("centroid"), Double)
'            Else
'                Me.centroid = Nothing
'            End If
'        Catch
'            Me.centroid = Nothing
'        End Try 'Centroid
'        Try
'            Me.istension = CType(PropReinfDataRow.Item("istension"), Boolean)
'        Catch
'            Me.istension = False
'        End Try 'Istension
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("matl_id"), Integer)) Then
'                Me.matl_id = CType(PropReinfDataRow.Item("matl_id"), Integer)
'            Else
'                Me.matl_id = 0
'            End If
'        Catch
'            Me.matl_id = 0
'        End Try 'Matl_Id
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("local_matl_id"), Integer)) Then
'                Me.local_matl_id = CType(PropReinfDataRow.Item("local_matl_id"), Integer)
'            Else
'                Me.local_matl_id = Nothing
'            End If
'        Catch
'            Me.local_matl_id = Nothing
'        End Try 'local_matl_id
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("Ix"), Double)) Then
'                Me.Ix = CType(PropReinfDataRow.Item("Ix"), Double)
'            Else
'                Me.Ix = Nothing
'            End If
'        Catch
'            Me.Ix = Nothing
'        End Try 'Ix
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("Iy"), Double)) Then
'                Me.Iy = CType(PropReinfDataRow.Item("Iy"), Double)
'            Else
'                Me.Iy = Nothing
'            End If
'        Catch
'            Me.Iy = Nothing
'        End Try 'Iy
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("Lu"), Double)) Then
'                Me.Lu = CType(PropReinfDataRow.Item("Lu"), Double)
'            Else
'                Me.Lu = Nothing
'            End If
'        Catch
'            Me.Lu = Nothing
'        End Try 'Lu
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("Kx"), Double)) Then
'                Me.Kx = CType(PropReinfDataRow.Item("Kx"), Double)
'            Else
'                Me.Kx = Nothing
'            End If
'        Catch
'            Me.Kx = Nothing
'        End Try 'Kx
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("Ky"), Double)) Then
'                Me.Ky = CType(PropReinfDataRow.Item("Ky"), Double)
'            Else
'                Me.Ky = Nothing
'            End If
'        Catch
'            Me.Ky = Nothing
'        End Try 'Ky
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("bolt_hole_size"), Double)) Then
'                Me.bolt_hole_size = CType(PropReinfDataRow.Item("bolt_hole_size"), Double)
'            Else
'                Me.bolt_hole_size = Nothing
'            End If
'        Catch
'            Me.bolt_hole_size = Nothing
'        End Try 'Bolt_Hole_Size
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("area_net"), Double)) Then
'                Me.area_net = CType(PropReinfDataRow.Item("area_net"), Double)
'            Else
'                Me.area_net = Nothing
'            End If
'        Catch
'            Me.area_net = Nothing
'        End Try 'Area_Net
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("shear_lag"), Double)) Then
'                Me.shear_lag = CType(PropReinfDataRow.Item("shear_lag"), Double)
'            Else
'                Me.shear_lag = Nothing
'            End If
'        Catch
'            Me.shear_lag = Nothing
'        End Try 'Shear_Lag
'        Try
'            Me.connection_type_bot = CType(PropReinfDataRow.Item("connection_type_bot"), String)
'        Catch
'            Me.connection_type_bot = ""
'        End Try 'Connection_Type_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("connection_cap_revF_bot"), Double)) Then
'                Me.connection_cap_revF_bot = CType(PropReinfDataRow.Item("connection_cap_revF_bot"), Double)
'            Else
'                Me.connection_cap_revF_bot = Nothing
'            End If
'        Catch
'            Me.connection_cap_revF_bot = Nothing
'        End Try 'Connection_Cap_Revf_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("connection_cap_revG_bot"), Double)) Then
'                Me.connection_cap_revG_bot = CType(PropReinfDataRow.Item("connection_cap_revG_bot"), Double)
'            Else
'                Me.connection_cap_revG_bot = Nothing
'            End If
'        Catch
'            Me.connection_cap_revG_bot = Nothing
'        End Try 'Connection_Cap_Revg_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("connection_cap_revH_bot"), Double)) Then
'                Me.connection_cap_revH_bot = CType(PropReinfDataRow.Item("connection_cap_revH_bot"), Double)
'            Else
'                Me.connection_cap_revH_bot = Nothing
'            End If
'        Catch
'            Me.connection_cap_revH_bot = Nothing
'        End Try 'Connection_Cap_Revh_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("bolt_type_id_bot"), Integer)) Then
'                Me.bolt_type_id_bot = CType(PropReinfDataRow.Item("bolt_type_id_bot"), Integer)
'            Else
'                Me.bolt_type_id_bot = 0
'            End If
'        Catch
'            Me.bolt_type_id_bot = 0
'        End Try 'bolt_type_id_bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("local_bolt_id_bot"), Integer)) Then
'                Me.local_bolt_id_bot = CType(PropReinfDataRow.Item("local_bolt_id_bot"), Integer)
'            Else
'                Me.local_bolt_id_bot = Nothing
'            End If
'        Catch
'            Me.local_bolt_id_bot = Nothing
'        End Try 'local_bolt_id_bot
'        Try
'            Me.bolt_N_or_X_bot = CType(PropReinfDataRow.Item("bolt_N_or_X_bot"), String)
'        Catch
'            Me.bolt_N_or_X_bot = ""
'        End Try 'Bolt_N_Or_X_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("bolt_num_bot"), Integer)) Then
'                Me.bolt_num_bot = CType(PropReinfDataRow.Item("bolt_num_bot"), Integer)
'            Else
'                Me.bolt_num_bot = Nothing
'            End If
'        Catch
'            Me.bolt_num_bot = Nothing
'        End Try 'Bolt_Num_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("bolt_spacing_bot"), Double)) Then
'                Me.bolt_spacing_bot = CType(PropReinfDataRow.Item("bolt_spacing_bot"), Double)
'            Else
'                Me.bolt_spacing_bot = Nothing
'            End If
'        Catch
'            Me.bolt_spacing_bot = Nothing
'        End Try 'Bolt_Spacing_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("bolt_edge_dist_bot"), Double)) Then
'                Me.bolt_edge_dist_bot = CType(PropReinfDataRow.Item("bolt_edge_dist_bot"), Double)
'            Else
'                Me.bolt_edge_dist_bot = Nothing
'            End If
'        Catch
'            Me.bolt_edge_dist_bot = Nothing
'        End Try 'Bolt_Edge_Dist_Bot
'        Try
'            Me.FlangeOrBP_connected_bot = CType(PropReinfDataRow.Item("FlangeOrBP_connected_bot"), Boolean)
'        Catch
'            Me.FlangeOrBP_connected_bot = False
'        End Try 'Flangeorbp_Connected_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_grade_bot"), Double)) Then
'                Me.weld_grade_bot = CType(PropReinfDataRow.Item("weld_grade_bot"), Double)
'            Else
'                Me.weld_grade_bot = Nothing
'            End If
'        Catch
'            Me.weld_grade_bot = Nothing
'        End Try 'Weld_Grade_Bot
'        Try
'            Me.weld_trans_type_bot = CType(PropReinfDataRow.Item("weld_trans_type_bot"), String)
'        Catch
'            Me.weld_trans_type_bot = ""
'        End Try 'Weld_Trans_Type_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_trans_length_bot"), Double)) Then
'                Me.weld_trans_length_bot = CType(PropReinfDataRow.Item("weld_trans_length_bot"), Double)
'            Else
'                Me.weld_trans_length_bot = Nothing
'            End If
'        Catch
'            Me.weld_trans_length_bot = Nothing
'        End Try 'Weld_Trans_Length_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_groove_depth_bot"), Double)) Then
'                Me.weld_groove_depth_bot = CType(PropReinfDataRow.Item("weld_groove_depth_bot"), Double)
'            Else
'                Me.weld_groove_depth_bot = Nothing
'            End If
'        Catch
'            Me.weld_groove_depth_bot = Nothing
'        End Try 'Weld_Groove_Depth_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_groove_angle_bot"), Integer)) Then
'                Me.weld_groove_angle_bot = CType(PropReinfDataRow.Item("weld_groove_angle_bot"), Integer)
'            Else
'                Me.weld_groove_angle_bot = Nothing
'            End If
'        Catch
'            Me.weld_groove_angle_bot = Nothing
'        End Try 'Weld_Groove_Angle_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_trans_fillet_size_bot"), Double)) Then
'                Me.weld_trans_fillet_size_bot = CType(PropReinfDataRow.Item("weld_trans_fillet_size_bot"), Double)
'            Else
'                Me.weld_trans_fillet_size_bot = Nothing
'            End If
'        Catch
'            Me.weld_trans_fillet_size_bot = Nothing
'        End Try 'Weld_Trans_Fillet_Size_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_trans_eff_throat_bot"), Double)) Then
'                Me.weld_trans_eff_throat_bot = CType(PropReinfDataRow.Item("weld_trans_eff_throat_bot"), Double)
'            Else
'                Me.weld_trans_eff_throat_bot = Nothing
'            End If
'        Catch
'            Me.weld_trans_eff_throat_bot = Nothing
'        End Try 'Weld_Trans_Eff_Throat_Bot
'        Try
'            Me.weld_long_type_bot = CType(PropReinfDataRow.Item("weld_long_type_bot"), String)
'        Catch
'            Me.weld_long_type_bot = ""
'        End Try 'Weld_Long_Type_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_long_length_bot"), Double)) Then
'                Me.weld_long_length_bot = CType(PropReinfDataRow.Item("weld_long_length_bot"), Double)
'            Else
'                Me.weld_long_length_bot = Nothing
'            End If
'        Catch
'            Me.weld_long_length_bot = Nothing
'        End Try 'Weld_Long_Length_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_long_fillet_size_bot"), Double)) Then
'                Me.weld_long_fillet_size_bot = CType(PropReinfDataRow.Item("weld_long_fillet_size_bot"), Double)
'            Else
'                Me.weld_long_fillet_size_bot = Nothing
'            End If
'        Catch
'            Me.weld_long_fillet_size_bot = Nothing
'        End Try 'Weld_Long_Fillet_Size_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_long_eff_throat_bot"), Double)) Then
'                Me.weld_long_eff_throat_bot = CType(PropReinfDataRow.Item("weld_long_eff_throat_bot"), Double)
'            Else
'                Me.weld_long_eff_throat_bot = Nothing
'            End If
'        Catch
'            Me.weld_long_eff_throat_bot = Nothing
'        End Try 'Weld_Long_Eff_Throat_Bot
'        Try
'            Me.top_bot_connections_symmetrical = CType(PropReinfDataRow.Item("top_bot_connections_symmetrical"), Boolean)
'        Catch
'            Me.top_bot_connections_symmetrical = False
'        End Try 'Top_Bot_Connections_Symmetrical
'        Try
'            Me.connection_type_top = CType(PropReinfDataRow.Item("connection_type_top"), String)
'        Catch
'            Me.connection_type_top = ""
'        End Try 'Connection_Type_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("connection_cap_revF_top"), Double)) Then
'                Me.connection_cap_revF_top = CType(PropReinfDataRow.Item("connection_cap_revF_top"), Double)
'            Else
'                Me.connection_cap_revF_top = Nothing
'            End If
'        Catch
'            Me.connection_cap_revF_top = Nothing
'        End Try 'Connection_Cap_Revf_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("connection_cap_revG_top"), Double)) Then
'                Me.connection_cap_revG_top = CType(PropReinfDataRow.Item("connection_cap_revG_top"), Double)
'            Else
'                Me.connection_cap_revG_top = Nothing
'            End If
'        Catch
'            Me.connection_cap_revG_top = Nothing
'        End Try 'Connection_Cap_Revg_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("connection_cap_revH_top"), Double)) Then
'                Me.connection_cap_revH_top = CType(PropReinfDataRow.Item("connection_cap_revH_top"), Double)
'            Else
'                Me.connection_cap_revH_top = Nothing
'            End If
'        Catch
'            Me.connection_cap_revH_top = Nothing
'        End Try 'Connection_Cap_Revh_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("bolt_type_id_top"), Integer)) Then
'                Me.bolt_type_id_top = CType(PropReinfDataRow.Item("bolt_type_id_top"), Integer)
'            Else
'                Me.bolt_type_id_top = 0
'            End If
'        Catch
'            Me.bolt_type_id_top = 0
'        End Try 'bolt_type_id_top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("local_bolt_id_top"), Integer)) Then
'                Me.local_bolt_id_top = CType(PropReinfDataRow.Item("local_bolt_id_top"), Integer)
'            Else
'                Me.local_bolt_id_top = Nothing
'            End If
'        Catch
'            Me.local_bolt_id_top = Nothing
'        End Try 'local_bolt_id_top
'        Try
'            Me.bolt_N_or_X_top = CType(PropReinfDataRow.Item("bolt_N_or_X_top"), String)
'        Catch
'            Me.bolt_N_or_X_top = ""
'        End Try 'Bolt_N_Or_X_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("bolt_num_top"), Integer)) Then
'                Me.bolt_num_top = CType(PropReinfDataRow.Item("bolt_num_top"), Integer)
'            Else
'                Me.bolt_num_top = Nothing
'            End If
'        Catch
'            Me.bolt_num_top = Nothing
'        End Try 'Bolt_Num_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("bolt_spacing_top"), Double)) Then
'                Me.bolt_spacing_top = CType(PropReinfDataRow.Item("bolt_spacing_top"), Double)
'            Else
'                Me.bolt_spacing_top = Nothing
'            End If
'        Catch
'            Me.bolt_spacing_top = Nothing
'        End Try 'Bolt_Spacing_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("bolt_edge_dist_top"), Double)) Then
'                Me.bolt_edge_dist_top = CType(PropReinfDataRow.Item("bolt_edge_dist_top"), Double)
'            Else
'                Me.bolt_edge_dist_top = Nothing
'            End If
'        Catch
'            Me.bolt_edge_dist_top = Nothing
'        End Try 'Bolt_Edge_Dist_Top
'        Try
'            Me.FlangeOrBP_connected_top = CType(PropReinfDataRow.Item("FlangeOrBP_connected_top"), Boolean)
'        Catch
'            Me.FlangeOrBP_connected_top = False
'        End Try 'Flangeorbp_Connected_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_grade_top"), Double)) Then
'                Me.weld_grade_top = CType(PropReinfDataRow.Item("weld_grade_top"), Double)
'            Else
'                Me.weld_grade_top = Nothing
'            End If
'        Catch
'            Me.weld_grade_top = Nothing
'        End Try 'Weld_Grade_Top
'        Try
'            Me.weld_trans_type_top = CType(PropReinfDataRow.Item("weld_trans_type_top"), String)
'        Catch
'            Me.weld_trans_type_top = ""
'        End Try 'Weld_Trans_Type_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_trans_length_top"), Double)) Then
'                Me.weld_trans_length_top = CType(PropReinfDataRow.Item("weld_trans_length_top"), Double)
'            Else
'                Me.weld_trans_length_top = Nothing
'            End If
'        Catch
'            Me.weld_trans_length_top = Nothing
'        End Try 'Weld_Trans_Length_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_groove_depth_top"), Double)) Then
'                Me.weld_groove_depth_top = CType(PropReinfDataRow.Item("weld_groove_depth_top"), Double)
'            Else
'                Me.weld_groove_depth_top = Nothing
'            End If
'        Catch
'            Me.weld_groove_depth_top = Nothing
'        End Try 'Weld_Groove_Depth_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_groove_angle_top"), Integer)) Then
'                Me.weld_groove_angle_top = CType(PropReinfDataRow.Item("weld_groove_angle_top"), Integer)
'            Else
'                Me.weld_groove_angle_top = Nothing
'            End If
'        Catch
'            Me.weld_groove_angle_top = Nothing
'        End Try 'Weld_Groove_Angle_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_trans_fillet_size_top"), Double)) Then
'                Me.weld_trans_fillet_size_top = CType(PropReinfDataRow.Item("weld_trans_fillet_size_top"), Double)
'            Else
'                Me.weld_trans_fillet_size_top = Nothing
'            End If
'        Catch
'            Me.weld_trans_fillet_size_top = Nothing
'        End Try 'Weld_Trans_Fillet_Size_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_trans_eff_throat_top"), Double)) Then
'                Me.weld_trans_eff_throat_top = CType(PropReinfDataRow.Item("weld_trans_eff_throat_top"), Double)
'            Else
'                Me.weld_trans_eff_throat_top = Nothing
'            End If
'        Catch
'            Me.weld_trans_eff_throat_top = Nothing
'        End Try 'Weld_Trans_Eff_Throat_Top
'        Try
'            Me.weld_long_type_top = CType(PropReinfDataRow.Item("weld_long_type_top"), String)
'        Catch
'            Me.weld_long_type_top = ""
'        End Try 'Weld_Long_Type_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_long_length_top"), Double)) Then
'                Me.weld_long_length_top = CType(PropReinfDataRow.Item("weld_long_length_top"), Double)
'            Else
'                Me.weld_long_length_top = Nothing
'            End If
'        Catch
'            Me.weld_long_length_top = Nothing
'        End Try 'Weld_Long_Length_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_long_fillet_size_top"), Double)) Then
'                Me.weld_long_fillet_size_top = CType(PropReinfDataRow.Item("weld_long_fillet_size_top"), Double)
'            Else
'                Me.weld_long_fillet_size_top = Nothing
'            End If
'        Catch
'            Me.weld_long_fillet_size_top = Nothing
'        End Try 'Weld_Long_Fillet_Size_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("weld_long_eff_throat_top"), Double)) Then
'                Me.weld_long_eff_throat_top = CType(PropReinfDataRow.Item("weld_long_eff_throat_top"), Double)
'            Else
'                Me.weld_long_eff_throat_top = Nothing
'            End If
'        Catch
'            Me.weld_long_eff_throat_top = Nothing
'        End Try 'Weld_Long_Eff_Throat_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("conn_length_channel"), Double)) Then
'                Me.conn_length_channel = CType(PropReinfDataRow.Item("conn_length_channel"), Double)
'            Else
'                Me.conn_length_channel = Nothing
'            End If
'        Catch
'            Me.conn_length_channel = Nothing
'        End Try 'conn_length_channel
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("conn_length_bot"), Double)) Then
'                Me.conn_length_bot = CType(PropReinfDataRow.Item("conn_length_bot"), Double)
'            Else
'                Me.conn_length_bot = Nothing
'            End If
'        Catch
'            Me.conn_length_bot = Nothing
'        End Try 'Conn_Length_Bot
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("conn_length_top"), Double)) Then
'                Me.conn_length_top = CType(PropReinfDataRow.Item("conn_length_top"), Double)
'            Else
'                Me.conn_length_top = Nothing
'            End If
'        Catch
'            Me.conn_length_top = Nothing
'        End Try 'Conn_Length_Top
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_comp_xx_f"), Double)) Then
'                Me.cap_comp_xx_f = CType(PropReinfDataRow.Item("cap_comp_xx_f"), Double)
'            Else
'                Me.cap_comp_xx_f = Nothing
'            End If
'        Catch
'            Me.cap_comp_xx_f = Nothing
'        End Try 'Cap_Comp_Xx_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_comp_yy_f"), Double)) Then
'                Me.cap_comp_yy_f = CType(PropReinfDataRow.Item("cap_comp_yy_f"), Double)
'            Else
'                Me.cap_comp_yy_f = Nothing
'            End If
'        Catch
'            Me.cap_comp_yy_f = Nothing
'        End Try 'Cap_Comp_Yy_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_tens_yield_f"), Double)) Then
'                Me.cap_tens_yield_f = CType(PropReinfDataRow.Item("cap_tens_yield_f"), Double)
'            Else
'                Me.cap_tens_yield_f = Nothing
'            End If
'        Catch
'            Me.cap_tens_yield_f = Nothing
'        End Try 'Cap_Tens_Yield_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_tens_rupture_f"), Double)) Then
'                Me.cap_tens_rupture_f = CType(PropReinfDataRow.Item("cap_tens_rupture_f"), Double)
'            Else
'                Me.cap_tens_rupture_f = Nothing
'            End If
'        Catch
'            Me.cap_tens_rupture_f = Nothing
'        End Try 'Cap_Tens_Rupture_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_shear_f"), Double)) Then
'                Me.cap_shear_f = CType(PropReinfDataRow.Item("cap_shear_f"), Double)
'            Else
'                Me.cap_shear_f = Nothing
'            End If
'        Catch
'            Me.cap_shear_f = Nothing
'        End Try 'Cap_Shear_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_bolt_shear_bot_f"), Double)) Then
'                Me.cap_bolt_shear_bot_f = CType(PropReinfDataRow.Item("cap_bolt_shear_bot_f"), Double)
'            Else
'                Me.cap_bolt_shear_bot_f = Nothing
'            End If
'        Catch
'            Me.cap_bolt_shear_bot_f = Nothing
'        End Try 'Cap_Bolt_Shear_Bot_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_bolt_shear_top_f"), Double)) Then
'                Me.cap_bolt_shear_top_f = CType(PropReinfDataRow.Item("cap_bolt_shear_top_f"), Double)
'            Else
'                Me.cap_bolt_shear_top_f = Nothing
'            End If
'        Catch
'            Me.cap_bolt_shear_top_f = Nothing
'        End Try 'Cap_Bolt_Shear_Top_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_bot_f"), Double)) Then
'                Me.cap_boltshaft_bearing_nodeform_bot_f = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_bot_f"), Double)
'            Else
'                Me.cap_boltshaft_bearing_nodeform_bot_f = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_nodeform_bot_f = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Nodeform_Bot_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_bot_f"), Double)) Then
'                Me.cap_boltshaft_bearing_deform_bot_f = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_bot_f"), Double)
'            Else
'                Me.cap_boltshaft_bearing_deform_bot_f = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_deform_bot_f = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Deform_Bot_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_top_f"), Double)) Then
'                Me.cap_boltshaft_bearing_nodeform_top_f = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_top_f"), Double)
'            Else
'                Me.cap_boltshaft_bearing_nodeform_top_f = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_nodeform_top_f = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Nodeform_Top_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_top_f"), Double)) Then
'                Me.cap_boltshaft_bearing_deform_top_f = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_top_f"), Double)
'            Else
'                Me.cap_boltshaft_bearing_deform_top_f = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_deform_top_f = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Deform_Top_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_bot_f"), Double)) Then
'                Me.cap_boltreinf_bearing_nodeform_bot_f = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_bot_f"), Double)
'            Else
'                Me.cap_boltreinf_bearing_nodeform_bot_f = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_nodeform_bot_f = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Nodeform_Bot_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_bot_f"), Double)) Then
'                Me.cap_boltreinf_bearing_deform_bot_f = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_bot_f"), Double)
'            Else
'                Me.cap_boltreinf_bearing_deform_bot_f = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_deform_bot_f = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Deform_Bot_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_top_f"), Double)) Then
'                Me.cap_boltreinf_bearing_nodeform_top_f = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_top_f"), Double)
'            Else
'                Me.cap_boltreinf_bearing_nodeform_top_f = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_nodeform_top_f = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Nodeform_Top_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_top_f"), Double)) Then
'                Me.cap_boltreinf_bearing_deform_top_f = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_top_f"), Double)
'            Else
'                Me.cap_boltreinf_bearing_deform_top_f = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_deform_top_f = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Deform_Top_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_trans_bot_f"), Double)) Then
'                Me.cap_weld_trans_bot_f = CType(PropReinfDataRow.Item("cap_weld_trans_bot_f"), Double)
'            Else
'                Me.cap_weld_trans_bot_f = Nothing
'            End If
'        Catch
'            Me.cap_weld_trans_bot_f = Nothing
'        End Try 'Cap_Weld_Trans_Bot_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_long_bot_f"), Double)) Then
'                Me.cap_weld_long_bot_f = CType(PropReinfDataRow.Item("cap_weld_long_bot_f"), Double)
'            Else
'                Me.cap_weld_long_bot_f = Nothing
'            End If
'        Catch
'            Me.cap_weld_long_bot_f = Nothing
'        End Try 'Cap_Weld_Long_Bot_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_trans_top_f"), Double)) Then
'                Me.cap_weld_trans_top_f = CType(PropReinfDataRow.Item("cap_weld_trans_top_f"), Double)
'            Else
'                Me.cap_weld_trans_top_f = Nothing
'            End If
'        Catch
'            Me.cap_weld_trans_top_f = Nothing
'        End Try 'Cap_Weld_Trans_Top_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_long_top_f"), Double)) Then
'                Me.cap_weld_long_top_f = CType(PropReinfDataRow.Item("cap_weld_long_top_f"), Double)
'            Else
'                Me.cap_weld_long_top_f = Nothing
'            End If
'        Catch
'            Me.cap_weld_long_top_f = Nothing
'        End Try 'Cap_Weld_Long_Top_F
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_comp_xx_g"), Double)) Then
'                Me.cap_comp_xx_g = CType(PropReinfDataRow.Item("cap_comp_xx_g"), Double)
'            Else
'                Me.cap_comp_xx_g = Nothing
'            End If
'        Catch
'            Me.cap_comp_xx_g = Nothing
'        End Try 'Cap_Comp_Xx_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_comp_yy_g"), Double)) Then
'                Me.cap_comp_yy_g = CType(PropReinfDataRow.Item("cap_comp_yy_g"), Double)
'            Else
'                Me.cap_comp_yy_g = Nothing
'            End If
'        Catch
'            Me.cap_comp_yy_g = Nothing
'        End Try 'Cap_Comp_Yy_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_tens_yield_g"), Double)) Then
'                Me.cap_tens_yield_g = CType(PropReinfDataRow.Item("cap_tens_yield_g"), Double)
'            Else
'                Me.cap_tens_yield_g = Nothing
'            End If
'        Catch
'            Me.cap_tens_yield_g = Nothing
'        End Try 'Cap_Tens_Yield_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_tens_rupture_g"), Double)) Then
'                Me.cap_tens_rupture_g = CType(PropReinfDataRow.Item("cap_tens_rupture_g"), Double)
'            Else
'                Me.cap_tens_rupture_g = Nothing
'            End If
'        Catch
'            Me.cap_tens_rupture_g = Nothing
'        End Try 'Cap_Tens_Rupture_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_shear_g"), Double)) Then
'                Me.cap_shear_g = CType(PropReinfDataRow.Item("cap_shear_g"), Double)
'            Else
'                Me.cap_shear_g = Nothing
'            End If
'        Catch
'            Me.cap_shear_g = Nothing
'        End Try 'Cap_Shear_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_bolt_shear_bot_g"), Double)) Then
'                Me.cap_bolt_shear_bot_g = CType(PropReinfDataRow.Item("cap_bolt_shear_bot_g"), Double)
'            Else
'                Me.cap_bolt_shear_bot_g = Nothing
'            End If
'        Catch
'            Me.cap_bolt_shear_bot_g = Nothing
'        End Try 'Cap_Bolt_Shear_Bot_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_bolt_shear_top_g"), Double)) Then
'                Me.cap_bolt_shear_top_g = CType(PropReinfDataRow.Item("cap_bolt_shear_top_g"), Double)
'            Else
'                Me.cap_bolt_shear_top_g = Nothing
'            End If
'        Catch
'            Me.cap_bolt_shear_top_g = Nothing
'        End Try 'Cap_Bolt_Shear_Top_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_bot_g"), Double)) Then
'                Me.cap_boltshaft_bearing_nodeform_bot_g = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_bot_g"), Double)
'            Else
'                Me.cap_boltshaft_bearing_nodeform_bot_g = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_nodeform_bot_g = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Nodeform_Bot_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_bot_g"), Double)) Then
'                Me.cap_boltshaft_bearing_deform_bot_g = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_bot_g"), Double)
'            Else
'                Me.cap_boltshaft_bearing_deform_bot_g = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_deform_bot_g = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Deform_Bot_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_top_g"), Double)) Then
'                Me.cap_boltshaft_bearing_nodeform_top_g = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_top_g"), Double)
'            Else
'                Me.cap_boltshaft_bearing_nodeform_top_g = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_nodeform_top_g = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Nodeform_Top_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_top_g"), Double)) Then
'                Me.cap_boltshaft_bearing_deform_top_g = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_top_g"), Double)
'            Else
'                Me.cap_boltshaft_bearing_deform_top_g = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_deform_top_g = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Deform_Top_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_bot_g"), Double)) Then
'                Me.cap_boltreinf_bearing_nodeform_bot_g = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_bot_g"), Double)
'            Else
'                Me.cap_boltreinf_bearing_nodeform_bot_g = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_nodeform_bot_g = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Nodeform_Bot_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_bot_g"), Double)) Then
'                Me.cap_boltreinf_bearing_deform_bot_g = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_bot_g"), Double)
'            Else
'                Me.cap_boltreinf_bearing_deform_bot_g = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_deform_bot_g = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Deform_Bot_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_top_g"), Double)) Then
'                Me.cap_boltreinf_bearing_nodeform_top_g = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_top_g"), Double)
'            Else
'                Me.cap_boltreinf_bearing_nodeform_top_g = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_nodeform_top_g = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Nodeform_Top_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_top_g"), Double)) Then
'                Me.cap_boltreinf_bearing_deform_top_g = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_top_g"), Double)
'            Else
'                Me.cap_boltreinf_bearing_deform_top_g = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_deform_top_g = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Deform_Top_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_trans_bot_g"), Double)) Then
'                Me.cap_weld_trans_bot_g = CType(PropReinfDataRow.Item("cap_weld_trans_bot_g"), Double)
'            Else
'                Me.cap_weld_trans_bot_g = Nothing
'            End If
'        Catch
'            Me.cap_weld_trans_bot_g = Nothing
'        End Try 'Cap_Weld_Trans_Bot_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_long_bot_g"), Double)) Then
'                Me.cap_weld_long_bot_g = CType(PropReinfDataRow.Item("cap_weld_long_bot_g"), Double)
'            Else
'                Me.cap_weld_long_bot_g = Nothing
'            End If
'        Catch
'            Me.cap_weld_long_bot_g = Nothing
'        End Try 'Cap_Weld_Long_Bot_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_trans_top_g"), Double)) Then
'                Me.cap_weld_trans_top_g = CType(PropReinfDataRow.Item("cap_weld_trans_top_g"), Double)
'            Else
'                Me.cap_weld_trans_top_g = Nothing
'            End If
'        Catch
'            Me.cap_weld_trans_top_g = Nothing
'        End Try 'Cap_Weld_Trans_Top_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_long_top_g"), Double)) Then
'                Me.cap_weld_long_top_g = CType(PropReinfDataRow.Item("cap_weld_long_top_g"), Double)
'            Else
'                Me.cap_weld_long_top_g = Nothing
'            End If
'        Catch
'            Me.cap_weld_long_top_g = Nothing
'        End Try 'Cap_Weld_Long_Top_G
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_comp_xx_h"), Double)) Then
'                Me.cap_comp_xx_h = CType(PropReinfDataRow.Item("cap_comp_xx_h"), Double)
'            Else
'                Me.cap_comp_xx_h = Nothing
'            End If
'        Catch
'            Me.cap_comp_xx_h = Nothing
'        End Try 'Cap_Comp_Xx_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_comp_yy_h"), Double)) Then
'                Me.cap_comp_yy_h = CType(PropReinfDataRow.Item("cap_comp_yy_h"), Double)
'            Else
'                Me.cap_comp_yy_h = Nothing
'            End If
'        Catch
'            Me.cap_comp_yy_h = Nothing
'        End Try 'Cap_Comp_Yy_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_tens_yield_h"), Double)) Then
'                Me.cap_tens_yield_h = CType(PropReinfDataRow.Item("cap_tens_yield_h"), Double)
'            Else
'                Me.cap_tens_yield_h = Nothing
'            End If
'        Catch
'            Me.cap_tens_yield_h = Nothing
'        End Try 'Cap_Tens_Yield_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_tens_rupture_h"), Double)) Then
'                Me.cap_tens_rupture_h = CType(PropReinfDataRow.Item("cap_tens_rupture_h"), Double)
'            Else
'                Me.cap_tens_rupture_h = Nothing
'            End If
'        Catch
'            Me.cap_tens_rupture_h = Nothing
'        End Try 'Cap_Tens_Rupture_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_shear_h"), Double)) Then
'                Me.cap_shear_h = CType(PropReinfDataRow.Item("cap_shear_h"), Double)
'            Else
'                Me.cap_shear_h = Nothing
'            End If
'        Catch
'            Me.cap_shear_h = Nothing
'        End Try 'Cap_Shear_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_bolt_shear_bot_h"), Double)) Then
'                Me.cap_bolt_shear_bot_h = CType(PropReinfDataRow.Item("cap_bolt_shear_bot_h"), Double)
'            Else
'                Me.cap_bolt_shear_bot_h = Nothing
'            End If
'        Catch
'            Me.cap_bolt_shear_bot_h = Nothing
'        End Try 'Cap_Bolt_Shear_Bot_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_bolt_shear_top_h"), Double)) Then
'                Me.cap_bolt_shear_top_h = CType(PropReinfDataRow.Item("cap_bolt_shear_top_h"), Double)
'            Else
'                Me.cap_bolt_shear_top_h = Nothing
'            End If
'        Catch
'            Me.cap_bolt_shear_top_h = Nothing
'        End Try 'Cap_Bolt_Shear_Top_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_bot_h"), Double)) Then
'                Me.cap_boltshaft_bearing_nodeform_bot_h = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_bot_h"), Double)
'            Else
'                Me.cap_boltshaft_bearing_nodeform_bot_h = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_nodeform_bot_h = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Nodeform_Bot_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_bot_h"), Double)) Then
'                Me.cap_boltshaft_bearing_deform_bot_h = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_bot_h"), Double)
'            Else
'                Me.cap_boltshaft_bearing_deform_bot_h = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_deform_bot_h = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Deform_Bot_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_top_h"), Double)) Then
'                Me.cap_boltshaft_bearing_nodeform_top_h = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_nodeform_top_h"), Double)
'            Else
'                Me.cap_boltshaft_bearing_nodeform_top_h = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_nodeform_top_h = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Nodeform_Top_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_top_h"), Double)) Then
'                Me.cap_boltshaft_bearing_deform_top_h = CType(PropReinfDataRow.Item("cap_boltshaft_bearing_deform_top_h"), Double)
'            Else
'                Me.cap_boltshaft_bearing_deform_top_h = Nothing
'            End If
'        Catch
'            Me.cap_boltshaft_bearing_deform_top_h = Nothing
'        End Try 'Cap_Boltshaft_Bearing_Deform_Top_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_bot_h"), Double)) Then
'                Me.cap_boltreinf_bearing_nodeform_bot_h = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_bot_h"), Double)
'            Else
'                Me.cap_boltreinf_bearing_nodeform_bot_h = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_nodeform_bot_h = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Nodeform_Bot_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_bot_h"), Double)) Then
'                Me.cap_boltreinf_bearing_deform_bot_h = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_bot_h"), Double)
'            Else
'                Me.cap_boltreinf_bearing_deform_bot_h = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_deform_bot_h = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Deform_Bot_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_top_h"), Double)) Then
'                Me.cap_boltreinf_bearing_nodeform_top_h = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_nodeform_top_h"), Double)
'            Else
'                Me.cap_boltreinf_bearing_nodeform_top_h = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_nodeform_top_h = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Nodeform_Top_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_top_h"), Double)) Then
'                Me.cap_boltreinf_bearing_deform_top_h = CType(PropReinfDataRow.Item("cap_boltreinf_bearing_deform_top_h"), Double)
'            Else
'                Me.cap_boltreinf_bearing_deform_top_h = Nothing
'            End If
'        Catch
'            Me.cap_boltreinf_bearing_deform_top_h = Nothing
'        End Try 'Cap_Boltreinf_Bearing_Deform_Top_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_trans_bot_h"), Double)) Then
'                Me.cap_weld_trans_bot_h = CType(PropReinfDataRow.Item("cap_weld_trans_bot_h"), Double)
'            Else
'                Me.cap_weld_trans_bot_h = Nothing
'            End If
'        Catch
'            Me.cap_weld_trans_bot_h = Nothing
'        End Try 'Cap_Weld_Trans_Bot_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_long_bot_h"), Double)) Then
'                Me.cap_weld_long_bot_h = CType(PropReinfDataRow.Item("cap_weld_long_bot_h"), Double)
'            Else
'                Me.cap_weld_long_bot_h = Nothing
'            End If
'        Catch
'            Me.cap_weld_long_bot_h = Nothing
'        End Try 'Cap_Weld_Long_Bot_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_trans_top_h"), Double)) Then
'                Me.cap_weld_trans_top_h = CType(PropReinfDataRow.Item("cap_weld_trans_top_h"), Double)
'            Else
'                Me.cap_weld_trans_top_h = Nothing
'            End If
'        Catch
'            Me.cap_weld_trans_top_h = Nothing
'        End Try 'Cap_Weld_Trans_Top_H
'        Try
'            If Not IsDBNull(CType(PropReinfDataRow.Item("cap_weld_long_top_h"), Double)) Then
'                Me.cap_weld_long_top_h = CType(PropReinfDataRow.Item("cap_weld_long_top_h"), Double)
'            Else
'                Me.cap_weld_long_top_h = Nothing
'            End If
'        Catch
'            Me.cap_weld_long_top_h = Nothing
'        End Try 'Cap_Weld_Long_Top_H
'    End Sub

'#End Region

'End Class 'Add Custom Reinforcement to CCIpole Object

'Partial Public Class PropBolt_old

'#Region "Define"
'    Private prop_bolt_db_id As Integer?
'    Private prop_local_id As Integer?
'    Private prop_name As String
'    Private prop_description As String
'    Private prop_diam As Double?
'    Private prop_area As Double?
'    Private prop_fu_bolt As Double?
'    Private prop_sleeve_diam_out As Double?
'    Private prop_sleeve_diam_in As Double?
'    Private prop_fu_sleeve As Double?
'    Private prop_bolt_n_sleeve_shear_revF As Double?
'    Private prop_bolt_x_sleeve_shear_revF As Double?
'    Private prop_bolt_n_sleeve_shear_revG As Double?
'    Private prop_bolt_x_sleeve_shear_revG As Double?
'    Private prop_bolt_n_sleeve_shear_revH As Double?
'    Private prop_bolt_x_sleeve_shear_revH As Double?
'    Private prop_rb_applied_revH As Boolean

'    <Category("PropBolt"), Description(""), DisplayName("Bolt_Db_Id")>
'    Public Property bolt_db_id() As Integer?
'        Get
'            Return Me.prop_bolt_db_id
'        End Get
'        Set
'            Me.prop_bolt_db_id = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Local_Id")>
'    Public Property local_id() As Integer?
'        Get
'            Return Me.prop_local_id
'        End Get
'        Set
'            Me.prop_local_id = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Name")>
'    Public Property name() As String
'        Get
'            Return Me.prop_name
'        End Get
'        Set
'            Me.prop_name = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Description")>
'    Public Property description() As String
'        Get
'            Return Me.prop_description
'        End Get
'        Set
'            Me.prop_description = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Diam")>
'    Public Property diam() As Double?
'        Get
'            Return Me.prop_diam
'        End Get
'        Set
'            Me.prop_diam = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Area")>
'    Public Property area() As Double?
'        Get
'            Return Me.prop_area
'        End Get
'        Set
'            Me.prop_area = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Fu_Bolt")>
'    Public Property fu_bolt() As Double?
'        Get
'            Return Me.prop_fu_bolt
'        End Get
'        Set
'            Me.prop_fu_bolt = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Sleeve_Diam_Out")>
'    Public Property sleeve_diam_out() As Double?
'        Get
'            Return Me.prop_sleeve_diam_out
'        End Get
'        Set
'            Me.prop_sleeve_diam_out = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Sleeve_Diam_In")>
'    Public Property sleeve_diam_in() As Double?
'        Get
'            Return Me.prop_sleeve_diam_in
'        End Get
'        Set
'            Me.prop_sleeve_diam_in = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Fu_Sleeve")>
'    Public Property fu_sleeve() As Double?
'        Get
'            Return Me.prop_fu_sleeve
'        End Get
'        Set
'            Me.prop_fu_sleeve = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Bolt_N_Sleeve_Shear_Revf")>
'    Public Property bolt_n_sleeve_shear_revF() As Double?
'        Get
'            Return Me.prop_bolt_n_sleeve_shear_revF
'        End Get
'        Set
'            Me.prop_bolt_n_sleeve_shear_revF = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Bolt_X_Sleeve_Shear_Revf")>
'    Public Property bolt_x_sleeve_shear_revF() As Double?
'        Get
'            Return Me.prop_bolt_x_sleeve_shear_revF
'        End Get
'        Set
'            Me.prop_bolt_x_sleeve_shear_revF = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Bolt_N_Sleeve_Shear_Revg")>
'    Public Property bolt_n_sleeve_shear_revG() As Double?
'        Get
'            Return Me.prop_bolt_n_sleeve_shear_revG
'        End Get
'        Set
'            Me.prop_bolt_n_sleeve_shear_revG = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Bolt_X_Sleeve_Shear_Revg")>
'    Public Property bolt_x_sleeve_shear_revG() As Double?
'        Get
'            Return Me.prop_bolt_x_sleeve_shear_revG
'        End Get
'        Set
'            Me.prop_bolt_x_sleeve_shear_revG = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Bolt_N_Sleeve_Shear_Revh")>
'    Public Property bolt_n_sleeve_shear_revH() As Double?
'        Get
'            Return Me.prop_bolt_n_sleeve_shear_revH
'        End Get
'        Set
'            Me.prop_bolt_n_sleeve_shear_revH = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Bolt_X_Sleeve_Shear_Revh")>
'    Public Property bolt_x_sleeve_shear_revH() As Double?
'        Get
'            Return Me.prop_bolt_x_sleeve_shear_revH
'        End Get
'        Set
'            Me.prop_bolt_x_sleeve_shear_revH = Value
'        End Set
'    End Property
'    <Category("PropBolt"), Description(""), DisplayName("Rb_Applied_Revh")>
'    Public Property rb_applied_revH() As Boolean
'        Get
'            Return Me.prop_rb_applied_revH
'        End Get
'        Set
'            Me.prop_rb_applied_revH = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PropBoltDataRow As DataRow)
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("bolt_db_id"), Integer)) Then
'                Me.bolt_db_id = CType(PropBoltDataRow.Item("bolt_db_id"), Integer)
'            Else
'                Me.bolt_db_id = 0
'            End If
'        Catch
'            Me.bolt_db_id = 0
'        End Try 'Bolt_Db_Id
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("local_id"), Integer)) Then
'                Me.local_id = CType(PropBoltDataRow.Item("local_id"), Integer)
'            Else
'                Me.local_id = Nothing
'            End If
'        Catch
'            Me.local_id = Nothing
'        End Try 'local_id
'        Try
'            Me.name = CType(PropBoltDataRow.Item("name"), String)
'        Catch
'            Me.name = ""
'        End Try 'Name
'        Try
'            Me.description = CType(PropBoltDataRow.Item("description"), String)
'        Catch
'            Me.description = ""
'        End Try 'Description
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("diam"), Double)) Then
'                Me.diam = CType(PropBoltDataRow.Item("diam"), Double)
'            Else
'                Me.diam = Nothing
'            End If
'        Catch
'            Me.diam = Nothing
'        End Try 'Diam
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("area"), Double)) Then
'                Me.area = CType(PropBoltDataRow.Item("area"), Double)
'            Else
'                Me.area = Nothing
'            End If
'        Catch
'            Me.area = Nothing
'        End Try 'Area
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("fu_bolt"), Double)) Then
'                Me.fu_bolt = CType(PropBoltDataRow.Item("fu_bolt"), Double)
'            Else
'                Me.fu_bolt = Nothing
'            End If
'        Catch
'            Me.fu_bolt = Nothing
'        End Try 'Fu_Bolt
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("sleeve_diam_out"), Double)) Then
'                Me.sleeve_diam_out = CType(PropBoltDataRow.Item("sleeve_diam_out"), Double)
'            Else
'                Me.sleeve_diam_out = Nothing
'            End If
'        Catch
'            Me.sleeve_diam_out = Nothing
'        End Try 'Sleeve_Diam_Out
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("sleeve_diam_in"), Double)) Then
'                Me.sleeve_diam_in = CType(PropBoltDataRow.Item("sleeve_diam_in"), Double)
'            Else
'                Me.sleeve_diam_in = Nothing
'            End If
'        Catch
'            Me.sleeve_diam_in = Nothing
'        End Try 'Sleeve_Diam_In
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("fu_sleeve"), Double)) Then
'                Me.fu_sleeve = CType(PropBoltDataRow.Item("fu_sleeve"), Double)
'            Else
'                Me.fu_sleeve = Nothing
'            End If
'        Catch
'            Me.fu_sleeve = Nothing
'        End Try 'Fu_Sleeve
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("bolt_n_sleeve_shear_revF"), Double)) Then
'                Me.bolt_n_sleeve_shear_revF = CType(PropBoltDataRow.Item("bolt_n_sleeve_shear_revF"), Double)
'            Else
'                Me.bolt_n_sleeve_shear_revF = Nothing
'            End If
'        Catch
'            Me.bolt_n_sleeve_shear_revF = Nothing
'        End Try 'Bolt_N_Sleeve_Shear_Revf
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("bolt_x_sleeve_shear_revF"), Double)) Then
'                Me.bolt_x_sleeve_shear_revF = CType(PropBoltDataRow.Item("bolt_x_sleeve_shear_revF"), Double)
'            Else
'                Me.bolt_x_sleeve_shear_revF = Nothing
'            End If
'        Catch
'            Me.bolt_x_sleeve_shear_revF = Nothing
'        End Try 'Bolt_X_Sleeve_Shear_Revf
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("bolt_n_sleeve_shear_revG"), Double)) Then
'                Me.bolt_n_sleeve_shear_revG = CType(PropBoltDataRow.Item("bolt_n_sleeve_shear_revG"), Double)
'            Else
'                Me.bolt_n_sleeve_shear_revG = Nothing
'            End If
'        Catch
'            Me.bolt_n_sleeve_shear_revG = Nothing
'        End Try 'Bolt_N_Sleeve_Shear_Revg
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("bolt_x_sleeve_shear_revG"), Double)) Then
'                Me.bolt_x_sleeve_shear_revG = CType(PropBoltDataRow.Item("bolt_x_sleeve_shear_revG"), Double)
'            Else
'                Me.bolt_x_sleeve_shear_revG = Nothing
'            End If
'        Catch
'            Me.bolt_x_sleeve_shear_revG = Nothing
'        End Try 'Bolt_X_Sleeve_Shear_Revg
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("bolt_n_sleeve_shear_revH"), Double)) Then
'                Me.bolt_n_sleeve_shear_revH = CType(PropBoltDataRow.Item("bolt_n_sleeve_shear_revH"), Double)
'            Else
'                Me.bolt_n_sleeve_shear_revH = Nothing
'            End If
'        Catch
'            Me.bolt_n_sleeve_shear_revH = Nothing
'        End Try 'Bolt_N_Sleeve_Shear_Revh
'        Try
'            If Not IsDBNull(CType(PropBoltDataRow.Item("bolt_x_sleeve_shear_revH"), Double)) Then
'                Me.bolt_x_sleeve_shear_revH = CType(PropBoltDataRow.Item("bolt_x_sleeve_shear_revH"), Double)
'            Else
'                Me.bolt_x_sleeve_shear_revH = Nothing
'            End If
'        Catch
'            Me.bolt_x_sleeve_shear_revH = Nothing
'        End Try 'Bolt_X_Sleeve_Shear_Revh
'        Try
'            Me.rb_applied_revH = CType(PropBoltDataRow.Item("rb_applied_revH"), Boolean)
'        Catch
'            Me.rb_applied_revH = False
'        End Try 'Rb_Applied_Revh
'    End Sub

'#End Region

'End Class 'Add Custom Bolt to CCIpole Object

'Partial Public Class PropMatl_old

'#Region "Define"
'    Private prop_matl_db_id As Integer?
'    Private prop_local_id As Integer?
'    Private prop_name As String
'    Private prop_fy As Double?
'    Private prop_fu As Double?

'    <Category("PropMatl"), Description(""), DisplayName("Matl_Db_Id")>
'    Public Property matl_db_id() As Integer?
'        Get
'            Return Me.prop_matl_db_id
'        End Get
'        Set
'            Me.prop_matl_db_id = Value
'        End Set
'    End Property
'    <Category("PropMatl"), Description(""), DisplayName("Local_Id")>
'    Public Property local_id() As Integer?
'        Get
'            Return Me.prop_local_id
'        End Get
'        Set
'            Me.prop_local_id = Value
'        End Set
'    End Property
'    <Category("PropMatl"), Description(""), DisplayName("Name")>
'    Public Property name() As String
'        Get
'            Return Me.prop_name
'        End Get
'        Set
'            Me.prop_name = Value
'        End Set
'    End Property
'    <Category("PropMatl"), Description(""), DisplayName("Fy")>
'    Public Property fy() As Double?
'        Get
'            Return Me.prop_fy
'        End Get
'        Set
'            Me.prop_fy = Value
'        End Set
'    End Property
'    <Category("PropMatl"), Description(""), DisplayName("Fu")>
'    Public Property fu() As Double?
'        Get
'            Return Me.prop_fu
'        End Get
'        Set
'            Me.prop_fu = Value
'        End Set
'    End Property

'#End Region

'#Region "Constructors"
'    Public Sub New()
'        'Leave Method Empty
'    End Sub

'    Public Sub New(ByVal PropMatlDataRow As DataRow)
'        Try
'            If Not IsDBNull(CType(PropMatlDataRow.Item("matl_db_id"), Integer)) Then
'                Me.matl_db_id = CType(PropMatlDataRow.Item("matl_db_id"), Integer)
'            Else
'                Me.matl_db_id = 0
'            End If
'        Catch
'            Me.matl_db_id = 0
'        End Try 'Matl_Db_Id
'        Try
'            If Not IsDBNull(CType(PropMatlDataRow.Item("local_id"), Integer)) Then
'                Me.local_id = CType(PropMatlDataRow.Item("local_id"), Integer)
'            Else
'                Me.local_id = Nothing
'            End If
'        Catch
'            Me.local_id = Nothing
'        End Try 'local_id
'        Try
'            Me.name = CType(PropMatlDataRow.Item("name"), String)
'        Catch
'            Me.name = ""
'        End Try 'Name
'        Try
'            If Not IsDBNull(CType(PropMatlDataRow.Item("fy"), Double)) Then
'                Me.fy = CType(PropMatlDataRow.Item("fy"), Double)
'            Else
'                Me.fy = Nothing
'            End If
'        Catch
'            Me.fy = Nothing
'        End Try 'Fy
'        Try
'            If Not IsDBNull(CType(PropMatlDataRow.Item("fu"), Double)) Then
'                Me.fu = CType(PropMatlDataRow.Item("fu"), Double)
'            Else
'                Me.fu = Nothing
'            End If
'        Catch
'            Me.fu = Nothing
'        End Try 'Fu
'    End Sub

'#End Region

'End Class 'Add Custom Material to CCIpole Object
