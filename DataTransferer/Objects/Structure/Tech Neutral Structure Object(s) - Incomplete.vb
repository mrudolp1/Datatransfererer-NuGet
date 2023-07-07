Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
Imports System.IO

Partial Public Class tower_structure
    Private prop_ID As Integer
    Private prop_bus_unit As Integer
    Private prop_mod_wo_seq_num As Integer
    Private prop_structure_id As String
    Private prop_base_twr_type As String
    Private prop_up_twr_type As String
    Private prop_twr_height As Double
    Private prop_base_str_height As Double
    Private prop_up_str_height As Double
    Private prop_base_elev As Double
    Private prop_tow_face_width As Double
    Private prop_base_face_width As Double
    Private prop_lat_pole_width As Double
    Private prop_lambda As Double
    Private prop_const_slope As Boolean
    Private prop_index_plate As Boolean
    Private prop_top_takup_lambda As Boolean
    Private prop_base_type As String
    Private prop_base_taper As Double
    Private prop_original_geometry As Boolean
    Private prop_existing_geometry As Boolean
    Private prop_modified_geometry As Boolean
    Private prop_lattice_sections As List(Of lattice_section)
    Private prop_pole_sections As List(Of pole_section)
    Private prop_guy_anchor_groups As List(Of guy_anchor_group)
    Private prop_guy_attachments As List(Of guy_attachment)

    Public Sub New()
        'Leave method empty
    End Sub

    <Category("Structure Model"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Bus_Unit")>
    Public Property bus_unit() As Integer
        Get
            Return Me.prop_bus_unit
        End Get
        Set
            Me.prop_bus_unit = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Mod_Wo_Seq_Num")>
    Public Property mod_wo_seq_num() As Integer
        Get
            Return Me.prop_mod_wo_seq_num
        End Get
        Set
            Me.prop_mod_wo_seq_num = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Structure_Id")>
    Public Property structure_id() As String
        Get
            Return Me.prop_structure_id
        End Get
        Set
            Me.prop_structure_id = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Base_Twr_Type")>
    Public Property base_twr_type() As String
        Get
            Return Me.prop_base_twr_type
        End Get
        Set
            Me.prop_base_twr_type = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Up_Twr_Type")>
    Public Property up_twr_type() As String
        Get
            Return Me.prop_up_twr_type
        End Get
        Set
            Me.prop_up_twr_type = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Twr_Height")>
    Public Property twr_height() As Double
        Get
            Return Me.prop_twr_height
        End Get
        Set
            Me.prop_twr_height = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Base_Str_Height")>
    Public Property base_str_height() As Double
        Get
            Return Me.prop_base_str_height
        End Get
        Set
            Me.prop_base_str_height = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Up_Str_Height")>
    Public Property up_str_height() As Double
        Get
            Return Me.prop_up_str_height
        End Get
        Set
            Me.prop_up_str_height = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Base_Elev")>
    Public Property base_elev() As Double
        Get
            Return Me.prop_base_elev
        End Get
        Set
            Me.prop_base_elev = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Tow_Face_Width")>
    Public Property tow_face_width() As Double
        Get
            Return Me.prop_tow_face_width
        End Get
        Set
            Me.prop_tow_face_width = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Base_Face_Width")>
    Public Property base_face_width() As Double
        Get
            Return Me.prop_base_face_width
        End Get
        Set
            Me.prop_base_face_width = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Lat_Pole_Width")>
    Public Property lat_pole_width() As Double
        Get
            Return Me.prop_lat_pole_width
        End Get
        Set
            Me.prop_lat_pole_width = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Lambda")>
    Public Property lambda() As Double
        Get
            Return Me.prop_lambda
        End Get
        Set
            Me.prop_lambda = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Const_Slope")>
    Public Property const_slope() As Boolean
        Get
            Return Me.prop_const_slope
        End Get
        Set
            Me.prop_const_slope = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Index_Plate")>
    Public Property index_plate() As Boolean
        Get
            Return Me.prop_index_plate
        End Get
        Set
            Me.prop_index_plate = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Top_Takup_Lambda")>
    Public Property top_takup_lambda() As Boolean
        Get
            Return Me.prop_top_takup_lambda
        End Get
        Set
            Me.prop_top_takup_lambda = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Base_Type")>
    Public Property base_type() As String
        Get
            Return Me.prop_base_type
        End Get
        Set
            Me.prop_base_type = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Base_Taper")>
    Public Property base_taper() As Double
        Get
            Return Me.prop_base_taper
        End Get
        Set
            Me.prop_base_taper = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Original_Geometry")>
    Public Property original_geometry() As Boolean
        Get
            Return Me.prop_original_geometry
        End Get
        Set
            Me.prop_original_geometry = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Existing_Geometry")>
    Public Property existing_geometry() As Boolean
        Get
            Return Me.prop_existing_geometry
        End Get
        Set
            Me.prop_existing_geometry = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Modified_Geometry")>
    Public Property modified_geometry() As Boolean
        Get
            Return Me.prop_modified_geometry
        End Get
        Set
            Me.prop_modified_geometry = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Lattice_Sections")>
    Public Property lattice_sections() As List(Of lattice_section)
        Get
            Return Me.prop_lattice_sections
        End Get
        Set
            Me.prop_lattice_sections = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Pole_Sections")>
    Public Property pole_sections() As List(Of pole_section)
        Get
            Return Me.prop_pole_sections
        End Get
        Set
            Me.prop_pole_sections = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Guy_Anchor_Groups")>
    Public Property guy_anchor_groups() As List(Of guy_anchor_group)
        Get
            Return Me.prop_guy_anchor_groups
        End Get
        Set
            Me.prop_guy_anchor_groups = Value
        End Set
    End Property
    <Category("Structure Model"), Description(""), DisplayName("Guy_Attachments")>
    Public Property guy_attachments() As List(Of guy_attachment)
        Get
            Return Me.prop_guy_attachments
        End Get
        Set
            Me.prop_guy_attachments = Value
        End Set
    End Property

End Class

#Region "Lattice"
Partial Public Class lattice_section
    Private prop_model_id As Integer
    Private prop_ID As Integer
    Private prop_section_num As Integer
    Private prop_upper_strc As Boolean
    Private prop_section_length As Double
    Private prop_sect_top_width As Double
    Private prop_sect_bot_width As Double
    Private prop_sect_top_girt_length As Double
    Private prop_sect_bot_girt_length As Double
    Private prop_sect_faces_identical As Boolean
    Private prop_sect_hips_identical As Boolean
    Private prop_sect_plans_identical As Boolean
    Private prop_sect_bracing_direct_to_leg As Boolean
    Private prop_sect_memb_end_adjust As Double
    Private prop_sect_Af As Boolean
    Private prop_sect_Ar As Boolean
    Private prop_sect_Ar_Ice As Boolean
    Private prop_sect_weight As Boolean
    Private prop_sect_conn_group_memb_type As Boolean
    Private prop_sect_check_gusset As Boolean
    Private prop_sect_dp_group_memb_type As Boolean
    Private prop_sect_cc_consider As Boolean
    Private prop_weight_mult As Double
    Private prop_wp_mult As Double
    Private prop_af_factor As Double
    Private prop_ar_factor As Double
    Private prop_round_area_ratio As Double
    Private prop_flat_area_ratio As Double

    Public Sub New()
        'Leave method empty
    End Sub

    <Category("Lattice Section"), Description(""), DisplayName("Model_Id")>
    Public Property model_id() As Integer
        Get
            Return Me.prop_model_id
        End Get
        Set
            Me.prop_model_id = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Section_Num")>
    Public Property section_num() As Integer
        Get
            Return Me.prop_section_num
        End Get
        Set
            Me.prop_section_num = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Upper_Strc")>
    Public Property upper_strc() As Boolean
        Get
            Return Me.prop_upper_strc
        End Get
        Set
            Me.prop_upper_strc = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Section_Length")>
    Public Property section_length() As Double
        Get
            Return Me.prop_section_length
        End Get
        Set
            Me.prop_section_length = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Top_Width")>
    Public Property sect_top_width() As Double
        Get
            Return Me.prop_sect_top_width
        End Get
        Set
            Me.prop_sect_top_width = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Bot_Width")>
    Public Property sect_bot_width() As Double
        Get
            Return Me.prop_sect_bot_width
        End Get
        Set
            Me.prop_sect_bot_width = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Top_Girt_Length")>
    Public Property sect_top_girt_length() As Double
        Get
            Return Me.prop_sect_top_girt_length
        End Get
        Set
            Me.prop_sect_top_girt_length = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Bot_Girt_Length")>
    Public Property sect_bot_girt_length() As Double
        Get
            Return Me.prop_sect_bot_girt_length
        End Get
        Set
            Me.prop_sect_bot_girt_length = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Faces_Identical")>
    Public Property sect_faces_identical() As Boolean
        Get
            Return Me.prop_sect_faces_identical
        End Get
        Set
            Me.prop_sect_faces_identical = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Hips_Identical")>
    Public Property sect_hips_identical() As Boolean
        Get
            Return Me.prop_sect_hips_identical
        End Get
        Set
            Me.prop_sect_hips_identical = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Plans_Identical")>
    Public Property sect_plans_identical() As Boolean
        Get
            Return Me.prop_sect_plans_identical
        End Get
        Set
            Me.prop_sect_plans_identical = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Bracing_Direct_To_Leg")>
    Public Property sect_bracing_direct_to_leg() As Boolean
        Get
            Return Me.prop_sect_bracing_direct_to_leg
        End Get
        Set
            Me.prop_sect_bracing_direct_to_leg = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Memb_End_Adjust")>
    Public Property sect_memb_end_adjust() As Double
        Get
            Return Me.prop_sect_memb_end_adjust
        End Get
        Set
            Me.prop_sect_memb_end_adjust = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Af")>
    Public Property sect_Af() As Boolean
        Get
            Return Me.prop_sect_Af
        End Get
        Set
            Me.prop_sect_Af = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Ar")>
    Public Property sect_Ar() As Boolean
        Get
            Return Me.prop_sect_Ar
        End Get
        Set
            Me.prop_sect_Ar = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Ar_Ice")>
    Public Property sect_Ar_Ice() As Boolean
        Get
            Return Me.prop_sect_Ar_Ice
        End Get
        Set
            Me.prop_sect_Ar_Ice = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Weight")>
    Public Property sect_weight() As Boolean
        Get
            Return Me.prop_sect_weight
        End Get
        Set
            Me.prop_sect_weight = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Conn_Group_Memb_Type")>
    Public Property sect_conn_group_memb_type() As Boolean
        Get
            Return Me.prop_sect_conn_group_memb_type
        End Get
        Set
            Me.prop_sect_conn_group_memb_type = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Check_Gusset")>
    Public Property sect_check_gusset() As Boolean
        Get
            Return Me.prop_sect_check_gusset
        End Get
        Set
            Me.prop_sect_check_gusset = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Dp_Group_Memb_Type")>
    Public Property sect_dp_group_memb_type() As Boolean
        Get
            Return Me.prop_sect_dp_group_memb_type
        End Get
        Set
            Me.prop_sect_dp_group_memb_type = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Sect_Cc_Consider")>
    Public Property sect_cc_consider() As Boolean
        Get
            Return Me.prop_sect_cc_consider
        End Get
        Set
            Me.prop_sect_cc_consider = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Weight_Mult")>
    Public Property weight_mult() As Double
        Get
            Return Me.prop_weight_mult
        End Get
        Set
            Me.prop_weight_mult = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Wp_Mult")>
    Public Property wp_mult() As Double
        Get
            Return Me.prop_wp_mult
        End Get
        Set
            Me.prop_wp_mult = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Af_Factor")>
    Public Property af_factor() As Double
        Get
            Return Me.prop_af_factor
        End Get
        Set
            Me.prop_af_factor = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Ar_Factor")>
    Public Property ar_factor() As Double
        Get
            Return Me.prop_ar_factor
        End Get
        Set
            Me.prop_ar_factor = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Round_Area_Ratio")>
    Public Property round_area_ratio() As Double
        Get
            Return Me.prop_round_area_ratio
        End Get
        Set
            Me.prop_round_area_ratio = Value
        End Set
    End Property
    <Category("Lattice Section"), Description(""), DisplayName("Flat_Area_Ratio")>
    Public Property flat_area_ratio() As Double
        Get
            Return Me.prop_flat_area_ratio
        End Get
        Set
            Me.prop_flat_area_ratio = Value
        End Set
    End Property

End Class
Partial Public Class lattice_face
    Private prop_section_id As Integer
    Private prop_ID As Integer
    Private prop_face_letter As String
    Private prop_face_bays_identical As Boolean

    <Category("Lattice Face"), Description(""), DisplayName("Section_Id")>
    Public Property section_id() As Integer
        Get
            Return Me.prop_section_id
        End Get
        Set
            Me.prop_section_id = Value
        End Set
    End Property
    <Category("Lattice Face"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Lattice Face"), Description(""), DisplayName("Face_Letter")>
    Public Property face_letter() As String
        Get
            Return Me.prop_face_letter
        End Get
        Set
            Me.prop_face_letter = Value
        End Set
    End Property
    <Category("Lattice Face"), Description(""), DisplayName("Face_Bays_Identical")>
    Public Property face_bays_identical() As Boolean
        Get
            Return Me.prop_face_bays_identical
        End Get
        Set
            Me.prop_face_bays_identical = Value
        End Set
    End Property

End Class
Partial Public Class lattice_face_bay
    Private prop_face_id As Integer
    Private prop_ID As Integer
    Private prop_bay_number As Integer
    Private prop_bay_length As Double
    Private prop_bay_face_pattern As String

    <Category("Lattice Face Bay"), Description(""), DisplayName("Face_Id")>
    Public Property face_id() As Integer
        Get
            Return Me.prop_face_id
        End Get
        Set
            Me.prop_face_id = Value
        End Set
    End Property
    <Category("Lattice Face Bay"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Lattice Face Bay"), Description(""), DisplayName("Bay_Number")>
    Public Property bay_number() As Integer
        Get
            Return Me.prop_bay_number
        End Get
        Set
            Me.prop_bay_number = Value
        End Set
    End Property
    <Category("Lattice Face Bay"), Description(""), DisplayName("Bay_Length")>
    Public Property bay_length() As Double
        Get
            Return Me.prop_bay_length
        End Get
        Set
            Me.prop_bay_length = Value
        End Set
    End Property
    <Category("Lattice Face Bay"), Description(""), DisplayName("Bay_Face_Pattern")>
    Public Property bay_face_pattern() As String
        Get
            Return Me.prop_bay_face_pattern
        End Get
        Set
            Me.prop_bay_face_pattern = Value
        End Set
    End Property

End Class
Partial Public Class lattice_hip
    Private prop_section_id As Integer
    Private prop_ID As Integer
    Private prop_leg_letter As String
    Private prop_face_bays_identical As Boolean

    <Category("Lattice Hip"), Description(""), DisplayName("Section_Id")>
    Public Property section_id() As Integer
        Get
            Return Me.prop_section_id
        End Get
        Set
            Me.prop_section_id = Value
        End Set
    End Property
    <Category("Lattice Hip"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Lattice Hip"), Description(""), DisplayName("Leg_Letter")>
    Public Property leg_letter() As String
        Get
            Return Me.prop_leg_letter
        End Get
        Set
            Me.prop_leg_letter = Value
        End Set
    End Property
    <Category("Lattice Hip"), Description(""), DisplayName("Face_Bays_Identical")>
    Public Property face_bays_identical() As Boolean
        Get
            Return Me.prop_face_bays_identical
        End Get
        Set
            Me.prop_face_bays_identical = Value
        End Set
    End Property

End Class
Partial Public Class lattice_hip_bay
    Private prop_hip_id As Integer
    Private prop_ID As Integer
    Private prop_bay_number As Integer
    Private prop_bay_length As Double
    Private prop_bay_hip_pattern As String

    <Category("Lattice Hip Bay"), Description(""), DisplayName("Hip_Id")>
    Public Property hip_id() As Integer
        Get
            Return Me.prop_hip_id
        End Get
        Set
            Me.prop_hip_id = Value
        End Set
    End Property
    <Category("Lattice Hip Bay"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Lattice Hip Bay"), Description(""), DisplayName("Bay_Number")>
    Public Property bay_number() As Integer
        Get
            Return Me.prop_bay_number
        End Get
        Set
            Me.prop_bay_number = Value
        End Set
    End Property
    <Category("Lattice Hip Bay"), Description(""), DisplayName("Bay_Length")>
    Public Property bay_length() As Double
        Get
            Return Me.prop_bay_length
        End Get
        Set
            Me.prop_bay_length = Value
        End Set
    End Property
    <Category("Lattice Hip Bay"), Description(""), DisplayName("Bay_Hip_Pattern")>
    Public Property bay_hip_pattern() As String
        Get
            Return Me.prop_bay_hip_pattern
        End Get
        Set
            Me.prop_bay_hip_pattern = Value
        End Set
    End Property

End Class
Partial Public Class lattice_plan
    Private prop_section_id As Integer
    Private prop_ID As Integer
    Private prop_plan_elevation As Double
    Private prop_plan_pattern As String

    <Category("Lattice Plan"), Description(""), DisplayName("Section_Id")>
    Public Property section_id() As Integer
        Get
            Return Me.prop_section_id
        End Get
        Set
            Me.prop_section_id = Value
        End Set
    End Property
    <Category("Lattice Plan"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Lattice Plan"), Description(""), DisplayName("Plan_Elevation")>
    Public Property plan_elevation() As Double
        Get
            Return Me.prop_plan_elevation
        End Get
        Set
            Me.prop_plan_elevation = Value
        End Set
    End Property
    <Category("Lattice Plan"), Description(""), DisplayName("Plan_Pattern")>
    Public Property plan_pattern() As String
        Get
            Return Me.prop_plan_pattern
        End Get
        Set
            Me.prop_plan_pattern = Value
        End Set
    End Property

End Class
Partial Public Class lattice_bracing_detail
    Private prop_bracing_id As Integer
    Private prop_ID As Integer
    Private prop_bracing_pattern_type As String
    Private prop_bracing_type As String
    Private prop_bracing_sect_type As String
    Private prop_bracing_sect_prop As Integer
    Private prop_bracing_mat_prop As Integer
    Private prop_bracing_conn_end_cond As String
    Private prop_bracing_conn_mirror As Boolean
    Private prop_bracing_conn_start_pattern As String
    Private prop_bracing_conn_start_bolt_size As Integer
    Private prop_bracing_conn_start_bolt_mat_prop As Integer
    Private prop_bracing_conn_start_bolt_thread As String
    Private prop_bracing_conn_start_edge As Double
    Private prop_bracing_conn_start_pitch As Double
    Private prop_bracing_conn_start_gage As Double
    Private prop_bracing_conn_start_gage_space As Double
    Private prop_bracing_conn_end_bolt_pattern As String
    Private prop_bracing_conn_end_bolt_size As Integer
    Private prop_bracing_conn_end_bolt_mat_prop As Integer
    Private prop_bracing_conn_end_bolt_thread As String
    Private prop_bracing_conn_end_edge As Double
    Private prop_bracing_conn_end_pitch As Double
    Private prop_bracing_conn_end_gage As Double
    Private prop_bracing_conn_end_gage_space As Double
    Private prop_bracing_dp_connector_type As String
    Private prop_bracing_dp_connector_spacing As Double
    Private prop_bracing_dp_end_connection As Boolean
    Private prop_bracing_dp_fully_comp As Boolean
    Private prop_bracing_dp_crushing As Boolean
    Private prop_bracing_dp_tension As Boolean
    Private prop_bracing_dp_eccentricity As Boolean
    Private prop_bracing_dp_U As Double
    Private prop_bracing_dp_Lex As Double
    Private prop_bracing_dp_Ley As Double
    Private prop_bracing_dp_Lez As Double
    Private prop_bracing_dp_Kx As Double
    Private prop_bracing_dp_Ky As Double
    Private prop_bracing_dp_Kz As Double

    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Id")>
    Public Property bracing_id() As Integer
        Get
            Return Me.prop_bracing_id
        End Get
        Set
            Me.prop_bracing_id = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Pattern_Type")>
    Public Property bracing_pattern_type() As String
        Get
            Return Me.prop_bracing_pattern_type
        End Get
        Set
            Me.prop_bracing_pattern_type = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Type")>
    Public Property bracing_type() As String
        Get
            Return Me.prop_bracing_type
        End Get
        Set
            Me.prop_bracing_type = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Sect_Type")>
    Public Property bracing_sect_type() As String
        Get
            Return Me.prop_bracing_sect_type
        End Get
        Set
            Me.prop_bracing_sect_type = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Sect_Prop")>
    Public Property bracing_sect_prop() As Integer
        Get
            Return Me.prop_bracing_sect_prop
        End Get
        Set
            Me.prop_bracing_sect_prop = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Mat_Prop")>
    Public Property bracing_mat_prop() As Integer
        Get
            Return Me.prop_bracing_mat_prop
        End Get
        Set
            Me.prop_bracing_mat_prop = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Cond")>
    Public Property bracing_conn_end_cond() As String
        Get
            Return Me.prop_bracing_conn_end_cond
        End Get
        Set
            Me.prop_bracing_conn_end_cond = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Mirror")>
    Public Property bracing_conn_mirror() As Boolean
        Get
            Return Me.prop_bracing_conn_mirror
        End Get
        Set
            Me.prop_bracing_conn_mirror = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Pattern")>
    Public Property bracing_conn_start_pattern() As String
        Get
            Return Me.prop_bracing_conn_start_pattern
        End Get
        Set
            Me.prop_bracing_conn_start_pattern = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Bolt_Size")>
    Public Property bracing_conn_start_bolt_size() As Integer
        Get
            Return Me.prop_bracing_conn_start_bolt_size
        End Get
        Set
            Me.prop_bracing_conn_start_bolt_size = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Bolt_Mat_Prop")>
    Public Property bracing_conn_start_bolt_mat_prop() As Integer
        Get
            Return Me.prop_bracing_conn_start_bolt_mat_prop
        End Get
        Set
            Me.prop_bracing_conn_start_bolt_mat_prop = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Bolt_Thread")>
    Public Property bracing_conn_start_bolt_thread() As String
        Get
            Return Me.prop_bracing_conn_start_bolt_thread
        End Get
        Set
            Me.prop_bracing_conn_start_bolt_thread = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Edge")>
    Public Property bracing_conn_start_edge() As Double
        Get
            Return Me.prop_bracing_conn_start_edge
        End Get
        Set
            Me.prop_bracing_conn_start_edge = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Pitch")>
    Public Property bracing_conn_start_pitch() As Double
        Get
            Return Me.prop_bracing_conn_start_pitch
        End Get
        Set
            Me.prop_bracing_conn_start_pitch = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Gage")>
    Public Property bracing_conn_start_gage() As Double
        Get
            Return Me.prop_bracing_conn_start_gage
        End Get
        Set
            Me.prop_bracing_conn_start_gage = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Gage_Space")>
    Public Property bracing_conn_start_gage_space() As Double
        Get
            Return Me.prop_bracing_conn_start_gage_space
        End Get
        Set
            Me.prop_bracing_conn_start_gage_space = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Bolt_Pattern")>
    Public Property bracing_conn_end_bolt_pattern() As String
        Get
            Return Me.prop_bracing_conn_end_bolt_pattern
        End Get
        Set
            Me.prop_bracing_conn_end_bolt_pattern = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Bolt_Size")>
    Public Property bracing_conn_end_bolt_size() As Integer
        Get
            Return Me.prop_bracing_conn_end_bolt_size
        End Get
        Set
            Me.prop_bracing_conn_end_bolt_size = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Bolt_Mat_Prop")>
    Public Property bracing_conn_end_bolt_mat_prop() As Integer
        Get
            Return Me.prop_bracing_conn_end_bolt_mat_prop
        End Get
        Set
            Me.prop_bracing_conn_end_bolt_mat_prop = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Bolt_Thread")>
    Public Property bracing_conn_end_bolt_thread() As String
        Get
            Return Me.prop_bracing_conn_end_bolt_thread
        End Get
        Set
            Me.prop_bracing_conn_end_bolt_thread = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Edge")>
    Public Property bracing_conn_end_edge() As Double
        Get
            Return Me.prop_bracing_conn_end_edge
        End Get
        Set
            Me.prop_bracing_conn_end_edge = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Pitch")>
    Public Property bracing_conn_end_pitch() As Double
        Get
            Return Me.prop_bracing_conn_end_pitch
        End Get
        Set
            Me.prop_bracing_conn_end_pitch = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Gage")>
    Public Property bracing_conn_end_gage() As Double
        Get
            Return Me.prop_bracing_conn_end_gage
        End Get
        Set
            Me.prop_bracing_conn_end_gage = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Gage_Space")>
    Public Property bracing_conn_end_gage_space() As Double
        Get
            Return Me.prop_bracing_conn_end_gage_space
        End Get
        Set
            Me.prop_bracing_conn_end_gage_space = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Connector_Type")>
    Public Property bracing_dp_connector_type() As String
        Get
            Return Me.prop_bracing_dp_connector_type
        End Get
        Set
            Me.prop_bracing_dp_connector_type = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Connector_Spacing")>
    Public Property bracing_dp_connector_spacing() As Double
        Get
            Return Me.prop_bracing_dp_connector_spacing
        End Get
        Set
            Me.prop_bracing_dp_connector_spacing = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_End_Connection")>
    Public Property bracing_dp_end_connection() As Boolean
        Get
            Return Me.prop_bracing_dp_end_connection
        End Get
        Set
            Me.prop_bracing_dp_end_connection = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Fully_Comp")>
    Public Property bracing_dp_fully_comp() As Boolean
        Get
            Return Me.prop_bracing_dp_fully_comp
        End Get
        Set
            Me.prop_bracing_dp_fully_comp = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Crushing")>
    Public Property bracing_dp_crushing() As Boolean
        Get
            Return Me.prop_bracing_dp_crushing
        End Get
        Set
            Me.prop_bracing_dp_crushing = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Tension")>
    Public Property bracing_dp_tension() As Boolean
        Get
            Return Me.prop_bracing_dp_tension
        End Get
        Set
            Me.prop_bracing_dp_tension = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Eccentricity")>
    Public Property bracing_dp_eccentricity() As Boolean
        Get
            Return Me.prop_bracing_dp_eccentricity
        End Get
        Set
            Me.prop_bracing_dp_eccentricity = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_U")>
    Public Property bracing_dp_U() As Double
        Get
            Return Me.prop_bracing_dp_U
        End Get
        Set
            Me.prop_bracing_dp_U = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Lex")>
    Public Property bracing_dp_Lex() As Double
        Get
            Return Me.prop_bracing_dp_Lex
        End Get
        Set
            Me.prop_bracing_dp_Lex = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Ley")>
    Public Property bracing_dp_Ley() As Double
        Get
            Return Me.prop_bracing_dp_Ley
        End Get
        Set
            Me.prop_bracing_dp_Ley = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Lez")>
    Public Property bracing_dp_Lez() As Double
        Get
            Return Me.prop_bracing_dp_Lez
        End Get
        Set
            Me.prop_bracing_dp_Lez = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Kx")>
    Public Property bracing_dp_Kx() As Double
        Get
            Return Me.prop_bracing_dp_Kx
        End Get
        Set
            Me.prop_bracing_dp_Kx = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Ky")>
    Public Property bracing_dp_Ky() As Double
        Get
            Return Me.prop_bracing_dp_Ky
        End Get
        Set
            Me.prop_bracing_dp_Ky = Value
        End Set
    End Property
    <Category("Lattice Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Kz")>
    Public Property bracing_dp_Kz() As Double
        Get
            Return Me.prop_bracing_dp_Kz
        End Get
        Set
            Me.prop_bracing_dp_Kz = Value
        End Set
    End Property

End Class
Partial Public Class lattice_leg_detail
    Private prop_section_id As Integer
    Private prop_ID As Integer
    Private prop_leg_sect_type As String
    Private prop_leg_sect_prop As Integer
    Private prop_leg_mat_prop As Integer
    Private prop_leg_dp_connector_type As String
    Private prop_leg_dp_connector_spacing As Double
    Private prop_leg_dp_end_connection As Boolean
    Private prop_leg_dp_fully_comp As Boolean
    Private prop_leg_dp_crushing As Boolean
    Private prop_leg_dp_tension As Boolean
    Private prop_leg_dp_eccentricity As Boolean
    Private prop_leg_dp_U As Double
    Private prop_leg_dp_Lex As Double
    Private prop_leg_dp_Ley As Double
    Private prop_leg_dp_Lez As Double
    Private prop_leg_dp_Kx As Double
    Private prop_leg_dp_Ky As Double
    Private prop_leg_dp_Kz As Double
    Private prop_leg_conn_type As String
    Private prop_leg_conn_flng_bolt_qty As Integer
    Private prop_leg_conn_flng_bolt_size As Integer
    Private prop_leg_conn_flng_bolt_mat_name As Integer
    Private prop_leg_conn_shear_slv_bolt_qty As Integer
    Private prop_leg_conn_shear_slv_bolt_size As Integer
    Private prop_leg_conn_shear_slv_bolt_mat_name As Integer
    Private prop_leg_conn_shear_slv_resist_comp As Boolean
    Private prop_leg_conn_shear_slv_type As String
    Private prop_leg_conn_shear_slv_edge_dist As Double

    <Category("Lattice Leg Detail"), Description(""), DisplayName("Section_Id")>
    Public Property section_id() As Integer
        Get
            Return Me.prop_section_id
        End Get
        Set
            Me.prop_section_id = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Sect_Type")>
    Public Property leg_sect_type() As String
        Get
            Return Me.prop_leg_sect_type
        End Get
        Set
            Me.prop_leg_sect_type = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Sect_Prop")>
    Public Property leg_sect_prop() As Integer
        Get
            Return Me.prop_leg_sect_prop
        End Get
        Set
            Me.prop_leg_sect_prop = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Mat_Prop")>
    Public Property leg_mat_prop() As Integer
        Get
            Return Me.prop_leg_mat_prop
        End Get
        Set
            Me.prop_leg_mat_prop = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Connector_Type")>
    Public Property leg_dp_connector_type() As String
        Get
            Return Me.prop_leg_dp_connector_type
        End Get
        Set
            Me.prop_leg_dp_connector_type = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Connector_Spacing")>
    Public Property leg_dp_connector_spacing() As Double
        Get
            Return Me.prop_leg_dp_connector_spacing
        End Get
        Set
            Me.prop_leg_dp_connector_spacing = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_End_Connection")>
    Public Property leg_dp_end_connection() As Boolean
        Get
            Return Me.prop_leg_dp_end_connection
        End Get
        Set
            Me.prop_leg_dp_end_connection = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Fully_Comp")>
    Public Property leg_dp_fully_comp() As Boolean
        Get
            Return Me.prop_leg_dp_fully_comp
        End Get
        Set
            Me.prop_leg_dp_fully_comp = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Crushing")>
    Public Property leg_dp_crushing() As Boolean
        Get
            Return Me.prop_leg_dp_crushing
        End Get
        Set
            Me.prop_leg_dp_crushing = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Tension")>
    Public Property leg_dp_tension() As Boolean
        Get
            Return Me.prop_leg_dp_tension
        End Get
        Set
            Me.prop_leg_dp_tension = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Eccentricity")>
    Public Property leg_dp_eccentricity() As Boolean
        Get
            Return Me.prop_leg_dp_eccentricity
        End Get
        Set
            Me.prop_leg_dp_eccentricity = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_U")>
    Public Property leg_dp_U() As Double
        Get
            Return Me.prop_leg_dp_U
        End Get
        Set
            Me.prop_leg_dp_U = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Lex")>
    Public Property leg_dp_Lex() As Double
        Get
            Return Me.prop_leg_dp_Lex
        End Get
        Set
            Me.prop_leg_dp_Lex = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Ley")>
    Public Property leg_dp_Ley() As Double
        Get
            Return Me.prop_leg_dp_Ley
        End Get
        Set
            Me.prop_leg_dp_Ley = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Lez")>
    Public Property leg_dp_Lez() As Double
        Get
            Return Me.prop_leg_dp_Lez
        End Get
        Set
            Me.prop_leg_dp_Lez = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Kx")>
    Public Property leg_dp_Kx() As Double
        Get
            Return Me.prop_leg_dp_Kx
        End Get
        Set
            Me.prop_leg_dp_Kx = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Ky")>
    Public Property leg_dp_Ky() As Double
        Get
            Return Me.prop_leg_dp_Ky
        End Get
        Set
            Me.prop_leg_dp_Ky = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Dp_Kz")>
    Public Property leg_dp_Kz() As Double
        Get
            Return Me.prop_leg_dp_Kz
        End Get
        Set
            Me.prop_leg_dp_Kz = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Conn_Type")>
    Public Property leg_conn_type() As String
        Get
            Return Me.prop_leg_conn_type
        End Get
        Set
            Me.prop_leg_conn_type = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Conn_Flng_Bolt_Qty")>
    Public Property leg_conn_flng_bolt_qty() As Integer
        Get
            Return Me.prop_leg_conn_flng_bolt_qty
        End Get
        Set
            Me.prop_leg_conn_flng_bolt_qty = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Conn_Flng_Bolt_Size")>
    Public Property leg_conn_flng_bolt_size() As Integer
        Get
            Return Me.prop_leg_conn_flng_bolt_size
        End Get
        Set
            Me.prop_leg_conn_flng_bolt_size = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Conn_Flng_Bolt_Mat_Name")>
    Public Property leg_conn_flng_bolt_mat_name() As Integer
        Get
            Return Me.prop_leg_conn_flng_bolt_mat_name
        End Get
        Set
            Me.prop_leg_conn_flng_bolt_mat_name = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Conn_Shear_Slv_Bolt_Qty")>
    Public Property leg_conn_shear_slv_bolt_qty() As Integer
        Get
            Return Me.prop_leg_conn_shear_slv_bolt_qty
        End Get
        Set
            Me.prop_leg_conn_shear_slv_bolt_qty = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Conn_Shear_Slv_Bolt_Size")>
    Public Property leg_conn_shear_slv_bolt_size() As Integer
        Get
            Return Me.prop_leg_conn_shear_slv_bolt_size
        End Get
        Set
            Me.prop_leg_conn_shear_slv_bolt_size = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Conn_Shear_Slv_Bolt_Mat_Name")>
    Public Property leg_conn_shear_slv_bolt_mat_name() As Integer
        Get
            Return Me.prop_leg_conn_shear_slv_bolt_mat_name
        End Get
        Set
            Me.prop_leg_conn_shear_slv_bolt_mat_name = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Conn_Shear_Slv_Resist_Comp")>
    Public Property leg_conn_shear_slv_resist_comp() As Boolean
        Get
            Return Me.prop_leg_conn_shear_slv_resist_comp
        End Get
        Set
            Me.prop_leg_conn_shear_slv_resist_comp = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Conn_Shear_Slv_Type")>
    Public Property leg_conn_shear_slv_type() As String
        Get
            Return Me.prop_leg_conn_shear_slv_type
        End Get
        Set
            Me.prop_leg_conn_shear_slv_type = Value
        End Set
    End Property
    <Category("Lattice Leg Detail"), Description(""), DisplayName("Leg_Conn_Shear_Slv_Edge_Dist")>
    Public Property leg_conn_shear_slv_edge_dist() As Double
        Get
            Return Me.prop_leg_conn_shear_slv_edge_dist
        End Get
        Set
            Me.prop_leg_conn_shear_slv_edge_dist = Value
        End Set
    End Property

End Class
Partial Public Class lattice_custom_capacity
    Private prop_member_detail_id As Integer
    Private prop_ID As Integer
    Private prop_cc_member_type As String
    Private prop_cc_code As String
    Private prop_cc_comp As Double
    Private prop_cc_ten As Double
    Private prop_cc_conn_comp As Double
    Private prop_cc_conn_ten As Double
    Private prop_cc_pass_rating As Decimal

    <Category("Lattice Custom Capacity"), Description(""), DisplayName("Member_Detail_Id")>
    Public Property member_detail_id() As Integer
        Get
            Return Me.prop_member_detail_id
        End Get
        Set
            Me.prop_member_detail_id = Value
        End Set
    End Property
    <Category("Lattice Custom Capacity"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Lattice Custom Capacity"), Description(""), DisplayName("Cc_Member_Type")>
    Public Property cc_member_type() As String
        Get
            Return Me.prop_cc_member_type
        End Get
        Set
            Me.prop_cc_member_type = Value
        End Set
    End Property
    <Category("Lattice Custom Capacity"), Description(""), DisplayName("Cc_Code")>
    Public Property cc_code() As String
        Get
            Return Me.prop_cc_code
        End Get
        Set
            Me.prop_cc_code = Value
        End Set
    End Property
    <Category("Lattice Custom Capacity"), Description(""), DisplayName("Cc_Comp")>
    Public Property cc_comp() As Double
        Get
            Return Me.prop_cc_comp
        End Get
        Set
            Me.prop_cc_comp = Value
        End Set
    End Property
    <Category("Lattice Custom Capacity"), Description(""), DisplayName("Cc_Ten")>
    Public Property cc_ten() As Double
        Get
            Return Me.prop_cc_ten
        End Get
        Set
            Me.prop_cc_ten = Value
        End Set
    End Property
    <Category("Lattice Custom Capacity"), Description(""), DisplayName("Cc_Conn_Comp")>
    Public Property cc_conn_comp() As Double
        Get
            Return Me.prop_cc_conn_comp
        End Get
        Set
            Me.prop_cc_conn_comp = Value
        End Set
    End Property
    <Category("Lattice Custom Capacity"), Description(""), DisplayName("Cc_Conn_Ten")>
    Public Property cc_conn_ten() As Double
        Get
            Return Me.prop_cc_conn_ten
        End Get
        Set
            Me.prop_cc_conn_ten = Value
        End Set
    End Property
    <Category("Lattice Custom Capacity"), Description(""), DisplayName("Cc_Pass_Rating")>
    Public Property cc_pass_rating() As Decimal
        Get
            Return Me.prop_cc_pass_rating
        End Get
        Set
            Me.prop_cc_pass_rating = Value
        End Set
    End Property

End Class
#End Region

#Region "Pole"
Partial Public Class pole_section
    Private prop_model_id As Integer
    Private prop_ID As Integer
    Private prop_analysis_section_id As Integer
    Private prop_elev_bot As Double
    Private prop_elev_top As Double
    Private prop_length_section As Double
    Private prop_length_splice As Double
    Private prop_num_sides As Integer
    Private prop_diam_bot As Double
    Private prop_diam_top As Double
    Private prop_wall_thickness As Double
    Private prop_bend_radius As Double
    Private prop_steel_grade As String
    Private prop_pole_type As String
    Private prop_section_name As String
    Private prop_socket_length As Double
    Private prop_weight_mult As Double
    Private prop_wp_mult As Double
    Private prop_af_factor As Double
    Private prop_ar_factor As Double
    Private prop_round_area_ratio As Double
    Private prop_flat_area_ratio As Double

    <Category("Pole Section"), Description(""), DisplayName("Model_Id")>
    Public Property model_id() As Integer
        Get
            Return Me.prop_model_id
        End Get
        Set
            Me.prop_model_id = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Analysis_Section_Id")>
    Public Property analysis_section_id() As Integer
        Get
            Return Me.prop_analysis_section_id
        End Get
        Set
            Me.prop_analysis_section_id = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Elev_Bot")>
    Public Property elev_bot() As Double
        Get
            Return Me.prop_elev_bot
        End Get
        Set
            Me.prop_elev_bot = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Elev_Top")>
    Public Property elev_top() As Double
        Get
            Return Me.prop_elev_top
        End Get
        Set
            Me.prop_elev_top = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Length_Section")>
    Public Property length_section() As Double
        Get
            Return Me.prop_length_section
        End Get
        Set
            Me.prop_length_section = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Length_Splice")>
    Public Property length_splice() As Double
        Get
            Return Me.prop_length_splice
        End Get
        Set
            Me.prop_length_splice = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Num_Sides")>
    Public Property num_sides() As Integer
        Get
            Return Me.prop_num_sides
        End Get
        Set
            Me.prop_num_sides = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Diam_Bot")>
    Public Property diam_bot() As Double
        Get
            Return Me.prop_diam_bot
        End Get
        Set
            Me.prop_diam_bot = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Diam_Top")>
    Public Property diam_top() As Double
        Get
            Return Me.prop_diam_top
        End Get
        Set
            Me.prop_diam_top = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Wall_Thickness")>
    Public Property wall_thickness() As Double
        Get
            Return Me.prop_wall_thickness
        End Get
        Set
            Me.prop_wall_thickness = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Bend_Radius")>
    Public Property bend_radius() As Double
        Get
            Return Me.prop_bend_radius
        End Get
        Set
            Me.prop_bend_radius = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Steel_Grade")>
    Public Property steel_grade() As String
        Get
            Return Me.prop_steel_grade
        End Get
        Set
            Me.prop_steel_grade = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Pole_Type")>
    Public Property pole_type() As String
        Get
            Return Me.prop_pole_type
        End Get
        Set
            Me.prop_pole_type = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Section_Name")>
    Public Property section_name() As String
        Get
            Return Me.prop_section_name
        End Get
        Set
            Me.prop_section_name = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Socket_Length")>
    Public Property socket_length() As Double
        Get
            Return Me.prop_socket_length
        End Get
        Set
            Me.prop_socket_length = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Weight_Mult")>
    Public Property weight_mult() As Double
        Get
            Return Me.prop_weight_mult
        End Get
        Set
            Me.prop_weight_mult = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Wp_Mult")>
    Public Property wp_mult() As Double
        Get
            Return Me.prop_wp_mult
        End Get
        Set
            Me.prop_wp_mult = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Af_Factor")>
    Public Property af_factor() As Double
        Get
            Return Me.prop_af_factor
        End Get
        Set
            Me.prop_af_factor = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Ar_Factor")>
    Public Property ar_factor() As Double
        Get
            Return Me.prop_ar_factor
        End Get
        Set
            Me.prop_ar_factor = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Round_Area_Ratio")>
    Public Property round_area_ratio() As Double
        Get
            Return Me.prop_round_area_ratio
        End Get
        Set
            Me.prop_round_area_ratio = Value
        End Set
    End Property
    <Category("Pole Section"), Description(""), DisplayName("Flat_Area_Ratio")>
    Public Property flat_area_ratio() As Double
        Get
            Return Me.prop_flat_area_ratio
        End Get
        Set
            Me.prop_flat_area_ratio = Value
        End Set
    End Property

End Class
Partial Public Class pole_reinf_section
    Private prop_model_id As Integer
    Private prop_ID As Integer
    Private prop_analysis_section_ID As Integer
    Private prop_elev_bot As Double
    Private prop_elev_top As Double
    Private prop_length_section As Double
    Private prop_length_splice As Double
    Private prop_num_sides As Integer
    Private prop_diam_bot As Double
    Private prop_diam_top As Double
    Private prop_wall_thickness As Double
    Private prop_bend_radius As Double
    Private prop_steel_grade As String
    Private prop_weight_mult As Double
    Private prop_section_name As String
    Private prop_wp_mult As Double
    Private prop_af_factor As Double
    Private prop_ar_factor As Double
    Private prop_round_area_ratio As Double
    Private prop_flat_area_ratio As Double

    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Model_Id")>
    Public Property model_id() As Integer
        Get
            Return Me.prop_model_id
        End Get
        Set
            Me.prop_model_id = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Analysis_Section_Id")>
    Public Property analysis_section_ID() As Integer
        Get
            Return Me.prop_analysis_section_ID
        End Get
        Set
            Me.prop_analysis_section_ID = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Elev_Bot")>
    Public Property elev_bot() As Double
        Get
            Return Me.prop_elev_bot
        End Get
        Set
            Me.prop_elev_bot = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Elev_Top")>
    Public Property elev_top() As Double
        Get
            Return Me.prop_elev_top
        End Get
        Set
            Me.prop_elev_top = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Length_Section")>
    Public Property length_section() As Double
        Get
            Return Me.prop_length_section
        End Get
        Set
            Me.prop_length_section = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Length_Splice")>
    Public Property length_splice() As Double
        Get
            Return Me.prop_length_splice
        End Get
        Set
            Me.prop_length_splice = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Num_Sides")>
    Public Property num_sides() As Integer
        Get
            Return Me.prop_num_sides
        End Get
        Set
            Me.prop_num_sides = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Diam_Bot")>
    Public Property diam_bot() As Double
        Get
            Return Me.prop_diam_bot
        End Get
        Set
            Me.prop_diam_bot = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Diam_Top")>
    Public Property diam_top() As Double
        Get
            Return Me.prop_diam_top
        End Get
        Set
            Me.prop_diam_top = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Wall_Thickness")>
    Public Property wall_thickness() As Double
        Get
            Return Me.prop_wall_thickness
        End Get
        Set
            Me.prop_wall_thickness = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Bend_Radius")>
    Public Property bend_radius() As Double
        Get
            Return Me.prop_bend_radius
        End Get
        Set
            Me.prop_bend_radius = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Steel_Grade")>
    Public Property steel_grade() As String
        Get
            Return Me.prop_steel_grade
        End Get
        Set
            Me.prop_steel_grade = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Weight_Mult")>
    Public Property weight_mult() As Double
        Get
            Return Me.prop_weight_mult
        End Get
        Set
            Me.prop_weight_mult = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Section_Name")>
    Public Property section_name() As String
        Get
            Return Me.prop_section_name
        End Get
        Set
            Me.prop_section_name = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Wp_Mult")>
    Public Property wp_mult() As Double
        Get
            Return Me.prop_wp_mult
        End Get
        Set
            Me.prop_wp_mult = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Af_Factor")>
    Public Property af_factor() As Double
        Get
            Return Me.prop_af_factor
        End Get
        Set
            Me.prop_af_factor = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Ar_Factor")>
    Public Property ar_factor() As Double
        Get
            Return Me.prop_ar_factor
        End Get
        Set
            Me.prop_ar_factor = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Round_Area_Ratio")>
    Public Property round_area_ratio() As Double
        Get
            Return Me.prop_round_area_ratio
        End Get
        Set
            Me.prop_round_area_ratio = Value
        End Set
    End Property
    <Category("Pole Reinforcement Section"), Description(""), DisplayName("Flat_Area_Ratio")>
    Public Property flat_area_ratio() As Double
        Get
            Return Me.prop_flat_area_ratio
        End Get
        Set
            Me.prop_flat_area_ratio = Value
        End Set
    End Property

End Class
Partial Public Class pole_reinf_group
    Private prop_model_id As Integer
    Private prop_ID As Integer
    Private prop_elev_bot_actual As Double
    Private prop_elev_bot_eff As Double
    Private prop_elev_top_actual As Double
    Private prop_elev_top_eff As Double
    Private prop_reinf_db_id As Integer
    Private prop_bolt_db_id_top As Integer
    Private prop_bolt_db_id_bot As Integer

    <Category("Pole Reinforcement Group"), Description(""), DisplayName("Model_Id")>
    Public Property model_id() As Integer
        Get
            Return Me.prop_model_id
        End Get
        Set
            Me.prop_model_id = Value
        End Set
    End Property
    <Category("Pole Reinforcement Group"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Pole Reinforcement Group"), Description(""), DisplayName("Elev_Bot_Actual")>
    Public Property elev_bot_actual() As Double
        Get
            Return Me.prop_elev_bot_actual
        End Get
        Set
            Me.prop_elev_bot_actual = Value
        End Set
    End Property
    <Category("Pole Reinforcement Group"), Description(""), DisplayName("Elev_Bot_Eff")>
    Public Property elev_bot_eff() As Double
        Get
            Return Me.prop_elev_bot_eff
        End Get
        Set
            Me.prop_elev_bot_eff = Value
        End Set
    End Property
    <Category("Pole Reinforcement Group"), Description(""), DisplayName("Elev_Top_Actual")>
    Public Property elev_top_actual() As Double
        Get
            Return Me.prop_elev_top_actual
        End Get
        Set
            Me.prop_elev_top_actual = Value
        End Set
    End Property
    <Category("Pole Reinforcement Group"), Description(""), DisplayName("Elev_Top_Eff")>
    Public Property elev_top_eff() As Double
        Get
            Return Me.prop_elev_top_eff
        End Get
        Set
            Me.prop_elev_top_eff = Value
        End Set
    End Property
    <Category("Pole Reinforcement Group"), Description(""), DisplayName("Reinf_Db_Id")>
    Public Property reinf_db_id() As Integer
        Get
            Return Me.prop_reinf_db_id
        End Get
        Set
            Me.prop_reinf_db_id = Value
        End Set
    End Property
    <Category("Pole Reinforcement Group"), Description(""), DisplayName("Bolt_Db_Id_Top")>
    Public Property bolt_db_id_top() As Integer
        Get
            Return Me.prop_bolt_db_id_top
        End Get
        Set
            Me.prop_bolt_db_id_top = Value
        End Set
    End Property
    <Category("Pole Reinforcement Group"), Description(""), DisplayName("Bolt_Db_Id_Bot")>
    Public Property bolt_db_id_bot() As Integer
        Get
            Return Me.prop_bolt_db_id_bot
        End Get
        Set
            Me.prop_bolt_db_id_bot = Value
        End Set
    End Property

End Class
Partial Public Class pole_reinf_details
    Private prop_reinf_group_id As Integer
    Private prop_ID As Integer
    Private prop_pole_flat As Integer
    Private prop_horizontal_offset As Double
    Private prop_rotation As Double

    <Category("Pole Reinforcement Details"), Description(""), DisplayName("Reinf_Group_Id")>
    Public Property reinf_group_id() As Integer
        Get
            Return Me.prop_reinf_group_id
        End Get
        Set
            Me.prop_reinf_group_id = Value
        End Set
    End Property
    <Category("Pole Reinforcement Details"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Pole Reinforcement Details"), Description(""), DisplayName("Pole_Flat")>
    Public Property pole_flat() As Integer
        Get
            Return Me.prop_pole_flat
        End Get
        Set
            Me.prop_pole_flat = Value
        End Set
    End Property
    <Category("Pole Reinforcement Details"), Description(""), DisplayName("Horizontal_Offset")>
    Public Property horizontal_offset() As Double
        Get
            Return Me.prop_horizontal_offset
        End Get
        Set
            Me.prop_horizontal_offset = Value
        End Set
    End Property
    <Category("Pole Reinforcement Details"), Description(""), DisplayName("Rotation")>
    Public Property rotation() As Double
        Get
            Return Me.prop_rotation
        End Get
        Set
            Me.prop_rotation = Value
        End Set
    End Property

End Class
Partial Public Class pole_interference_group
    Private prop_model_id As Integer
    Private prop_ID As Integer
    Private prop_elev_bot As Double
    Private prop_elev_top As Double
    Private prop_width As Double
    Private prop_description As String

    <Category("Pole Interference Group"), Description(""), DisplayName("Model_Id")>
    Public Property model_id() As Integer
        Get
            Return Me.prop_model_id
        End Get
        Set
            Me.prop_model_id = Value
        End Set
    End Property
    <Category("Pole Interference Group"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Pole Interference Group"), Description(""), DisplayName("Elev_Bot")>
    Public Property elev_bot() As Double
        Get
            Return Me.prop_elev_bot
        End Get
        Set
            Me.prop_elev_bot = Value
        End Set
    End Property
    <Category("Pole Interference Group"), Description(""), DisplayName("Elev_Top")>
    Public Property elev_top() As Double
        Get
            Return Me.prop_elev_top
        End Get
        Set
            Me.prop_elev_top = Value
        End Set
    End Property
    <Category("Pole Interference Group"), Description(""), DisplayName("Width")>
    Public Property width() As Double
        Get
            Return Me.prop_width
        End Get
        Set
            Me.prop_width = Value
        End Set
    End Property
    <Category("Pole Interference Group"), Description(""), DisplayName("Description")>
    Public Property description() As String
        Get
            Return Me.prop_description
        End Get
        Set
            Me.prop_description = Value
        End Set
    End Property

End Class
Partial Public Class pole_interference_detail
    Private prop_interference_group_id As Integer
    Private prop_ID As Integer
    Private prop_pole_flat As Integer
    Private prop_horizontal_offset As Double
    Private prop_rotation As Double

    <Category("Pole Interference Detail"), Description(""), DisplayName("Interference_Group_Id")>
    Public Property interference_group_id() As Integer
        Get
            Return Me.prop_interference_group_id
        End Get
        Set
            Me.prop_interference_group_id = Value
        End Set
    End Property
    <Category("Pole Interference Detail"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Pole Interference Detail"), Description(""), DisplayName("Pole_Flat")>
    Public Property pole_flat() As Integer
        Get
            Return Me.prop_pole_flat
        End Get
        Set
            Me.prop_pole_flat = Value
        End Set
    End Property
    <Category("Pole Interference Detail"), Description(""), DisplayName("Horizontal_Offset")>
    Public Property horizontal_offset() As Double
        Get
            Return Me.prop_horizontal_offset
        End Get
        Set
            Me.prop_horizontal_offset = Value
        End Set
    End Property
    <Category("Pole Interference Detail"), Description(""), DisplayName("Rotation")>
    Public Property rotation() As Double
        Get
            Return Me.prop_rotation
        End Get
        Set
            Me.prop_rotation = Value
        End Set
    End Property

End Class
Partial Public Class memb_prop_flat_plate
    Private prop_ID As Integer
    Private prop_name As String
    Private prop_type As String
    Private prop_b As Double
    Private prop_h As Double
    Private prop_sr_diam As Double
    Private prop_channel_thkns_web As Double
    Private prop_channel_thkns_flange As Double
    Private prop_channel_eo As Double
    Private prop_channel_J As Double
    Private prop_channel_Cw As Double
    Private prop_area_gross As Double
    Private prop_centroid As Double
    Private prop_istension As Boolean
    Private prop_matl As String
    Private prop_Ix As Double
    Private prop_Iy As Double
    Private prop_Lu As Double
    Private prop_Kx As Double
    Private prop_Ky As Double
    Private prop_bolt_hole_size As Double
    Private prop_area_net As Double
    Private prop_shear_lag As Double
    Private prop_connection_type_bot As String
    Private prop_connection_cap_revF_bot As Double
    Private prop_connection_cap_revG_bot As Double
    Private prop_connection_cap_revH_bot As Double
    Private prop_bolt_id_bot As Integer
    Private prop_bolt_N_or_X_bot As String
    Private prop_bolt_num_bot As Integer
    Private prop_bolt_spacing_bot As Double
    Private prop_bolt_edge_dist_bot As Double
    Private prop_FlangeOrBP_connected_bot As Boolean
    Private prop_weld_grade_bot As Double
    Private prop_weld_trans_type_bot As String
    Private prop_weld_trans_length_bot As Double
    Private prop_weld_groove_depth_bot As Double
    Private prop_weld_groove_angle_bot As Integer
    Private prop_weld_trans_fillet_size_bot As Double
    Private prop_weld_trans_eff_throat_bot As Double
    Private prop_weld_long_type_bot As String
    Private prop_weld_long_length_bot As Double
    Private prop_weld_long_fillet_size_bot As Double
    Private prop_weld_long_eff_throat_bot As Double
    Private prop_top_bot_connections_symmetrical As Boolean
    Private prop_connection_type_top As String
    Private prop_connection_cap_revF_top As Double
    Private prop_connection_cap_revG_top As Double
    Private prop_connection_cap_revH_top As Double
    Private prop_bolt_id_top As Integer
    Private prop_bolt_N_or_X_top As String
    Private prop_bolt_num_top As Integer
    Private prop_bolt_spacing_top As Double
    Private prop_bolt_edge_dist_top As Double
    Private prop_FlangeOrBP_connected_top As Boolean
    Private prop_weld_grade_top As Double
    Private prop_weld_trans_type_top As String
    Private prop_weld_trans_length_top As Double
    Private prop_weld_groove_depth_top As Double
    Private prop_weld_groove_angle_top As Integer
    Private prop_weld_trans_fillet_size_top As Double
    Private prop_weld_trans_eff_throat_top As Double
    Private prop_weld_long_type_top As String
    Private prop_weld_long_length_top As Double
    Private prop_weld_long_fillet_size_top As Double
    Private prop_weld_long_eff_throat_top As Double
    Private prop_conn_length_bot As Double
    Private prop_conn_length_top As Double

    <Category("Member Property Flat Plate"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Name")>
    Public Property name() As String
        Get
            Return Me.prop_name
        End Get
        Set
            Me.prop_name = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Type")>
    Public Property type() As String
        Get
            Return Me.prop_type
        End Get
        Set
            Me.prop_type = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("B")>
    Public Property b() As Double
        Get
            Return Me.prop_b
        End Get
        Set
            Me.prop_b = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("H")>
    Public Property h() As Double
        Get
            Return Me.prop_h
        End Get
        Set
            Me.prop_h = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Sr_Diam")>
    Public Property sr_diam() As Double
        Get
            Return Me.prop_sr_diam
        End Get
        Set
            Me.prop_sr_diam = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Channel_Thkns_Web")>
    Public Property channel_thkns_web() As Double
        Get
            Return Me.prop_channel_thkns_web
        End Get
        Set
            Me.prop_channel_thkns_web = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Channel_Thkns_Flange")>
    Public Property channel_thkns_flange() As Double
        Get
            Return Me.prop_channel_thkns_flange
        End Get
        Set
            Me.prop_channel_thkns_flange = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Channel_Eo")>
    Public Property channel_eo() As Double
        Get
            Return Me.prop_channel_eo
        End Get
        Set
            Me.prop_channel_eo = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Channel_J")>
    Public Property channel_J() As Double
        Get
            Return Me.prop_channel_J
        End Get
        Set
            Me.prop_channel_J = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Channel_Cw")>
    Public Property channel_Cw() As Double
        Get
            Return Me.prop_channel_Cw
        End Get
        Set
            Me.prop_channel_Cw = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Area_Gross")>
    Public Property area_gross() As Double
        Get
            Return Me.prop_area_gross
        End Get
        Set
            Me.prop_area_gross = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Centroid")>
    Public Property centroid() As Double
        Get
            Return Me.prop_centroid
        End Get
        Set
            Me.prop_centroid = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Istension")>
    Public Property istension() As Boolean
        Get
            Return Me.prop_istension
        End Get
        Set
            Me.prop_istension = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Matl")>
    Public Property matl() As String
        Get
            Return Me.prop_matl
        End Get
        Set
            Me.prop_matl = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Ix")>
    Public Property Ix() As Double
        Get
            Return Me.prop_Ix
        End Get
        Set
            Me.prop_Ix = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Iy")>
    Public Property Iy() As Double
        Get
            Return Me.prop_Iy
        End Get
        Set
            Me.prop_Iy = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Lu")>
    Public Property Lu() As Double
        Get
            Return Me.prop_Lu
        End Get
        Set
            Me.prop_Lu = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Kx")>
    Public Property Kx() As Double
        Get
            Return Me.prop_Kx
        End Get
        Set
            Me.prop_Kx = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Ky")>
    Public Property Ky() As Double
        Get
            Return Me.prop_Ky
        End Get
        Set
            Me.prop_Ky = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Bolt_Hole_Size")>
    Public Property bolt_hole_size() As Double
        Get
            Return Me.prop_bolt_hole_size
        End Get
        Set
            Me.prop_bolt_hole_size = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Area_Net")>
    Public Property area_net() As Double
        Get
            Return Me.prop_area_net
        End Get
        Set
            Me.prop_area_net = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Shear_Lag")>
    Public Property shear_lag() As Double
        Get
            Return Me.prop_shear_lag
        End Get
        Set
            Me.prop_shear_lag = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Connection_Type_Bot")>
    Public Property connection_type_bot() As String
        Get
            Return Me.prop_connection_type_bot
        End Get
        Set
            Me.prop_connection_type_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Connection_Cap_Revf_Bot")>
    Public Property connection_cap_revF_bot() As Double
        Get
            Return Me.prop_connection_cap_revF_bot
        End Get
        Set
            Me.prop_connection_cap_revF_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Connection_Cap_Revg_Bot")>
    Public Property connection_cap_revG_bot() As Double
        Get
            Return Me.prop_connection_cap_revG_bot
        End Get
        Set
            Me.prop_connection_cap_revG_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Connection_Cap_Revh_Bot")>
    Public Property connection_cap_revH_bot() As Double
        Get
            Return Me.prop_connection_cap_revH_bot
        End Get
        Set
            Me.prop_connection_cap_revH_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("bolt_id_Bot")>
    Public Property bolt_id_bot() As Integer
        Get
            Return Me.prop_bolt_id_bot
        End Get
        Set
            Me.prop_bolt_id_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Bolt_N_Or_X_Bot")>
    Public Property bolt_N_or_X_bot() As String
        Get
            Return Me.prop_bolt_N_or_X_bot
        End Get
        Set
            Me.prop_bolt_N_or_X_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Bolt_Num_Bot")>
    Public Property bolt_num_bot() As Integer
        Get
            Return Me.prop_bolt_num_bot
        End Get
        Set
            Me.prop_bolt_num_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Bolt_Spacing_Bot")>
    Public Property bolt_spacing_bot() As Double
        Get
            Return Me.prop_bolt_spacing_bot
        End Get
        Set
            Me.prop_bolt_spacing_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Bolt_Edge_Dist_Bot")>
    Public Property bolt_edge_dist_bot() As Double
        Get
            Return Me.prop_bolt_edge_dist_bot
        End Get
        Set
            Me.prop_bolt_edge_dist_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Flangeorbp_Connected_Bot")>
    Public Property FlangeOrBP_connected_bot() As Boolean
        Get
            Return Me.prop_FlangeOrBP_connected_bot
        End Get
        Set
            Me.prop_FlangeOrBP_connected_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Grade_Bot")>
    Public Property weld_grade_bot() As Double
        Get
            Return Me.prop_weld_grade_bot
        End Get
        Set
            Me.prop_weld_grade_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Trans_Type_Bot")>
    Public Property weld_trans_type_bot() As String
        Get
            Return Me.prop_weld_trans_type_bot
        End Get
        Set
            Me.prop_weld_trans_type_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Trans_Length_Bot")>
    Public Property weld_trans_length_bot() As Double
        Get
            Return Me.prop_weld_trans_length_bot
        End Get
        Set
            Me.prop_weld_trans_length_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Groove_Depth_Bot")>
    Public Property weld_groove_depth_bot() As Double
        Get
            Return Me.prop_weld_groove_depth_bot
        End Get
        Set
            Me.prop_weld_groove_depth_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Groove_Angle_Bot")>
    Public Property weld_groove_angle_bot() As Integer
        Get
            Return Me.prop_weld_groove_angle_bot
        End Get
        Set
            Me.prop_weld_groove_angle_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Trans_Fillet_Size_Bot")>
    Public Property weld_trans_fillet_size_bot() As Double
        Get
            Return Me.prop_weld_trans_fillet_size_bot
        End Get
        Set
            Me.prop_weld_trans_fillet_size_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Trans_Eff_Throat_Bot")>
    Public Property weld_trans_eff_throat_bot() As Double
        Get
            Return Me.prop_weld_trans_eff_throat_bot
        End Get
        Set
            Me.prop_weld_trans_eff_throat_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Long_Type_Bot")>
    Public Property weld_long_type_bot() As String
        Get
            Return Me.prop_weld_long_type_bot
        End Get
        Set
            Me.prop_weld_long_type_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Long_Length_Bot")>
    Public Property weld_long_length_bot() As Double
        Get
            Return Me.prop_weld_long_length_bot
        End Get
        Set
            Me.prop_weld_long_length_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Long_Fillet_Size_Bot")>
    Public Property weld_long_fillet_size_bot() As Double
        Get
            Return Me.prop_weld_long_fillet_size_bot
        End Get
        Set
            Me.prop_weld_long_fillet_size_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Long_Eff_Throat_Bot")>
    Public Property weld_long_eff_throat_bot() As Double
        Get
            Return Me.prop_weld_long_eff_throat_bot
        End Get
        Set
            Me.prop_weld_long_eff_throat_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Top_Bot_Connections_Symmetrical")>
    Public Property top_bot_connections_symmetrical() As Boolean
        Get
            Return Me.prop_top_bot_connections_symmetrical
        End Get
        Set
            Me.prop_top_bot_connections_symmetrical = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Connection_Type_Top")>
    Public Property connection_type_top() As String
        Get
            Return Me.prop_connection_type_top
        End Get
        Set
            Me.prop_connection_type_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Connection_Cap_Revf_Top")>
    Public Property connection_cap_revF_top() As Double
        Get
            Return Me.prop_connection_cap_revF_top
        End Get
        Set
            Me.prop_connection_cap_revF_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Connection_Cap_Revg_Top")>
    Public Property connection_cap_revG_top() As Double
        Get
            Return Me.prop_connection_cap_revG_top
        End Get
        Set
            Me.prop_connection_cap_revG_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Connection_Cap_Revh_Top")>
    Public Property connection_cap_revH_top() As Double
        Get
            Return Me.prop_connection_cap_revH_top
        End Get
        Set
            Me.prop_connection_cap_revH_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("bolt_id_Top")>
    Public Property bolt_id_top() As Integer
        Get
            Return Me.prop_bolt_id_top
        End Get
        Set
            Me.prop_bolt_id_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Bolt_N_Or_X_Top")>
    Public Property bolt_N_or_X_top() As String
        Get
            Return Me.prop_bolt_N_or_X_top
        End Get
        Set
            Me.prop_bolt_N_or_X_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Bolt_Num_Top")>
    Public Property bolt_num_top() As Integer
        Get
            Return Me.prop_bolt_num_top
        End Get
        Set
            Me.prop_bolt_num_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Bolt_Spacing_Top")>
    Public Property bolt_spacing_top() As Double
        Get
            Return Me.prop_bolt_spacing_top
        End Get
        Set
            Me.prop_bolt_spacing_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Bolt_Edge_Dist_Top")>
    Public Property bolt_edge_dist_top() As Double
        Get
            Return Me.prop_bolt_edge_dist_top
        End Get
        Set
            Me.prop_bolt_edge_dist_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Flangeorbp_Connected_Top")>
    Public Property FlangeOrBP_connected_top() As Boolean
        Get
            Return Me.prop_FlangeOrBP_connected_top
        End Get
        Set
            Me.prop_FlangeOrBP_connected_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Grade_Top")>
    Public Property weld_grade_top() As Double
        Get
            Return Me.prop_weld_grade_top
        End Get
        Set
            Me.prop_weld_grade_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Trans_Type_Top")>
    Public Property weld_trans_type_top() As String
        Get
            Return Me.prop_weld_trans_type_top
        End Get
        Set
            Me.prop_weld_trans_type_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Trans_Length_Top")>
    Public Property weld_trans_length_top() As Double
        Get
            Return Me.prop_weld_trans_length_top
        End Get
        Set
            Me.prop_weld_trans_length_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Groove_Depth_Top")>
    Public Property weld_groove_depth_top() As Double
        Get
            Return Me.prop_weld_groove_depth_top
        End Get
        Set
            Me.prop_weld_groove_depth_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Groove_Angle_Top")>
    Public Property weld_groove_angle_top() As Integer
        Get
            Return Me.prop_weld_groove_angle_top
        End Get
        Set
            Me.prop_weld_groove_angle_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Trans_Fillet_Size_Top")>
    Public Property weld_trans_fillet_size_top() As Double
        Get
            Return Me.prop_weld_trans_fillet_size_top
        End Get
        Set
            Me.prop_weld_trans_fillet_size_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Trans_Eff_Throat_Top")>
    Public Property weld_trans_eff_throat_top() As Double
        Get
            Return Me.prop_weld_trans_eff_throat_top
        End Get
        Set
            Me.prop_weld_trans_eff_throat_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Long_Type_Top")>
    Public Property weld_long_type_top() As String
        Get
            Return Me.prop_weld_long_type_top
        End Get
        Set
            Me.prop_weld_long_type_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Long_Length_Top")>
    Public Property weld_long_length_top() As Double
        Get
            Return Me.prop_weld_long_length_top
        End Get
        Set
            Me.prop_weld_long_length_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Long_Fillet_Size_Top")>
    Public Property weld_long_fillet_size_top() As Double
        Get
            Return Me.prop_weld_long_fillet_size_top
        End Get
        Set
            Me.prop_weld_long_fillet_size_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Weld_Long_Eff_Throat_Top")>
    Public Property weld_long_eff_throat_top() As Double
        Get
            Return Me.prop_weld_long_eff_throat_top
        End Get
        Set
            Me.prop_weld_long_eff_throat_top = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Conn_Length_Bot")>
    Public Property conn_length_bot() As Double
        Get
            Return Me.prop_conn_length_bot
        End Get
        Set
            Me.prop_conn_length_bot = Value
        End Set
    End Property
    <Category("Member Property Flat Plate"), Description(""), DisplayName("Conn_Length_Top")>
    Public Property conn_length_top() As Double
        Get
            Return Me.prop_conn_length_top
        End Get
        Set
            Me.prop_conn_length_top = Value
        End Set
    End Property

End Class
Partial Public Class bolt_prop_flat_plate
    Private prop_ID As Integer
    Private prop_name As String
    Private prop_description As String
    Private prop_diam As Double
    Private prop_area As Double
    Private prop_fu_bolt As Double
    Private prop_sleeve_diam_out As Double
    Private prop_sleeve_diam_in As Double
    Private prop_fu_sleeve As Double
    Private prop_bolt_n_sleeve_shear_revF As Double
    Private prop_bolt_x_sleeve_shear_revF As Double
    Private prop_bolt_n_sleeve_shear_revG As Double
    Private prop_bolt_x_sleeve_shear_revG As Double
    Private prop_bolt_n_sleeve_shear_revH As Double
    Private prop_bolt_x_sleeve_shear_revH As Double
    Private prop_rb_applied_revH As Boolean

    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Name")>
    Public Property name() As String
        Get
            Return Me.prop_name
        End Get
        Set
            Me.prop_name = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Description")>
    Public Property description() As String
        Get
            Return Me.prop_description
        End Get
        Set
            Me.prop_description = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Diam")>
    Public Property diam() As Double
        Get
            Return Me.prop_diam
        End Get
        Set
            Me.prop_diam = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Area")>
    Public Property area() As Double
        Get
            Return Me.prop_area
        End Get
        Set
            Me.prop_area = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Fu_Bolt")>
    Public Property fu_bolt() As Double
        Get
            Return Me.prop_fu_bolt
        End Get
        Set
            Me.prop_fu_bolt = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Sleeve_Diam_Out")>
    Public Property sleeve_diam_out() As Double
        Get
            Return Me.prop_sleeve_diam_out
        End Get
        Set
            Me.prop_sleeve_diam_out = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Sleeve_Diam_In")>
    Public Property sleeve_diam_in() As Double
        Get
            Return Me.prop_sleeve_diam_in
        End Get
        Set
            Me.prop_sleeve_diam_in = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Fu_Sleeve")>
    Public Property fu_sleeve() As Double
        Get
            Return Me.prop_fu_sleeve
        End Get
        Set
            Me.prop_fu_sleeve = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Bolt_N_Sleeve_Shear_Revf")>
    Public Property bolt_n_sleeve_shear_revF() As Double
        Get
            Return Me.prop_bolt_n_sleeve_shear_revF
        End Get
        Set
            Me.prop_bolt_n_sleeve_shear_revF = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Bolt_X_Sleeve_Shear_Revf")>
    Public Property bolt_x_sleeve_shear_revF() As Double
        Get
            Return Me.prop_bolt_x_sleeve_shear_revF
        End Get
        Set
            Me.prop_bolt_x_sleeve_shear_revF = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Bolt_N_Sleeve_Shear_Revg")>
    Public Property bolt_n_sleeve_shear_revG() As Double
        Get
            Return Me.prop_bolt_n_sleeve_shear_revG
        End Get
        Set
            Me.prop_bolt_n_sleeve_shear_revG = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Bolt_X_Sleeve_Shear_Revg")>
    Public Property bolt_x_sleeve_shear_revG() As Double
        Get
            Return Me.prop_bolt_x_sleeve_shear_revG
        End Get
        Set
            Me.prop_bolt_x_sleeve_shear_revG = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Bolt_N_Sleeve_Shear_Revh")>
    Public Property bolt_n_sleeve_shear_revH() As Double
        Get
            Return Me.prop_bolt_n_sleeve_shear_revH
        End Get
        Set
            Me.prop_bolt_n_sleeve_shear_revH = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Bolt_X_Sleeve_Shear_Revh")>
    Public Property bolt_x_sleeve_shear_revH() As Double
        Get
            Return Me.prop_bolt_x_sleeve_shear_revH
        End Get
        Set
            Me.prop_bolt_x_sleeve_shear_revH = Value
        End Set
    End Property
    <Category("Bolt Propoerties Flat Plate"), Description(""), DisplayName("Rb_Applied_Revh")>
    Public Property rb_applied_revH() As Boolean
        Get
            Return Me.prop_rb_applied_revH
        End Get
        Set
            Me.prop_rb_applied_revH = Value
        End Set
    End Property

End Class

#End Region

#Region "Guys"
Partial Public Class guy_anchor_group
    Private prop_model_id As Integer
    Private prop_ID As Integer
    Private prop_guy_anchor_group_name As String
    Private prop_guy_anchor_qty As Integer
    Private prop_guy_anchor_identical As Boolean
    Private prop_guy_anchor_symetrical As Boolean

    <Category("Guy Anchor Group"), Description(""), DisplayName("Model_Id")>
    Public Property model_id() As Integer
        Get
            Return Me.prop_model_id
        End Get
        Set
            Me.prop_model_id = Value
        End Set
    End Property
    <Category("Guy Anchor Group"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Guy Anchor Group"), Description(""), DisplayName("Guy_Anchor_Group_Name")>
    Public Property guy_anchor_group_name() As String
        Get
            Return Me.prop_guy_anchor_group_name
        End Get
        Set
            Me.prop_guy_anchor_group_name = Value
        End Set
    End Property
    <Category("Guy Anchor Group"), Description(""), DisplayName("Guy_Anchor_Qty")>
    Public Property guy_anchor_qty() As Integer
        Get
            Return Me.prop_guy_anchor_qty
        End Get
        Set
            Me.prop_guy_anchor_qty = Value
        End Set
    End Property
    <Category("Guy Anchor Group"), Description(""), DisplayName("Guy_Anchor_Identical")>
    Public Property guy_anchor_identical() As Boolean
        Get
            Return Me.prop_guy_anchor_identical
        End Get
        Set
            Me.prop_guy_anchor_identical = Value
        End Set
    End Property
    <Category("Guy Anchor Group"), Description(""), DisplayName("Guy_Anchor_Symetrical")>
    Public Property guy_anchor_symetrical() As Boolean
        Get
            Return Me.prop_guy_anchor_symetrical
        End Get
        Set
            Me.prop_guy_anchor_symetrical = Value
        End Set
    End Property

End Class
Partial Public Class guy_anchor_detail
    Private prop_guy_anchor_group_id As Integer
    Private prop_ID As Integer
    Private prop_guy_anchor_name As String
    Private prop_guy_anchor_radius As Double
    Private prop_guy_anchor_elevation As Double
    Private prop_guy_anchor_azimuth As Double

    <Category("Guy Anchor Detail"), Description(""), DisplayName("Guy_Anchor_Group_Id")>
    Public Property guy_anchor_group_id() As Integer
        Get
            Return Me.prop_guy_anchor_group_id
        End Get
        Set
            Me.prop_guy_anchor_group_id = Value
        End Set
    End Property
    <Category("Guy Anchor Detail"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Guy Anchor Detail"), Description(""), DisplayName("Guy_Anchor_Name")>
    Public Property guy_anchor_name() As String
        Get
            Return Me.prop_guy_anchor_name
        End Get
        Set
            Me.prop_guy_anchor_name = Value
        End Set
    End Property
    <Category("Guy Anchor Detail"), Description(""), DisplayName("Guy_Anchor_Radius")>
    Public Property guy_anchor_radius() As Double
        Get
            Return Me.prop_guy_anchor_radius
        End Get
        Set
            Me.prop_guy_anchor_radius = Value
        End Set
    End Property
    <Category("Guy Anchor Detail"), Description(""), DisplayName("Guy_Anchor_Elevation")>
    Public Property guy_anchor_elevation() As Double
        Get
            Return Me.prop_guy_anchor_elevation
        End Get
        Set
            Me.prop_guy_anchor_elevation = Value
        End Set
    End Property
    <Category("Guy Anchor Detail"), Description(""), DisplayName("Guy_Anchor_Azimuth")>
    Public Property guy_anchor_azimuth() As Double
        Get
            Return Me.prop_guy_anchor_azimuth
        End Get
        Set
            Me.prop_guy_anchor_azimuth = Value
        End Set
    End Property

End Class
Partial Public Class guy_attachment
    Private prop_model_id As Integer
    Private prop_ID As Integer
    Private prop_guy_attachment_type As String
    Private prop_guy_attachment_elev As Double
    Private prop_guy_attachment_top_mount_elev As Double
    Private prop_guy_attachment_bot_mount_elev As Double
    Private prop_guy_attachment_spread As Double
    Private prop_guy_attachment_members_identical As Boolean

    <Category("Guy Attachment"), Description(""), DisplayName("Model_Id")>
    Public Property model_id() As Integer
        Get
            Return Me.prop_model_id
        End Get
        Set
            Me.prop_model_id = Value
        End Set
    End Property
    <Category("Guy Attachment"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Guy Attachment"), Description(""), DisplayName("Guy_Attachment_Type")>
    Public Property guy_attachment_type() As String
        Get
            Return Me.prop_guy_attachment_type
        End Get
        Set
            Me.prop_guy_attachment_type = Value
        End Set
    End Property
    <Category("Guy Attachment"), Description(""), DisplayName("Guy_Attachment_Elev")>
    Public Property guy_attachment_elev() As Double
        Get
            Return Me.prop_guy_attachment_elev
        End Get
        Set
            Me.prop_guy_attachment_elev = Value
        End Set
    End Property
    <Category("Guy Attachment"), Description(""), DisplayName("Guy_Attachment_Top_Mount_Elev")>
    Public Property guy_attachment_top_mount_elev() As Double
        Get
            Return Me.prop_guy_attachment_top_mount_elev
        End Get
        Set
            Me.prop_guy_attachment_top_mount_elev = Value
        End Set
    End Property
    <Category("Guy Attachment"), Description(""), DisplayName("Guy_Attachment_Bot_Mount_Elev")>
    Public Property guy_attachment_bot_mount_elev() As Double
        Get
            Return Me.prop_guy_attachment_bot_mount_elev
        End Get
        Set
            Me.prop_guy_attachment_bot_mount_elev = Value
        End Set
    End Property
    <Category("Guy Attachment"), Description(""), DisplayName("Guy_Attachment_Spread")>
    Public Property guy_attachment_spread() As Double
        Get
            Return Me.prop_guy_attachment_spread
        End Get
        Set
            Me.prop_guy_attachment_spread = Value
        End Set
    End Property
    <Category("Guy Attachment"), Description(""), DisplayName("Guy_Attachment_Members_Identical")>
    Public Property guy_attachment_members_identical() As Boolean
        Get
            Return Me.prop_guy_attachment_members_identical
        End Get
        Set
            Me.prop_guy_attachment_members_identical = Value
        End Set
    End Property

End Class
Partial Public Class guy_attachment_bracing_detail
    Private prop_guy_attachment_id As Integer
    Private prop_ID As Integer
    Private prop_face_letter As String
    Private prop_bracing_type As String
    Private prop_bracing_sect_type As String
    Private prop_bracing_sect_prop As Integer
    Private prop_bracing_mat_prop As Integer
    Private prop_bracing_conn_end_cond As String
    Private prop_bracing_conn_mirror As Boolean
    Private prop_bracing_conn_start_pattern As String
    Private prop_bracing_conn_start_bolt_size As Integer
    Private prop_bracing_conn_start_bolt_mat_prop As Integer
    Private prop_bracing_conn_start_bolt_thread As String
    Private prop_bracing_conn_start_edge As Double
    Private prop_bracing_conn_start_pitch As Double
    Private prop_bracing_conn_start_gage As Double
    Private prop_bracing_conn_start_gage_space As Double
    Private prop_bracing_conn_end_bolt_pattern As String
    Private prop_bracing_conn_end_bolt_size As Integer
    Private prop_bracing_conn_end_bolt_mat_prop As Integer
    Private prop_bracing_conn_end_bolt_thread As String
    Private prop_bracing_conn_end_edge As Double
    Private prop_bracing_conn_end_pitch As Double
    Private prop_bracing_conn_end_gage As Double
    Private prop_bracing_conn_end_gage_space As Double
    Private prop_bracing_dp_connector_type As String
    Private prop_bracing_dp_connector_spacing As Double
    Private prop_bracing_dp_end_connection As Boolean
    Private prop_bracing_dp_fully_comp As Boolean
    Private prop_bracing_dp_crushing As Boolean
    Private prop_bracing_dp_tension As Boolean
    Private prop_bracing_dp_eccentricity As Boolean
    Private prop_bracing_dp_U As Double
    Private prop_bracing_dp_Lex As Double
    Private prop_bracing_dp_Ley As Double
    Private prop_bracing_dp_Lez As Double
    Private prop_bracing_dp_Kx As Double
    Private prop_bracing_dp_Ky As Double
    Private prop_bracing_dp_Kz As Double

    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Guy_Attachment_Id")>
    Public Property guy_attachment_id() As Integer
        Get
            Return Me.prop_guy_attachment_id
        End Get
        Set
            Me.prop_guy_attachment_id = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Face_Letter")>
    Public Property face_letter() As String
        Get
            Return Me.prop_face_letter
        End Get
        Set
            Me.prop_face_letter = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Type")>
    Public Property bracing_type() As String
        Get
            Return Me.prop_bracing_type
        End Get
        Set
            Me.prop_bracing_type = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Sect_Type")>
    Public Property bracing_sect_type() As String
        Get
            Return Me.prop_bracing_sect_type
        End Get
        Set
            Me.prop_bracing_sect_type = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Sect_Prop")>
    Public Property bracing_sect_prop() As Integer
        Get
            Return Me.prop_bracing_sect_prop
        End Get
        Set
            Me.prop_bracing_sect_prop = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Mat_Prop")>
    Public Property bracing_mat_prop() As Integer
        Get
            Return Me.prop_bracing_mat_prop
        End Get
        Set
            Me.prop_bracing_mat_prop = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Cond")>
    Public Property bracing_conn_end_cond() As String
        Get
            Return Me.prop_bracing_conn_end_cond
        End Get
        Set
            Me.prop_bracing_conn_end_cond = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Mirror")>
    Public Property bracing_conn_mirror() As Boolean
        Get
            Return Me.prop_bracing_conn_mirror
        End Get
        Set
            Me.prop_bracing_conn_mirror = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Pattern")>
    Public Property bracing_conn_start_pattern() As String
        Get
            Return Me.prop_bracing_conn_start_pattern
        End Get
        Set
            Me.prop_bracing_conn_start_pattern = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Bolt_Size")>
    Public Property bracing_conn_start_bolt_size() As Integer
        Get
            Return Me.prop_bracing_conn_start_bolt_size
        End Get
        Set
            Me.prop_bracing_conn_start_bolt_size = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Bolt_Mat_Prop")>
    Public Property bracing_conn_start_bolt_mat_prop() As Integer
        Get
            Return Me.prop_bracing_conn_start_bolt_mat_prop
        End Get
        Set
            Me.prop_bracing_conn_start_bolt_mat_prop = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Bolt_Thread")>
    Public Property bracing_conn_start_bolt_thread() As String
        Get
            Return Me.prop_bracing_conn_start_bolt_thread
        End Get
        Set
            Me.prop_bracing_conn_start_bolt_thread = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Edge")>
    Public Property bracing_conn_start_edge() As Double
        Get
            Return Me.prop_bracing_conn_start_edge
        End Get
        Set
            Me.prop_bracing_conn_start_edge = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Pitch")>
    Public Property bracing_conn_start_pitch() As Double
        Get
            Return Me.prop_bracing_conn_start_pitch
        End Get
        Set
            Me.prop_bracing_conn_start_pitch = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Gage")>
    Public Property bracing_conn_start_gage() As Double
        Get
            Return Me.prop_bracing_conn_start_gage
        End Get
        Set
            Me.prop_bracing_conn_start_gage = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_Start_Gage_Space")>
    Public Property bracing_conn_start_gage_space() As Double
        Get
            Return Me.prop_bracing_conn_start_gage_space
        End Get
        Set
            Me.prop_bracing_conn_start_gage_space = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Bolt_Pattern")>
    Public Property bracing_conn_end_bolt_pattern() As String
        Get
            Return Me.prop_bracing_conn_end_bolt_pattern
        End Get
        Set
            Me.prop_bracing_conn_end_bolt_pattern = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Bolt_Size")>
    Public Property bracing_conn_end_bolt_size() As Integer
        Get
            Return Me.prop_bracing_conn_end_bolt_size
        End Get
        Set
            Me.prop_bracing_conn_end_bolt_size = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Bolt_Mat_Prop")>
    Public Property bracing_conn_end_bolt_mat_prop() As Integer
        Get
            Return Me.prop_bracing_conn_end_bolt_mat_prop
        End Get
        Set
            Me.prop_bracing_conn_end_bolt_mat_prop = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Bolt_Thread")>
    Public Property bracing_conn_end_bolt_thread() As String
        Get
            Return Me.prop_bracing_conn_end_bolt_thread
        End Get
        Set
            Me.prop_bracing_conn_end_bolt_thread = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Edge")>
    Public Property bracing_conn_end_edge() As Double
        Get
            Return Me.prop_bracing_conn_end_edge
        End Get
        Set
            Me.prop_bracing_conn_end_edge = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Pitch")>
    Public Property bracing_conn_end_pitch() As Double
        Get
            Return Me.prop_bracing_conn_end_pitch
        End Get
        Set
            Me.prop_bracing_conn_end_pitch = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Gage")>
    Public Property bracing_conn_end_gage() As Double
        Get
            Return Me.prop_bracing_conn_end_gage
        End Get
        Set
            Me.prop_bracing_conn_end_gage = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Conn_End_Gage_Space")>
    Public Property bracing_conn_end_gage_space() As Double
        Get
            Return Me.prop_bracing_conn_end_gage_space
        End Get
        Set
            Me.prop_bracing_conn_end_gage_space = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Connector_Type")>
    Public Property bracing_dp_connector_type() As String
        Get
            Return Me.prop_bracing_dp_connector_type
        End Get
        Set
            Me.prop_bracing_dp_connector_type = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Connector_Spacing")>
    Public Property bracing_dp_connector_spacing() As Double
        Get
            Return Me.prop_bracing_dp_connector_spacing
        End Get
        Set
            Me.prop_bracing_dp_connector_spacing = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_End_Connection")>
    Public Property bracing_dp_end_connection() As Boolean
        Get
            Return Me.prop_bracing_dp_end_connection
        End Get
        Set
            Me.prop_bracing_dp_end_connection = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Fully_Comp")>
    Public Property bracing_dp_fully_comp() As Boolean
        Get
            Return Me.prop_bracing_dp_fully_comp
        End Get
        Set
            Me.prop_bracing_dp_fully_comp = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Crushing")>
    Public Property bracing_dp_crushing() As Boolean
        Get
            Return Me.prop_bracing_dp_crushing
        End Get
        Set
            Me.prop_bracing_dp_crushing = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Tension")>
    Public Property bracing_dp_tension() As Boolean
        Get
            Return Me.prop_bracing_dp_tension
        End Get
        Set
            Me.prop_bracing_dp_tension = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Eccentricity")>
    Public Property bracing_dp_eccentricity() As Boolean
        Get
            Return Me.prop_bracing_dp_eccentricity
        End Get
        Set
            Me.prop_bracing_dp_eccentricity = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_U")>
    Public Property bracing_dp_U() As Double
        Get
            Return Me.prop_bracing_dp_U
        End Get
        Set
            Me.prop_bracing_dp_U = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Lex")>
    Public Property bracing_dp_Lex() As Double
        Get
            Return Me.prop_bracing_dp_Lex
        End Get
        Set
            Me.prop_bracing_dp_Lex = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Ley")>
    Public Property bracing_dp_Ley() As Double
        Get
            Return Me.prop_bracing_dp_Ley
        End Get
        Set
            Me.prop_bracing_dp_Ley = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Lez")>
    Public Property bracing_dp_Lez() As Double
        Get
            Return Me.prop_bracing_dp_Lez
        End Get
        Set
            Me.prop_bracing_dp_Lez = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Kx")>
    Public Property bracing_dp_Kx() As Double
        Get
            Return Me.prop_bracing_dp_Kx
        End Get
        Set
            Me.prop_bracing_dp_Kx = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Ky")>
    Public Property bracing_dp_Ky() As Double
        Get
            Return Me.prop_bracing_dp_Ky
        End Get
        Set
            Me.prop_bracing_dp_Ky = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Detail"), Description(""), DisplayName("Bracing_Dp_Kz")>
    Public Property bracing_dp_Kz() As Double
        Get
            Return Me.prop_bracing_dp_Kz
        End Get
        Set
            Me.prop_bracing_dp_Kz = Value
        End Set
    End Property

End Class
Partial Public Class guy_attachment_bracing_cust_cap
    Private prop_guy_attachment_id As Integer
    Private prop_ID As Integer
    Private prop_face_letter As String
    Private prop_bracing_type As String
    Private prop_bracing_cc_code As String
    Private prop_bracing_cc_comp As Double
    Private prop_bracing_cc_ten As Double
    Private prop_bracing_cc_conn_comp As Double
    Private prop_bracing_cc_conn_ten As Double
    Private prop_bracing_cc_pass_rating As Decimal

    <Category("Guy Attachment Bracing Custom Capacity"), Description(""), DisplayName("Guy_Attachment_Id")>
    Public Property guy_attachment_id() As Integer
        Get
            Return Me.prop_guy_attachment_id
        End Get
        Set
            Me.prop_guy_attachment_id = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Custom Capacity"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Custom Capacity"), Description(""), DisplayName("Face_Letter")>
    Public Property face_letter() As String
        Get
            Return Me.prop_face_letter
        End Get
        Set
            Me.prop_face_letter = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Custom Capacity"), Description(""), DisplayName("Bracing_Type")>
    Public Property bracing_type() As String
        Get
            Return Me.prop_bracing_type
        End Get
        Set
            Me.prop_bracing_type = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Custom Capacity"), Description(""), DisplayName("Bracing_Cc_Code")>
    Public Property bracing_cc_code() As String
        Get
            Return Me.prop_bracing_cc_code
        End Get
        Set
            Me.prop_bracing_cc_code = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Custom Capacity"), Description(""), DisplayName("Bracing_Cc_Comp")>
    Public Property bracing_cc_comp() As Double
        Get
            Return Me.prop_bracing_cc_comp
        End Get
        Set
            Me.prop_bracing_cc_comp = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Custom Capacity"), Description(""), DisplayName("Bracing_Cc_Ten")>
    Public Property bracing_cc_ten() As Double
        Get
            Return Me.prop_bracing_cc_ten
        End Get
        Set
            Me.prop_bracing_cc_ten = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Custom Capacity"), Description(""), DisplayName("Bracing_Cc_Conn_Comp")>
    Public Property bracing_cc_conn_comp() As Double
        Get
            Return Me.prop_bracing_cc_conn_comp
        End Get
        Set
            Me.prop_bracing_cc_conn_comp = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Custom Capacity"), Description(""), DisplayName("Bracing_Cc_Conn_Ten")>
    Public Property bracing_cc_conn_ten() As Double
        Get
            Return Me.prop_bracing_cc_conn_ten
        End Get
        Set
            Me.prop_bracing_cc_conn_ten = Value
        End Set
    End Property
    <Category("Guy Attachment Bracing Custom Capacity"), Description(""), DisplayName("Bracing_Cc_Pass_Rating")>
    Public Property bracing_cc_pass_rating() As Decimal
        Get
            Return Me.prop_bracing_cc_pass_rating
        End Get
        Set
            Me.prop_bracing_cc_pass_rating = Value
        End Set
    End Property

End Class
Partial Public Class guy_junction_detail
    Private prop_guy_attachment_id As Integer
    Private prop_guy_anchor_group_id As Integer
    Private prop_ID As Integer
    Private prop_attachment_point As Integer
    Private prop_guy_wire_type As String
    Private prop_guy_wire_prop As Integer
    Private prop_initial_tension As Double
    Private prop_end_fitting_efficiency As Double

    <Category("Guy Junction Detail"), Description(""), DisplayName("Guy_Attachment_Id")>
    Public Property guy_attachment_id() As Integer
        Get
            Return Me.prop_guy_attachment_id
        End Get
        Set
            Me.prop_guy_attachment_id = Value
        End Set
    End Property
    <Category("Guy Junction Detail"), Description(""), DisplayName("Guy_Anchor_Group_Id")>
    Public Property guy_anchor_group_id() As Integer
        Get
            Return Me.prop_guy_anchor_group_id
        End Get
        Set
            Me.prop_guy_anchor_group_id = Value
        End Set
    End Property
    <Category("Guy Junction Detail"), Description(""), DisplayName("Id")>
    Public Property ID() As Integer
        Get
            Return Me.prop_ID
        End Get
        Set
            Me.prop_ID = Value
        End Set
    End Property
    <Category("Guy Junction Detail"), Description(""), DisplayName("Attachment_Point")>
    Public Property attachment_point() As Integer
        Get
            Return Me.prop_attachment_point
        End Get
        Set
            Me.prop_attachment_point = Value
        End Set
    End Property
    <Category("Guy Junction Detail"), Description(""), DisplayName("Guy_Wire_Type")>
    Public Property guy_wire_type() As String
        Get
            Return Me.prop_guy_wire_type
        End Get
        Set
            Me.prop_guy_wire_type = Value
        End Set
    End Property
    <Category("Guy Junction Detail"), Description(""), DisplayName("Guy_Wire_Prop")>
    Public Property guy_wire_prop() As Integer
        Get
            Return Me.prop_guy_wire_prop
        End Get
        Set
            Me.prop_guy_wire_prop = Value
        End Set
    End Property
    <Category("Guy Junction Detail"), Description(""), DisplayName("Initial_Tension")>
    Public Property initial_tension() As Double
        Get
            Return Me.prop_initial_tension
        End Get
        Set
            Me.prop_initial_tension = Value
        End Set
    End Property
    <Category("Guy Junction Detail"), Description(""), DisplayName("End_Fitting_Efficiency")>
    Public Property end_fitting_efficiency() As Double
        Get
            Return Me.prop_end_fitting_efficiency
        End Get
        Set
            Me.prop_end_fitting_efficiency = Value
        End Set
    End Property

End Class
#End Region

