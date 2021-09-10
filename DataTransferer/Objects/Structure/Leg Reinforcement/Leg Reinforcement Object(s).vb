Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
Public Class LegReinforcement

#Region "Define"
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
    Private prop_intermediate_conn_spacing As Integer?
    Private prop_ki_override As Double?
    Private prop_leg_dia As Double?
    Private prop_leg_thickness As Double?
    Private prop_leg_yield_strength As Double?
    Private prop_leg_unbraced_length As Double?
    Private prop_rein_dia As Double?
    Private prop_rein_thickness As Double?
    Private prop_rein_yield_strength As Double?
    Private prop_print_bolton_conn_info As Boolean

    'assign inputs for leg location and location in tool


#End Region

#Region "Constructors"

#End Region




#Region "Leg Reinforcement Extras"
    Partial Public Class BoltOnConnections
        Private prop_leg_length_of_tower_section As Double?
        Private prop_split_pip_length As Double?
        Private prop_set_top_to_bottom As Boolean
        Private prop_qty_flange_bolt_bot As Integer?
        Private prop_colt_circle_bot As Double?
        Private prop_bolt_orientation_bot As Integer?
        Private prop_qty_flange_bolt_top As Integer?
        Private prop_bolt_circle_top As Double?
        Private prop_bolt_orientation_top As Integer?
        Private prop_threaded_rod_dia_bot As Double?
        Private prop_threaded_rod_mat_bot As String
        Private prop_threaded_rod_qty_bot As Double?
        Private prop_threaded_rod_unbraced_length_bot As Double?
        Private prop_threaded_rod_dia_top As Double?
        Private prop_threaded_rod_mat_top As String
        Private prop_threaded_rod_qty_top As Double?
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


    End Class

    Partial Public Class ArbitraryShape
        Private prop_us_name As String
        Private prop_si_name As String
        Private prop_height As Double?
        Private prop_width As Double?
        Private prop_wind_projected_width As Double?
        Private prop_perimeter As Double?
        Private prop_modulus_of_elasticity As Double?
        Private prop_density As Double?
        Private prop_area As Double?
        Private prop_QaQs As Double?
        Private prop_cw As Double?
        Private prop_Ix As Double?
        Private prop_Iy As Double?
        Private prop_J As Double?
        Private prop_Sx_top As Double?
        Private prop_Sy_left As Double?
        Private prop_Sx_bot As Double?
        Private prop_Sy_right As Double?
        Private prop_rx As Double?
        Private prop_ry As Double?
        Private prop_SFx As Double?
        Private prop_SFy As Double?
        Private prop_K_factor_adj As Double?






    End Class

    Partial Public Class tnxSectionDatabaseInfo
        Private prop_file As String
        Private prop_us_name As String
        Private prop_si_name As String
        Private prop_values As String



    End Class

    Partial Public Class tnxMaterialDatabaseInfo
        Private prop_member_mat_file As String
        Private prop_mat_name As String
        Private prop_mat_values As String




    End Class
#End Region

End Class
