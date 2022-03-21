[EXISTING MODEL]

SELECT
	sm.ID model_id
	,pstr.ID pole_structure_id
	,pr.ID reinf_db_id
	,pr.local_id
	,pr.name
	,pr.type
	,pr.b
	,pr.h
	,pr.sr_diam
	,pr.channel_thkns_web
	,pr.channel_thkns_flange
	,pr.channel_eo
	,pr.channel_J
	,pr.channel_Cw
	,pr.area_gross
	,pr.centroid
	,pr.istension
	,pr.matl_id
	,pr.local_matl_id
	,pr.Ix
	,pr.Iy
	,pr.Lu
	,pr.Kx
	,pr.Ky
	,pr.bolt_hole_size
	,pr.area_net
	,pr.shear_lag
	,pr.connection_type_bot
	,pr.connection_cap_revF_bot
	,pr.connection_cap_revG_bot
	,pr.connection_cap_revH_bot
	,pr.bolt_type_id_bot
	,pr.local_bolt_id_bot
	,pr.bolt_N_or_X_bot
	,pr.bolt_num_bot
	,pr.bolt_spacing_bot
	,pr.bolt_edge_dist_bot
	,pr.FlangeOrBP_connected_bot
	,pr.weld_grade_bot
	,pr.weld_trans_type_bot
	,pr.weld_trans_length_bot
	,pr.weld_groove_depth_bot
	,pr.weld_groove_angle_bot
	,pr.weld_trans_fillet_size_bot
	,pr.weld_trans_eff_throat_bot
	,pr.weld_long_type_bot
	,pr.weld_long_length_bot
	,pr.weld_long_fillet_size_bot
	,pr.weld_long_eff_throat_bot
	,pr.top_bot_connections_symmetrical
	,pr.connection_type_top
	,pr.connection_cap_revF_top
	,pr.connection_cap_revG_top
	,pr.connection_cap_revH_top
	,pr.bolt_type_id_top
	,pr.local_bolt_id_top
	,pr.bolt_N_or_X_top
	,pr.bolt_num_top
	,pr.bolt_spacing_top
	,pr.bolt_edge_dist_top
	,pr.FlangeOrBP_connected_top
	,pr.weld_grade_top
	,pr.weld_trans_type_top
	,pr.weld_trans_length_top
	,pr.weld_groove_depth_top
	,pr.weld_groove_angle_top
	,pr.weld_trans_fillet_size_top
	,pr.weld_trans_eff_throat_top
	,pr.weld_long_type_top
	,pr.weld_long_length_top
	,pr.weld_long_fillet_size_top
	,pr.weld_long_eff_throat_top
	,pr.conn_length_channel
	,pr.conn_length_bot
	,pr.conn_length_top
	,pr.cap_comp_xx_f
	,pr.cap_comp_yy_f
	,pr.cap_tens_yield_f
	,pr.cap_tens_rupture_f
	,pr.cap_shear_f
	,pr.cap_bolt_shear_bot_f
	,pr.cap_bolt_shear_top_f
	,pr.cap_boltshaft_bearing_nodeform_bot_f
	,pr.cap_boltshaft_bearing_deform_bot_f
	,pr.cap_boltshaft_bearing_nodeform_top_f
	,pr.cap_boltshaft_bearing_deform_top_f
	,pr.cap_boltreinf_bearing_nodeform_bot_f
	,pr.cap_boltreinf_bearing_deform_bot_f
	,pr.cap_boltreinf_bearing_nodeform_top_f
	,pr.cap_boltreinf_bearing_deform_top_f
	,pr.cap_weld_trans_bot_f
	,pr.cap_weld_long_bot_f
	,pr.cap_weld_trans_top_f
	,pr.cap_weld_long_top_f
	,pr.cap_comp_xx_g
	,pr.cap_comp_yy_g
	,pr.cap_tens_yield_g
	,pr.cap_tens_rupture_g
	,pr.cap_shear_g
	,pr.cap_bolt_shear_bot_g
	,pr.cap_bolt_shear_top_g
	,pr.cap_boltshaft_bearing_nodeform_bot_g
	,pr.cap_boltshaft_bearing_deform_bot_g
	,pr.cap_boltshaft_bearing_nodeform_top_g
	,pr.cap_boltshaft_bearing_deform_top_g
	,pr.cap_boltreinf_bearing_nodeform_bot_g
	,pr.cap_boltreinf_bearing_deform_bot_g
	,pr.cap_boltreinf_bearing_nodeform_top_g
	,pr.cap_boltreinf_bearing_deform_top_g
	,pr.cap_weld_trans_bot_g
	,pr.cap_weld_long_bot_g
	,pr.cap_weld_trans_top_g
	,pr.cap_weld_long_top_g
	,pr.cap_comp_xx_h
	,pr.cap_comp_yy_h
	,pr.cap_tens_yield_h
	,pr.cap_tens_rupture_h
	,pr.cap_shear_h
	,pr.cap_bolt_shear_bot_h
	,pr.cap_bolt_shear_top_h
	,pr.cap_boltshaft_bearing_nodeform_bot_h
	,pr.cap_boltshaft_bearing_deform_bot_h
	,pr.cap_boltshaft_bearing_nodeform_top_h
	,pr.cap_boltshaft_bearing_deform_top_h
	,pr.cap_boltreinf_bearing_nodeform_bot_h
	,pr.cap_boltreinf_bearing_deform_bot_h
	,pr.cap_boltreinf_bearing_nodeform_top_h
	,pr.cap_boltreinf_bearing_deform_top_h
	,pr.cap_weld_trans_bot_h
	,pr.cap_weld_long_bot_h
	,pr.cap_weld_trans_top_h
	,pr.cap_weld_long_top_h
FROM
	gen.structure_model_xref smx
	,gen.structure_model sm
	,pole.pole_structure pstr
	,pole.memb_prop_flat_plate_xref prx
	,pole.memb_prop_flat_plate pr
WHERE
	smx.model_id=@ModelID
	AND smx.model_id=sm.ID
	AND sm.pole_structure_id=pstr.ID
	AND prx.pole_structure_id = pstr.ID
	AND pr.ID = prx.reinf_id
