[EXISTING MODEL]

	--Drilled Pier Details 
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
--,dpd.*
,dpd.id pier_id
,dpd.local_drilled_pier_id
,dpd.local_drilled_pier_profile
,dpd.foundation_depth
,dpd.extension_above_grade
,dpd.groundwater_depth
,dpd.assume_min_steel
,dpd.check_shear_along_depth
,dpd.utilize_shear_friction_methodology
,dpd.embedded_pole
,dpd.belled_pier
,dpd.soil_layer_quantity
,dpd.concrete_compressive_strength
,dpd.tie_yield_strength
,dpd.longitudinal_rebar_yield_strength
,dpd.rebar_effective_depths
,dpd.rebar_cage_2_fy_override
,dpd.rebar_cage_3_fy_override
,dpd.shear_override_crit_depth
,dpd.shear_crit_depth_override_comp
,dpd.shear_crit_depth_override_uplift
,dpd.bearing_toggle_type
,dpd.tool_version
,dpd.modified

FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.drilled_pier_details dpd 
WHERE 
smx.model_id = @ModelID 
AND smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = dpd.id 
--AND fd.foundation_type = @FndType
--AND smx.bus_unit = @BU 
--AND smx.structure_id = @STR_ID

--SELECT 
--	sm.bus_unit
--	,sm.structure_id str_id
--	,sm.id model_id
--	,fd.ID foundation_id
--	,dpd.ID drilled_pier_id
--	,fd.foundation_type 	
--	,dpd.foundation_depth
--	,dpd.extension_above_grade
--	,dpd.groundwater_depth
--	,dpd.assume_min_steel
--	,dpd.check_shear_along_depth
--	,dpd.utilize_shear_friction_methodology
--	,dpd.embedded_pole
--	,dpd.belled_pier
--	,dpd.soil_layer_quantity
--	,dpd.concrete_compressive_strength
--	,dpd.tie_yield_strength
--	,dpd.longitudinal_rebar_yield_strength
--	,dpd.rebar_effective_depths
--	,dpd.rebar_cage_2_fy_override
--	,dpd.rebar_cage_3_fy_override
--	,dpd.shear_override_crit_depth
--	,dpd.shear_crit_depth_override_comp
--	,dpd.shear_crit_depth_override_uplift
--	--,dpd.drilled_pier_profile_qty
--	--,dpd.soil_profiles
--	,dpd.local_drilled_pier_id
--	--,dpd.rho_override_1
--	--,dpd.rho_override_2
--	--,dpd.rho_override_3
--	--,dpd.rho_override_4
--	--,dpd.rho_override_5
--	,dpd.bearing_type_toggle
--	,dpd.modified,dpd.local_drilled_pier_profile
--FROM 
--	foundation_details fd
--	,drilled_pier_details dpd 
--	,structure_model sm
--WHERE 
--	dpd.foundation_id=fd.ID
--	AND fd.model_id=sm.id
--	AND sm.ID=@Model