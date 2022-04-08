[EXISTING MODEL]

--SELECT --WORK IN PROGRESS
--	sm.bus_unit
--	,sm.structure_id str_id
--	,sm.id model_id
--	,fd.ID foundation_id
--	,gabd.ID anchor_id
--	,fd.foundation_type 
--	,gabd.anchor_depth
--	,gabd.anchor_width
--	,gabd.anchor_thickness
--	,gabd.anchor_length
--	,gabd.anchor_toe_width
--	,gabd.anchor_top_rebar_size
--	,gabd.anchor_top_rebar_quantity
--	,gabd.anchor_front_rebar_size
--	,gabd.anchor_front_rebar_quantity
--	,gabd.anchor_stirrup_size
--	,gabd.anchor_shaft_diameter
--	,gabd.anchor_shaft_quantity
--	,gabd.anchor_shaft_area_override
--	,gabd.anchor_shaft_shear_lag_factor
--	,gabd.concrete_compressive_strength
--	,gabd.clear_cover
--	,gabd.anchor_shaft_yield_strength
--	,gabd.anchor_shaft_ultimate_strength
--	,gabd.neglect_depth
--	,gabd.groundwater_depth
--	,gabd.soil_layer_quantity
--	,gabd.tool_version
--	,gabd.anchor_shaft_section
--	,gabd.anchor_rebar_grade
--	,gabd.anchor_shaft_known
--	,gabd.basic_soil_check
--	,gabd.structural_check
--	,gabd.rebar_known
--	,gabd.local_anchor_id
--	,gabd.local_anchor_profile
--	,gabd.foundation_id
--FROM 
--	foundation_details fd
--	,anchor_details gabd 
--	,structure_model sm
--WHERE 
--	gabd.foundation_id=fd.ID
--	AND fd.model_id=sm.id
--	AND sm.ID=@Model

	--Anchor Block Details 
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
,abd.* 
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.anchor_block_details abd 
WHERE 
smx.model_id=@ModelID
AND smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = abd.id 
--AND fd.foundation_type = @FndType
--AND smx.bus_unit = @BU 
--AND smx.structure_id = @STR_ID