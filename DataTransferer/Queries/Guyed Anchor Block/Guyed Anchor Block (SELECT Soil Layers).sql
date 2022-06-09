[EXISTING MODEL]

--SELECT --WORK IN PROGRESS
--	sm.id model_id
--	,fd.id foundation_id
--	,gabd.ID anchor_id
--	--,sl.anchor_id
--	,sl.ID soil_layer_id
--	,sl.bottom_depth
--	,sl.effective_soil_density
--	,sl.cohesion
--	,sl.friction_angle
--	,sl.skin_friction_override_uplift
--	,sl.spt_blow_count
--	,sl.local_soil_layer_id
--	,sl.local_soil_profile
--FROM 
--	anchor_soil_layer sl
--	,foundation_details fd
--	,anchor_details gabd 
--	,structure_model sm
--WHERE 
--	sl.anchor_id=gabd.ID
--	AND gabd.foundation_id=fd.ID
--	AND fd.model_id=sm.id
--	AND sm.ID=@ModelID
--ORDER BY
--	sl.anchor_id
--	,sl.bottom_dep

	--Anchor Block Soil Layer 
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
,absl.* 
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.anchor_block_details abd 
,fnd.anchor_block_soil_layer absl 
WHERE 
smx.model_id=@ModelID
AND smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = abd.id 
AND absl.anchor_id = abd.id 
--AND fd.foundation_type = @FndType
--AND smx.bus_unit = @BU 
--AND smx.structure_id = @STR_ID 