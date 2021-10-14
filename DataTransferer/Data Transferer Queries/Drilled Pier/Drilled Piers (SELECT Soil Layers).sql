[EXISTING MODEL]

	--Drilled Pier Soil Layer 
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
,sl.* 
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.drilled_pier_soil_layer sl
WHERE 
smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = sl.id 
--AND fd.foundation_type = @FndType
AND smx.bus_unit = @BU 
AND smx.structure_id = @STR_ID


--SELECT 
--	sm.id model_id
--	,fd.id foundation_id
--	--,dpd.ID drilled_pier_id
--	,sl.drilled_pier_id
--	,sl.ID soil_layer_id
--	,sl.bottom_depth
--	,sl.effective_soil_density
--	,sl.cohesion
--	,sl.friction_angle
--	,sl.skin_friction_override_comp
--	,sl.skin_friction_override_uplift
--	,sl.nominal_bearing_capacity
--	,sl.spt_blow_count
--	,sl.local_soil_layer_id
--	--,sl.local_drilled_pier_id
--FROM 
--	drilled_pier_soil_layer sl
--	,foundation_details fd
--	,drilled_pier_details dpd
--	,structure_model sm
--WHERE 
--	sl.drilled_pier_id=dpd.ID
--	AND dpd.foundation_id=fd.ID
--	AND fd.model_id=sm.id
--	AND sm.ID=@ModelID
--ORDER BY
--	sl.drilled_pier_id
--	,sl.bottom_depth