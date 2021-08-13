[EXISTING MODEL]

SELECT 
	sm.id model_id
	,fd.id foundation_id
	,dpd.ID drilled_pier_id
	,sl.drilled_pier_id
	,sl.ID soil_layer_id
	,sl.bottom_depth
	,sl.effective_soil_density
	,sl.cohesion
	,sl.friction_angle
	,sl.skin_friction_override_comp
	,sl.skin_friction_override_uplift
	,sl.nominal_bearing_capacity
	,sl.spt_blow_count
	,sl.local_soil_layer_id
	,sl.local_drilled_pier_id
FROM 
	drilled_pier_soil_layer sl
	,foundation_details fd
	,drilled_pier_details dpd
	,structure_model sm
WHERE 
	sl.drilled_pier_id=dpd.ID
	AND dpd.foundation_id=fd.ID
	AND fd.model_id=sm.id
	AND sm.ID=@ModelID
ORDER BY
	sl.drilled_pier_id
	,sl.bottom_depth