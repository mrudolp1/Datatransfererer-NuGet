[EXISTING MODEL]

SELECT 
	sm.id model_id
	,fd.id foundation_id
	,dpp.drilled_pier_id
	,dpp.ID profile_id
	,dpp.local_drilled_pier_id
	,dpp.reaction_position
	,dpp.reaction_location
	,dpp.drilled_pier_profile
	,dpp.soil_profile
FROM 
	drilled_pier_profile dpp
	,foundation_details fd
	,drilled_pier_details dpd
	,structure_model sm
WHERE 
	dpp.drilled_pier_id=dpd.ID
	AND dpd.foundation_id=fd.ID
	AND fd.model_id=sm.id
	AND sm.ID=@ModelID
ORDER BY
	dpp.drilled_pier_id
	,dpp.local_drilled_pier_id