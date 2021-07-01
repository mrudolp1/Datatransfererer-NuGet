[EXISTING MODEL]

SELECT 
	sm.id model_id
	,fd.ID foundation_id
	,dpd.ID drilled_pier_id
	,ep.ID embedded_id
	,eps.ID embedded_section_ID
    ,eps.pier_diameter
FROM 
	foundation_details fd
	,drilled_pier_details dpd 
	,structure_model sm
	,embedded_pole_details ep
	,embedded_pole_section eps
WHERE 
	dpd.foundation_id=fd.ID
	AND fd.model_id=sm.id
	AND ep.drilled_pier_id=dpd.ID
	AND eps.embedded_pier_id=ep.ID
	AND sm.ID=@ModelID