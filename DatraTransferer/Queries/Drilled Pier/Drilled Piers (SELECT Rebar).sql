[EXISTING MODEL]

SELECT 
	sm.id model_id
	,fd.id foundation_id
	,s.drilled_pier_id
	,r.section_id
	,r.ID rebar_id
	,r.longitudinal_rebar_quantity
	,r.longitudinal_rebar_size
	,r.longitudinal_rebar_cage_diameter
	,r.longitudinal_rebar_yield_strength
FROM 
	drilled_pier_rebar r
	,drilled_pier_section s 
	,foundation_details fd
	,drilled_pier_details dpd
	,structure_model sm
WHERE 
	r.section_id=s.ID 
	AND s.drilled_pier_id=dpd.ID
	AND dpd.foundation_id=fd.ID
	AND fd.model_id=sm.id
	AND sm.ID=@ModelID
ORDER BY
	s.drilled_pier_id
	,s.top_elevation
	,s.bottom_elevation