[EXISTING MODEL]

SELECT 
	sm.id model_id
	,fd.id foundation_id
	,s.drilled_pier_id
	,s.ID section_id
	,s.pier_diameter
	,s.clear_cover
	,s.clear_cover_rebar_cage_option
	,s.tie_size
	,s.tie_spacing
	,s.top_elevation
	,s.bottom_elevation
	,s.tie_yield_strength
	,s.concrete_compressive_strength
	,s.assum_min_steel_rho_override
FROM 
	drilled_pier_section s 
	,foundation_details fd
	,drilled_pier_details dpd
	,structure_model sm
WHERE 
	s.drilled_pier_id=dpd.ID
	AND dpd.foundation_id=fd.ID
	AND fd.model_id=sm.id
	AND sm.ID=@ModelID
ORDER BY
	s.drilled_pier_id
	,s.top_elevation
	,s.bottom_elevation
