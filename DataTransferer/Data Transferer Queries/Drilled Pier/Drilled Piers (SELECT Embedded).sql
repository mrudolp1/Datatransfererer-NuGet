[EXISTING MODEL]

	--Drilled Pier Embedded Details 
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
,dpd.id drilled_pier_id
--,e.* 
,e.id embedded_id
,e.local_drilled_pier_id
,e.embedded_pole_option
,e.encased_in_concrete
,e.pole_side_quantity
,e.pole_yield_strength
,e.pole_thickness
,e.embedded_pole_input_type
,e.pole_diameter_toc
,e.pole_top_diameter
,e.pole_bottom_diameter
,e.pole_section_length
,e.pole_taper_factor
,e.pole_bend_radius_override
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.drilled_pier_details dpd
,fnd.embedded_pole_details e
WHERE 
smx.model_id = @ModelID
AND smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = dpd.id 
AND e.drilled_pier_id = dpd.id
--AND fd.foundation_type = @FndType
--AND smx.bus_unit = @BU 
--AND smx.structure_id = @STR_ID


--SELECT 
--	sm.id model_id
--	,fd.ID foundation_id
--	,dpd.ID drilled_pier_id
--	,ep.ID embedded_id
--    ,ep.embedded_pole_option
--    ,ep.encased_in_concrete
--    ,ep.pole_side_quantity
--    ,ep.pole_yield_strength
--    ,ep.pole_thickness
--    ,ep.embedded_pole_input_type
--    ,ep.pole_diameter_toc
--    ,ep.pole_top_diameter
--    ,ep.pole_bottom_diameter
--    ,ep.pole_section_length
--    ,ep.pole_taper_factor
--    ,ep.pole_bend_radius_override
--	--,ep.local_drilled_pier_id
--FROM 
--	foundation_details fd
--	,drilled_pier_details dpd 
--	,structure_model sm
--	,embedded_pole_details ep
--WHERE 
--	dpd.foundation_id=fd.ID
--	AND fd.model_id=sm.id
--	AND ep.drilled_pier_id=dpd.ID
--	AND sm.ID=@Model