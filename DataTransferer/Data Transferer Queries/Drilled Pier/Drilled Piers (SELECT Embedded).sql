[EXISTING MODEL]

	--Drilled Pier Embedded Details 
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
,e.* 
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.embedded_pole_details e
WHERE 
smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = e.id 
--AND fd.foundation_type = @FndType
AND smx.bus_unit = @BU 
AND smx.structure_id = @STR_ID


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