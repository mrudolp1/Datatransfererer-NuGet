[EXISTING MODEL]

	--Drilled Pier Rebar 
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
,dpd.id drilled_pier_id
,dps.id section_id
--,dpr.* 
,dpr.id rebar_id
,dpr.local_section_id
,dpr.longitudinal_rebar_quantity
,dpr.longitudinal_rebar_size
,dpr.longitudinal_rebar_cage_diameter
,dpr.local_rebar_id
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.drilled_pier_details dpd
,fnd.drilled_pier_section dps
,fnd.drilled_pier_rebar dpr
WHERE 
smx.model_id = @ModelID
AND smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = dpd.id 
AND dps.drilled_pier_id = dpd.id
AND dpr.section_id = dps.id
--AND fd.foundation_type = @FndType
--AND smx.bus_unit = @BU 
--AND smx.structure_id = @STR_ID


--SELECT 
--	sm.id model_id
--	,fd.id foundation_id
--	--,s.drilled_pier_id
--	,r.section_id
--	,r.ID rebar_id
--	,r.longitudinal_rebar_quantity
--	,r.longitudinal_rebar_size
--	,r.longitudinal_rebar_cage_diameter
--	,r.local_rebar_id
--	--,r.local_drilled_pier_id
--	--,r.local_section_id
--FROM 
--	drilled_pier_rebar r
--	,drilled_pier_section s 
--	,foundation_details fd
--	,drilled_pier_details dpd
--	,structure_model sm
--WHERE 
--	r.section_id=s.ID 
--	AND s.drilled_pier_id=dpd.ID
--	AND dpd.foundation_id=fd.ID
--	AND fd.model_id=sm.id
--	AND sm.ID=@ModelID
--ORDER BY
--	s.drilled_pier_id
--	,s.bottom_elevation