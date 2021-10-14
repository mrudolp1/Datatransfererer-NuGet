[EXISTING MODEL]

--Drilled Pier Profile 
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
,dpp.* 
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.drilled_pier_profile dpp
WHERE 
smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = dpp.id 
--AND fd.foundation_type = @FndType
AND smx.bus_unit = @BU 
AND smx.structure_id = @STR_ID


--SELECT 
--	sm.id model_id
--	,fd.id foundation_id
--	,dpp.drilled_pier_id
--	,dpp.ID profile_id
--	--,dpp.local_drilled_pier_id
--	,dpp.reaction_position
--	,dpp.reaction_location
--	,dpp.drilled_pier_profile
--	,dpp.soil_profile
--FROM 
--	drilled_pier_profile dpp
--	,foundation_details fd
--	,drilled_pier_details dpd
--	,structure_model sm
--WHERE 
--	dpp.drilled_pier_id=dpd.ID
--	AND dpd.foundation_id=fd.ID
--	AND fd.model_id=sm.id
--	AND sm.ID=@ModelID
--ORDER BY
--	dpp.drilled_pier_id
--	--,dpp.local_drilled_pier_id