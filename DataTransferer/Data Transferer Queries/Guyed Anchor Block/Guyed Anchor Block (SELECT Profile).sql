[EXISTING MODEL]

--SELECT --WORK IN PROGRESS
--	sm.id model_id
--	,fd.id foundation_id
--	,gabd.ID anchor_id
--	,gabp.ID profile_id
--	,gabp.reaction_location
--	,gabp.anchor_profile
--	,gabp.soil_profile
--	,gabp.local_anchor_id
--FROM 
--	anchor_block_profile gabp
--	,foundation_details fd
--	,anchor_details gabd 
--	,structure_model sm
--WHERE 
--	gabp.anchor_id=gabd.ID
--	AND gabd.foundation_id=fd.ID
--	AND fd.model_id=sm.id
--	AND sm.ID=@ModelID
--ORDER BY
--	gabp.anchor_id
--	,gabp.local_anchor_

	--Anchor Block Profile
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
,abp.* 
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.anchor_block_details abd 
,fnd.anchor_block_profile abp 
WHERE 
smx.model_id=@ModelID
AND smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = abd.id 
AND abp.anchor_id = abd.id 
--AND fd.foundation_type = @FndType
--AND smx.bus_unit = @BU 
--AND smx.structure_id = @STR_ID