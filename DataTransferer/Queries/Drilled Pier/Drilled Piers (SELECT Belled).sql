[EXISTING MODEL]

--Drilled Pier Belled Details 
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
,dpd.id drilled_pier_id
--,bp.*
,bp.ID belled_pier_id
,bp.local_drilled_pier_id
,bp.belled_pier_option
,bp.bottom_diameter_of_bell
,bp.bell_input_type
,bp.bell_angle
,bp.bell_height
,bp.bell_toe_height
,bp.neglect_top_soil_layer
,bp.swelling_expansive_soil
,bp.depth_of_expansive_soil
,bp.expansive_soil_force
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.drilled_pier_details dpd
,fnd.belled_pier_details bp
WHERE 
smx.model_id = @ModelID
AND smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = dpd.id 
AND bp.drilled_pier_id = dpd.id
--AND fd.foundation_type = @FndType
--AND smx.bus_unit = @BU 
--AND smx.structure_id = @STR_ID


--SELECT 
--	sm.id model_id
--	,fd.ID foundation_id
--	,dpd.ID drilled_pier_id
--	,bp.ID belled_pier_id
--    ,bp.belled_pier_option
--    ,bp.bottom_diameter_of_bell
--    ,bp.bell_input_type
--    ,bp.bell_angle
--    ,bp.bell_height
--    ,bp.bell_toe_height
--    ,bp.neglect_top_soil_layer
--    ,bp.swelling_expansive_soil
--    ,bp.depth_of_expansive_soil
--    ,bp.expansive_soil_force
--	--,bp.local_drilled_pier_id
--FROM 
--	foundation_details fd
--	,drilled_pier_details dpd 
--	,structure_model sm
--	,belled_pier_details bp
--WHERE 
--	dpd.foundation_id=fd.ID
--	AND fd.model_id=sm.id
--	AND bp.drilled_pier_id=dpd.ID
--	AND sm.ID=@ModelID