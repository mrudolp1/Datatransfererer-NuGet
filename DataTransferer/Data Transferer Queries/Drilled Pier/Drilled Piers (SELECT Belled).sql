[EXISTING MODEL]

SELECT 
	sm.id model_id
	,fd.ID foundation_id
	,dpd.ID drilled_pier_id
	,bp.ID belled_pier_id
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
	--,bp.local_drilled_pier_id
FROM 
	foundation_details fd
	,drilled_pier_details dpd 
	,structure_model sm
	,belled_pier_details bp
WHERE 
	dpd.foundation_id=fd.ID
	AND fd.model_id=sm.id
	AND bp.drilled_pier_id=dpd.ID
	AND sm.ID=@ModelID