[EXISTING MODEL]

SELECT 
	sm.bus_unit
	,sm.structure_id str_id
	,sm.id model_id
	,fd.ID foundation_id
	,dpd.ID drilled_pier_id
	,fd.foundation_type 	
	,dpd.foundation_depth
	,dpd.extension_above_grade
	,dpd.groundwater_depth
	,dpd.assume_min_steel
	,dpd.check_shear_along_depth
	,dpd.utilize_skin_friction_methodology
	,dpd.embedded_pole
	,dpd.belled_pier
	,dpd.soil_layer_quantity
FROM 
	foundation_details fd
	,drilled_pier_details dpd 
	,structure_model sm
WHERE 
	dpd.foundation_id=fd.ID
	AND fd.model_id=sm.id
	AND sm.ID=@ModelID