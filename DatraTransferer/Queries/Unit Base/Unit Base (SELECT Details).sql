[EXISTING MODEL]

SELECT 
	sm.bus_unit
	,sm.structure_id str_id
	,sm.id model_id
	,fd.ID foundation_id
	,fd.foundation_type 	
	,ub.ID unit_base_id
	,ub.extension_above_grade
	,ub.foundation_depth
	,ub.concrete_compressive_strength
	,ub.dry_concrete_density
	,ub.rebar_grade
	,ub.top_and_bottom_rebar_different
	,ub.block_foundation
	,ub.rectangular_foundation
	,ub.base_plate_distance_above_foundation
	,ub.bolt_circle_bearing_plate_width
	,ub.tower_centroid_offset
	,ub.pier_shape
	,ub.pier_diameter
	,ub.pier_rebar_quantity
	,ub.pier_rebar_size
	,ub.pier_tie_quantity
	,ub.pier_tie_size
	,ub.pier_reinforcement_type
	,ub.pier_clear_cover
	,ub.pad_width_1
	,ub.pad_width_2
	,ub.pad_thickness
	,ub.pad_rebar_size_top_dir1
	,ub.pad_rebar_size_bottom_dir1
	,ub.pad_rebar_size_top_dir2
	,ub.pad_rebar_size_bottom_dir2
	,ub.pad_rebar_quantity_top_dir1
	,ub.pad_rebar_quantity_bottom_dir1
	,ub.pad_rebar_quantity_top_dir2
	,ub.pad_rebar_quantity_bottom_dir2
	,ub.pad_clear_cover
	,ub.total_soil_unit_weight
	,ub.bearing_type
	,ub.nominal_bearing_capacity
	,ub.cohesion
	,ub.friction_angle
	,ub.spt_blow_count
	,ub.base_friction_factor
	,ub.neglect_depth
	,ub.bearing_distribution_type
	,ub.groundwater_depth
	,ub.tower_centroid_offset
FROM 
	foundation_details fd
	,unit_base_details ub 
	,structure_model sm
WHERE 
	ub.foundation_id=fd.ID
	AND fd.model_id=sm.id
	AND sm.ID=@ModelID