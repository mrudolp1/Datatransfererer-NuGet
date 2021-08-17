[EXISTING MODEL]

SELECT 
	sm.bus_unit
	,sm.structure_id str_id
	,sm.id model_id
	,fd.ID foundation_id
	,fd.foundation_type 	
	,ppd.id pier_pad_id
	,ppd.extension_above_grade
	,ppd.foundation_depth
	,ppd.concrete_compressive_strength
	,ppd.dry_concrete_density
	,ppd.rebar_grade
	,ppd.top_and_bottom_rebar_different
	,ppd.block_foundation
	,ppd.rectangular_foundation
	,ppd.base_plate_distance_above_foundation
	,ppd.bolt_circle_bearing_plate_width
	,ppd.pier_shape
	,ppd.pier_diameter
	,ppd.pier_rebar_quantity
	,ppd.pier_rebar_size
	,ppd.pier_tie_quantity
	,ppd.pier_tie_size
	,ppd.pier_reinforcement_type
	,ppd.pier_clear_cover
	,ppd.pad_width_1
	,ppd.pad_width_2
	,ppd.pad_thickness
	,ppd.pad_rebar_size_top_dir1
	,ppd.pad_rebar_size_bottom_dir1
	,ppd.pad_rebar_size_top_dir2
	,ppd.pad_rebar_size_bottom_dir2
	,ppd.pad_rebar_quantity_top_dir1
	,ppd.pad_rebar_quantity_bottom_dir1
	,ppd.pad_rebar_quantity_top_dir2
	,ppd.pad_rebar_quantity_bottom_dir2
	,ppd.pad_clear_cover
	,ppd.total_soil_unit_weight
	,ppd.bearing_type
	,ppd.nominal_bearing_capacity
	,ppd.cohesion
	,ppd.friction_angle
	,ppd.spt_blow_count
	,ppd.base_friction_factor
	,ppd.neglect_depth
	,ppd.bearing_distribution_type
	,ppd.groundwater_depth
	,ppd.basic_soil_check
	,ppd.structural_check

FROM 
	foundation_details fd
	,pier_pad_details ppd 
	,structure_model sm
WHERE 
	ppd.foundation_id=fd.ID
	AND fd.model_id=sm.id
	AND sm.ID=@ModelID