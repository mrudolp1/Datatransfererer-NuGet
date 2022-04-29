SELECT 
    fd.ID foundation_detail_id
    ,fd.foundation_group_id
    ,fd.foundation_type
    ,fd.guy_group_id
    ,det.ID foundation_id
    ,det.pier_shape
    ,det.pier_diameter
    ,det.extension_above_grade
    ,det.pier_rebar_size
    ,det.pier_tie_size
    ,det.pier_tie_quantity
    ,det.pier_reinforcement_type
    ,det.pier_clear_cover
    ,det.foundation_depth
    ,det.pad_width_1
    ,det.pad_width_2
    ,det.pad_thickness
    ,det.pad_rebar_size_top_dir1
    ,det.pad_rebar_size_bottom_dir1
    ,det.pad_rebar_size_top_dir2
    ,det.pad_rebar_size_bottom_dir2
    ,det.pad_rebar_quantity_top_dir1
    ,det.pad_rebar_quantity_bottom_dir1
    ,det.pad_rebar_quantity_top_dir2
    ,det.pad_rebar_quantity_bottom_dir2
    ,det.pad_clear_cover
    ,det.rebar_grade
    ,det.concrete_compressive_strength
    ,det.dry_concrete_density
    ,det.total_soil_unit_weight
    ,det.bearing_type
    ,det.nominal_bearing_capacity
    ,det.cohesion
    ,det.friction_angle
    ,det.spt_blow_count
    ,det.base_friction_factor
    ,det.neglect_depth
    ,det.bearing_distribution_type
    ,det.groundwater_depth
    ,det.top_and_bottom_rebar_different
    ,det.block_foundation
    ,det.rectangular_foundation
    ,det.base_plate_distance_above_foundation
    ,det.bolt_circle_bearing_plate_width
    ,det.tower_centroid_offset
    ,det.pier_rebar_quantity
    ,det.basic_soil_check
    ,det.structural_check
    ,det.tool_version
    ,det.modified
FROM
    fnd.foundation_details fd
    ,fnd.unit_base_details det
WHERE
    fd.foundation_group_id=[FNDGRPID]
    AND fd.foundation_type ='Unit Base'
    AND det.ID=fd.details_id
