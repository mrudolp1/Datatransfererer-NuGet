SELECT 
    --fd.ID foundation_detail_id
    --,fd.foundation_group_id
    --,fd.foundation_type
    --,fd.guy_group_id
    ub.ID unit_base_id
    ,ub.pier_shape
    ,ub.pier_diameter
    ,ub.extension_above_grade
    ,ub.pier_rebar_size
    ,ub.pier_tie_size
    ,ub.pier_tie_quantity
    ,ub.pier_reinforcement_type
    ,ub.pier_clear_cover
    ,ub.foundation_depth
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
    ,ub.rebar_grade
    ,ub.concrete_compressive_strength
    ,ub.dry_concrete_density
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
    ,ub.top_and_bottom_rebar_different
    ,ub.block_foundation
    ,ub.rectangular_foundation
    ,ub.base_plate_distance_above_foundation
    ,ub.bolt_circle_bearing_plate_width
    ,ub.tower_centroid_offset
    ,ub.pier_rebar_quantity
    ,ub.basic_soil_check
    ,ub.structural_check
    ,ub.tool_version
    ,ub.modified
FROM
    --fnd.foundation_details fd
    --,fnd.pier_pad_details det
    fnd.unit_base ub
WHERE
    --fd.foundation_group_id=[FNDGRPID]
    --AND fd.foundation_type ='Unit Base'
    --AND det.ID=fd.details_id
    ub.bus_unit=[BU]
    AND ub.structure_id=[STRC ID]
