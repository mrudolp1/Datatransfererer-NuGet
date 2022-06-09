SELECT

    --fd.foundation_type --might need to remove/adjust
    --,sm.ID model_id
    --,fg.ID foundation_group_id
    --,fd.ID foundation_id
    pp.ID pp_id
    ,pp.pier_shape
    ,pp.pier_diameter
    ,pp.extension_above_grade
    ,pp.pier_rebar_size
    ,pp.pier_tie_size
    ,pp.pier_tie_quantity
    ,pp.pier_reinforcement_type
    ,pp.pier_clear_cover
    ,pp.foundation_depth
    ,pp.pad_width_1
    ,pp.pad_width_2
    ,pp.pad_thickness
    ,pp.pad_rebar_size_top_dir1
    ,pp.pad_rebar_size_bottom_dir1
    ,pp.pad_rebar_size_top_dir2
    ,pp.pad_rebar_size_bottom_dir2
    ,pp.pad_rebar_quantity_top_dir1
    ,pp.pad_rebar_quantity_bottom_dir1
    ,pp.pad_rebar_quantity_top_dir2
    ,pp.pad_rebar_quantity_bottom_dir2
    ,pp.pad_clear_cover
    ,pp.rebar_grade
    ,pp.concrete_compressive_strength
    ,pp.dry_concrete_density
    ,pp.total_soil_unit_weight
    ,pp.bearing_type
    ,pp.nominal_bearing_capacity
    ,pp.cohesion
    ,pp.friction_angle
    ,pp.spt_blow_count
    ,pp.base_friction_factor
    ,pp.neglect_depth
    ,pp.bearing_distribution_type
    ,pp.groundwater_depth
    ,pp.top_and_bottom_rebar_different
    ,pp.block_foundation
    ,pp.rectangular_foundation
    ,pp.base_plate_distance_above_foundation
    ,pp.bolt_circle_bearing_plate_width
    ,pp.pier_rebar_quantity
    ,pp.basic_soil_check
    ,pp.structural_check
    ,pp.tool_version
    --,pp.modified

FROM
    fnd.pier_pad pp
WHERE
    pp.bus_unit=[BU]
    AND pp.structure_id=[STRC ID]
