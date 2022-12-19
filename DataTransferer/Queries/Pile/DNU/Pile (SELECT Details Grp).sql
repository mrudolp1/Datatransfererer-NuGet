﻿SELECT 
    fd.ID foundation_detail_id
    ,fd.foundation_group_id
    ,fd.foundation_type
    ,fd.guy_group_id
    ,det.ID foundation_id
    ,det.load_eccentricity
    ,det.bolt_circle_bearing_plate_width
    ,det.pile_shape
    ,det.pile_material
    ,det.pile_length
    ,det.pile_diameter_width
    ,det.pile_pipe_thickness
    ,det.pile_soil_capacity_given
    ,det.steel_yield_strength
    ,det.pile_type_option
    ,det.rebar_quantity
    ,det.pile_group_config
    ,det.foundation_depth
    ,det.pad_thickness
    ,det.pad_width_dir1
    ,det.pad_width_dir2
    ,det.pad_rebar_size_bottom
    ,det.pad_rebar_size_top
    ,det.pad_rebar_quantity_bottom_dir1
    ,det.pad_rebar_quantity_top_dir1
    ,det.pad_rebar_quantity_bottom_dir2
    ,det.pad_rebar_quantity_top_dir2
    ,det.pier_shape
    ,det.pier_diameter
    ,det.extension_above_grade
    ,det.pier_rebar_size
    ,det.pier_rebar_quantity
    ,det.pier_tie_size
    ,det.rebar_grade
    ,det.concrete_compressive_strength
    ,det.groundwater_depth
    ,det.total_soil_unit_weight
    ,det.cohesion
    ,det.friction_angle
    ,det.neglect_depth
    ,det.spt_blow_count
    ,det.pile_negative_friction_force
    ,det.pile_ultimate_compression
    ,det.pile_ultimate_tension
    ,det.top_and_bottom_rebar_different
    ,det.ultimate_gross_end_bearing
    ,det.skin_friction_given
    ,det.pile_quantity_circular
    ,det.group_diameter_circular
    ,det.pile_column_quantity
    ,det.pile_row_quantity
    ,det.pile_columns_spacing
    ,det.pile_row_spacing
    ,det.group_efficiency_factor_given
    ,det.group_efficiency_factor
    ,det.cap_type
    ,det.pile_quantity_asymmetric
    ,det.pile_spacing_min_asymmetric
    ,det.quantity_piles_surrounding
    ,det.pile_cap_reference
    ,det.Soil_110
    ,det.Structural_105
    ,det.tool_version
    --,det.modified

FROM
    fnd.foundation_details fd
    ,fnd.pile_details det
WHERE
    fd.foundation_group_id=[FNDGRPID]
    AND fd.foundation_type ='Pile'
    AND det.ID=fd.details_id