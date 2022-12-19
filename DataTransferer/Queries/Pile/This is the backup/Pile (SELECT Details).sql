﻿[EXISTING MODEL]

SELECT


    fd.foundation_type --might need to remove/adjust
    ,sm.ID model_id
    ,fg.ID foundation_group_id
    ,fd.ID foundation_id
    ,pd.ID pile_id
    ,pd.load_eccentricity
    ,pd.bolt_circle_bearing_plate_width
    ,pd.pile_shape
    ,pd.pile_material
    ,pd.pile_length
    ,pd.pile_diameter_width
    ,pd.pile_pipe_thickness
    ,pd.pile_soil_capacity_given
    ,pd.steel_yield_strength
    ,pd.pile_type_option
    ,pd.rebar_quantity
    ,pd.pile_group_config
    ,pd.foundation_depth
    ,pd.pad_thickness
    ,pd.pad_width_dir1
    ,pd.pad_width_dir2
    ,pd.pad_rebar_size_bottom
    ,pd.pad_rebar_size_top
    ,pd.pad_rebar_quantity_bottom_dir1
    ,pd.pad_rebar_quantity_top_dir1
    ,pd.pad_rebar_quantity_bottom_dir2
    ,pd.pad_rebar_quantity_top_dir2
    ,pd.pier_shape
    ,pd.pier_diameter
    ,pd.extension_above_grade
    ,pd.pier_rebar_size
    ,pd.pier_rebar_quantity
    ,pd.pier_tie_size
    ,pd.rebar_grade
    ,pd.concrete_compressive_strength
    ,pd.groundwater_depth
    ,pd.total_soil_unit_weight
    ,pd.cohesion
    ,pd.friction_angle
    ,pd.neglect_depth
    ,pd.spt_blow_count
    ,pd.pile_negative_friction_force
    ,pd.pile_ultimate_compression
    ,pd.pile_ultimate_tension
    ,pd.top_and_bottom_rebar_different
    ,pd.ultimate_gross_end_bearing
    ,pd.skin_friction_given
    ,pd.pile_quantity_circular
    ,pd.group_diameter_circular
    ,pd.pile_column_quantity
    ,pd.pile_row_quantity
    ,pd.pile_columns_spacing
    ,pd.pile_row_spacing
    ,pd.group_efficiency_factor_given
    ,pd.group_efficiency_factor
    ,pd.cap_type
    ,pd.pile_quantity_asymmetric
    ,pd.pile_spacing_min_asymmetric
    ,pd.quantity_piles_surrounding
    ,pd.pile_cap_reference
    ,pd.tool_version
    ,pd.Soil_110
    ,pd.Structural_105

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,fnd.foundation_group fg
    ,fnd.foundation_details fd
    ,fnd.pile_details pd
WHERE
    smx.model_id=@ModelID
    AND smx.model_id=sm.ID
    AND sm.foundation_group_id=fg.ID
    AND fg.ID=fd.foundation_group_id
    AND fd.details_id=pd.ID