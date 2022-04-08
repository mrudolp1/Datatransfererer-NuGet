[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,pstr.ID pole_structure_id
    ,psec.ID section_id
    ,psec.pole_structure_id
    ,psec.analysis_section_id
    ,psec.elev_bot
    ,psec.elev_top
    ,psec.length_section
    ,psec.length_splice
    ,psec.num_sides
    ,psec.diam_bot
    ,psec.diam_top
    ,psec.wall_thickness
    ,psec.bend_radius
    ,psec.steel_grade_id
    ,psec.pole_type
    ,psec.section_name
    ,psec.socket_length
    ,psec.weight_mult
    ,psec.wp_mult
    ,psec.af_factor
    ,psec.ar_factor
    ,psec.round_area_ratio
    ,psec.flat_area_ratio

FROM
    structure_model sm
    ,pole_structure pstr
    ,pole_section psec
WHERE
    sm.ID=@ModelID
    AND pstr.model_id=sm.ID
    AND psec.pole_structure_id=pstr.ID
