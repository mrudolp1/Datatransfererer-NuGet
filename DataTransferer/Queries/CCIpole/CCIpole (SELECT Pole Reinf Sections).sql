[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,ps.ID pole_structure_id
    ,prs.ID section_id
    ,prs.pole_structure_id
    ,prs.analysis_section_ID
    ,prs.elev_bot
    ,prs.elev_top
    ,prs.length_section
    ,prs.length_splice
    ,prs.num_sides
    ,prs.diam_bot
    ,prs.diam_top
    ,prs.wall_thickness
    ,prs.bend_radius
    ,prs.steel_grade_id
    ,prs.pole_type
    ,prs.weight_mult
    ,prs.section_name
    ,prs.socket_length
    ,prs.wp_mult
    ,prs.af_factor
    ,prs.ar_factor
    ,prs.round_area_ratio
    ,prs.flat_area_ratio

FROM
    structure_model sm
    ,pole_structure ps
    ,pole_reinf_section prs
WHERE
    sm.ID=@ModelID
    AND ps.model_id=sm.ID
    AND prs.pole_structure_id=ps.ID
