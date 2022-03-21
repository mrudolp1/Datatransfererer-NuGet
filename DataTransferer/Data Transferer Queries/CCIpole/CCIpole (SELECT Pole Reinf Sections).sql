[EXISTING MODEL]

SELECT
sm.ID model_id
,pstr.ID pole_structure_id
,prs.ID section_id
,prs.local_section_id
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
,prs.local_matl_id
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
gen.structure_model_xref smx
,gen.structure_model sm
,pole.pole_structure pstr
,pole.pole_reinf_section_xref prsx
,pole.pole_reinf_section prs
WHERE
smx.model_id=@ModelID
AND smx.model_id=sm.ID
AND sm.pole_structure_id=pstr.ID
AND prsx.pole_structure_id = pstr.ID
AND prs.ID = prsx.section_id
