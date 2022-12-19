[EXISTING MODEL]

SELECT
sm.ID model_id
,pstr.ID pole_structure_id
,psec.ID section_id
,psec.local_section_id
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
,psec.local_matl_id
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
gen.structure_model_xref smx
,gen.structure_model sm
,pole.pole_structure pstr
,pole.pole_section_xref psecx
,pole.pole_section psec
WHERE
smx.model_id=@ModelID
AND smx.model_id=sm.ID
AND sm.pole_structure_id=pstr.ID
AND psecx.pole_structure_id = pstr.ID
AND psec.ID = psecx.section_id