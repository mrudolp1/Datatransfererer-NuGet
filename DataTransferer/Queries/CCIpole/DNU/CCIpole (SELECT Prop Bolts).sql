[EXISTING MODEL]

SELECT
    sm.ID model_id
    ,pstr.ID pole_structure_id
    ,pb.ID bolt_db_id
    ,pb.local_id
	,pb.name
    ,pb.description
    ,pb.diam
    ,pb.area
    ,pb.fu_bolt
    ,pb.sleeve_diam_out
    ,pb.sleeve_diam_in
    ,pb.fu_sleeve
    ,pb.bolt_n_sleeve_shear_revF
    ,pb.bolt_x_sleeve_shear_revF
    ,pb.bolt_n_sleeve_shear_revG
    ,pb.bolt_x_sleeve_shear_revG
    ,pb.bolt_n_sleeve_shear_revH
    ,pb.bolt_x_sleeve_shear_revH
    ,pb.rb_applied_revH
FROM
	gen.structure_model_xref smx
	,gen.structure_model sm
	,pole.pole_structure pstr
	,pole.bolt_prop_flat_plate_xref pbx
	,pole.bolt_prop_flat_plate pb
WHERE
	smx.model_id=@ModelID
	AND smx.model_id=sm.ID
	AND sm.pole_structure_id=pstr.ID
	AND pbx.pole_structure_id = pstr.ID
	AND pb.ID = pbx.bolt_id
