[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,ps.ID pole_structure_id
    ,pb.ID bolt_db_id
    ,pb.ID bolt_id
    ,pb.pole_structure_id
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
    structure_model sm
    ,pole_structure ps
    ,bolt_prop_flat_plate pb
WHERE
    sm.ID=@ModelID
    AND ps.model_id=sm.ID
    AND pb.pole_structure_id=ps.ID
