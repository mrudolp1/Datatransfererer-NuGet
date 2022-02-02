[EXISTING MODEL]

SELECT
    sm.ID model_id
    ,pstr.ID pole_structure_id

    ,pig.ID interference_group_id
    ,pid.ID interference_id
    ,pid.interference_group_id
    ,pid.pole_flat
    ,pid.horizontal_offset
    ,pid.rotation
    ,pid.note

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,pole.pole_structure pstr
    ,pole.pole_interference_group pig
    ,pole.pole_interference_details pid
WHERE
    smx.model_id=@ModelID
    AND smx.model_id=sm.ID
    AND sm.pole_structure_id=pstr.ID

    AND pig.pole_structure_id=pstr.ID
    AND pid.interference_group_id=pig.ID

