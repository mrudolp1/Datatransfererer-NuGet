[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,ps.ID pole_structure_id
    ,pig.ID interference_group_id
    ,pid.ID interference_id
    ,pid.interference_group_id
    ,pid.pole_flat
    ,pid.horizontal_offset
    ,pid.rotation
    ,pid.note

FROM
    structure_model sm
    ,pole_structure ps
    ,pole_interference_group pig
    ,pole_interference_details pid
WHERE
    sm.ID=@ModelID
    AND ps.model_id=sm.ID
    AND pig.pole_structure_id=ps.ID
    AND pid.interference_group_id=pig.ID

