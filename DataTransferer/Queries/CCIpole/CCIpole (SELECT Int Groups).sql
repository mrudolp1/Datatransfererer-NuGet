[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,ps.ID pole_structure_id
    ,pig.ID interference_group_id
    ,pig.pole_structure_id
    ,pig.elev_bot
    ,pig.elev_top
    ,pig.width
    ,pig.description

FROM
    structure_model sm
    ,pole_structure ps
    ,pole_interference_group pig
WHERE
    sm.ID=@ModelID
    AND ps.model_id=sm.ID
    AND pig.pole_structure_id=ps.ID
