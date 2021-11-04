[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,ps.ID pole_structure_id
    ,pm.ID matl_db_id
    ,pm.ID matl_id
    ,pm.pole_structure_id
    ,pm.name
    ,pm.fy
    ,pm.fu

FROM
    structure_model sm
    ,pole_structure ps
    ,matl_prop_flat_plate pm
WHERE
    sm.ID=@ModelID
    AND ps.model_id=sm.ID
    AND pm.pole_structure_id=ps.ID
