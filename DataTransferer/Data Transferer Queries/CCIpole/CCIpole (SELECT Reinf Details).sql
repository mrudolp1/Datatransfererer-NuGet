[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,ps.ID pole_structure_id
    ,prg.ID reinf_group_id
    ,prd.ID reinf_id
    ,prd.reinf_group_id
    ,prd.pole_flat
    ,prd.horizontal_offset
    ,prd.rotation
    ,prd.note

FROM
    structure_model sm
    ,pole_structure ps
    ,pole_reinf_group prg
    ,pole_reinf_details prd
WHERE
    sm.ID=@ModelID
    AND ps.model_id=sm.ID
    AND prg.pole_structure_id=ps.ID
    AND prd.reinf_group_id=prg.ID
