[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,ps.ID pole_structure_id
    ,prg.ID reinf_group_id
    ,prg.pole_structure_id
    ,prg.elev_bot_actual
    ,prg.elev_bot_eff
    ,prg.elev_top_actual
    ,prg.elev_top_eff
    ,prg.reinf_db_id

FROM
    structure_model sm
    ,pole_structure ps
    ,pole_reinf_group prg
WHERE
    sm.ID=@ModelID
    AND ps.model_id=sm.ID
    AND prg.pole_structure_id=ps.ID
