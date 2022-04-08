[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,ps.ID pole_structure_id
    ,prr.ID section_id
    ,prr.pole_structure_id
    ,prr.work_order_seq_num
    ,prr.reinf_group_id
    ,prr.result_lkup_value
    ,prr.rating

FROM
    structure_model sm
    ,pole_structure ps
    ,pole_reinf_results prr
WHERE
    sm.ID=@ModelID
    AND ps.model_id=sm.ID
    AND prr.pole_structure_id=ps.ID
