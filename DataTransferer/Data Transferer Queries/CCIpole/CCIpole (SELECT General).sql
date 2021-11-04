[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,pstr.ID pole_structure_id
    ,pc.ID criteria_id

FROM
    structure_model sm
    ,pole_structure pstr
    ,pole_analysis_criteria pc
WHERE
    sm.ID=@ModelID
    AND pstr.model_id=sm.ID
    AND pc.pole_structure_id=pstr.ID

