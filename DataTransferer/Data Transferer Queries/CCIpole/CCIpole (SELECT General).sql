[EXISTING MODEL]

SELECT
    smx.bus_unit
    ,smx.structure_id str_id
    ,sm.ID model_id
    ,pstr.ID pole_structure_id
    ,pc.ID criteria_id

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,pole.pole_structure pstr
    ,pole.pole_analysis_criteria pc
WHERE
    smx.model_id=@ModelID
    AND smx.model_id=sm.ID
    AND pstr.ID=sm.pole_structure_id
    AND pc.ID=pstr.criteria_id

