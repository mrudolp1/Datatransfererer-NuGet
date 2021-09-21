[EXISTING MODEL]

SELECT
    sm.bus_unit
    ,sm.structure_id str_id
    ,sm.ID model_id
    ,pstr.ID pole_structure_id
    ,pc.ID criteria_id
    ,pc.pole_structure_id
    ,pc.upper_structure_type
    ,pc.analysis_deg
    ,pc.geom_increment_length
    ,pc.vnum
    ,pc.check_connections
    ,pc.hole_deformation
    ,pc.ineff_mod_check
    ,pc.modified

FROM
    structure_model sm
    ,pole_structure pstr
    ,pole_analysis_criteria pc
WHERE
    sm.ID=@ModelID
    AND pstr.model_id=sm.ID
    AND pc.pole_structure_id=pstr.ID

