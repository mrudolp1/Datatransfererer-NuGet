[EXISTING MODEL]

SELECT
    sm.ID model_id
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

    --Do I need to get rid of 'pc.ID criteria_id'?
FROM
     gen.structure_model_xref smx
    ,gen.structure_model sm
    ,pole.pole_structure pstr
    ,pole.pole_analysis_criteria pc
WHERE
    smx.model_id=@ModelID
    AND smx.model_id=sm.ID
    AND sm.pole_structure_id=pstr.ID
    AND pstr.criteria_id=pc.ID

