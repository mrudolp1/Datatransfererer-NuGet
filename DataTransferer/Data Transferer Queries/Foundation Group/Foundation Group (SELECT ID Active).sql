SELECT
    sm.foundation_group_id
FROM
    gen.Structure_model_xref smx
    ,gen.structure_model sm
WHERE
    smx.bus_unit=[BU]
    AND smx.structure_id=[STRC ID]
	AND smx.isActive = 1
	AND sm.ID = smx.model_id