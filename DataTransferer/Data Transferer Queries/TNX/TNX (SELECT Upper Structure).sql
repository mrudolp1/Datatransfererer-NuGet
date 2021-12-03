SELECT us.*

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,tnx.tnx_structure ts
    ,tnx.upper_structure us
	,tnx.upper_structure_xref usx
WHERE
    smx.bus_unit=[BU]
    AND smx.structure_id=[STRC ID]
	AND smx.isActive = 1
	AND sm.ID = smx.model_id
    AND ts.ID = sm.tnx_id
    AND usx.tnx_structure_id = ts.ID
    AND us.ID = usx.upper_structure_id

ORDER BY
	us.AntennaRec ASC
