SELECT bs.*

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,tnx.tnx_structure ts
    ,tnx.base_structure bs
	,tnx.base_structure_xref bsx
WHERE
    smx.bus_unit=[BU]
    AND smx.structure_id=[STRC ID]
	AND smx.isActive = 1
	AND sm.ID = smx.model_id
    AND ts.ID = sm.tnx_id
    AND bsx.tnx_structure_id = ts.ID
    AND bs.ID = bsx.base_structure_id

ORDER BY
	bs.TowerRec ASC
