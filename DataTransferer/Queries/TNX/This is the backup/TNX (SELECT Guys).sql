SELECT guys.*

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,tnx.tnx_structure ts
    ,tnx.guys guys
	,tnx.guys_xref guyx
WHERE
    smx.bus_unit=[BU]
    AND smx.structure_id=[STRC ID]
	AND smx.isActive = 1
	AND sm.ID = smx.model_id
    AND ts.ID = sm.tnx_id
    AND guyx.tnx_structure_id = ts.ID
    AND guys.ID = guyx.guy_id

ORDER BY
	guys.GuyRec ASC
