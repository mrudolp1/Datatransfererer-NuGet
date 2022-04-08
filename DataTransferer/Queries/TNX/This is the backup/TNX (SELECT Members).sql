SELECT mems.*

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,tnx.tnx_structure ts
    ,tnx.members mems
	,tnx.members_xref memx
WHERE
    smx.bus_unit=[BU]
    AND smx.structure_id=[STRC ID]
	AND smx.isActive = 1
	AND sm.ID = smx.model_id
    AND ts.ID = sm.tnx_id
    AND memx.tnx_structure_id = ts.ID
    AND mems.ID = memx.member_id
