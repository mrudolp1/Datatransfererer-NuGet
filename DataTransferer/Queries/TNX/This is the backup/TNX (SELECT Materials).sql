SELECT mats.*

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,tnx.tnx_structure ts
    ,tnx.materials mats
	,tnx.materials_xref matx
WHERE
    smx.bus_unit=[BU]
    AND smx.structure_id=[STRC ID]
	AND smx.isActive = 1
	AND sm.ID = smx.model_id
    AND ts.ID = sm.tnx_id
    AND matx.tnx_structure_id = ts.ID
    AND mats.ID = matx.material_id

