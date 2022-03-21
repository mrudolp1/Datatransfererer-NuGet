[EXISTING MODEL]

SELECT
	sm.ID model_id
    ,pstr.ID pole_structure_id
    ,pm.ID matl_db_id
    ,pm.local_id
    ,pm.name
    ,pm.fy
    ,pm.fu
FROM
	gen.structure_model_xref smx
	,gen.structure_model sm
	,pole.pole_structure pstr
	,pole.matl_prop_flat_plate_xref pmx
	,pole.matl_prop_flat_plate pm
WHERE
	smx.model_id=@ModelID
	AND smx.model_id=sm.ID
	AND sm.pole_structure_id=pstr.ID
	AND pmx.pole_structure_id = pstr.ID
	AND pm.ID = pmx.matl_id
