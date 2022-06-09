[EXISTING MODEL]

SELECT
	sm.ID model_id
	,pstr.ID pole_structure_id
	,pig.ID group_id
	,pid.ID int_id
	,pid.local_group_id
	,pid.local_int_id
	,pid.pole_flat
	,pid.horizontal_offset
	,pid.rotation
	,pid.note
FROM
	gen.structure_model_xref smx
	,gen.structure_model sm
	,pole.pole_structure pstr
	,pole.pole_interference_group_xref pigx
	,pole.pole_interference_group pig
	,pole.pole_interference_details_xref pidx
	,pole.pole_interference_details pid
WHERE
	smx.model_id=@ModelID
	AND smx.model_id=sm.ID
	AND sm.pole_structure_id=pstr.ID
	AND pigx.pole_structure_id = pstr.ID
	AND pig.ID = pigx.interference_group_id
	AND pidx.interference_group_id = pig.ID
	AND pid.ID = pidx.interference_id

