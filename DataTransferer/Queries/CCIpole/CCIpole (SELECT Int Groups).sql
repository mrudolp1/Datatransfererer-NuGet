[EXISTING MODEL]

SELECT
	sm.ID model_id
	,pstr.ID pole_structure_id
	,pig.ID group_id
	,pig.local_group_id
	,pig.elev_bot
	,pig.elev_top
	,pig.width
	,pig.description
	,pig.qty
FROM
	gen.structure_model_xref smx
	,gen.structure_model sm
	,pole.pole_structure pstr
	,pole.pole_interference_group_xref pigx
	,pole.pole_interference_group pig
WHERE
	smx.model_id=@ModelID
	AND smx.model_id=sm.ID
	AND sm.pole_structure_id=pstr.ID
	AND pigx.pole_structure_id = pstr.ID
	AND pig.ID = pigx.interference_group_id