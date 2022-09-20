[EXISTING MODEL]

SELECT
	sm.ID model_id
	,pstr.ID pole_structure_id
	,prg.ID group_id
	,prg.local_group_id
	,prg.elev_bot_actual
	,prg.elev_bot_eff
	,prg.elev_top_actual
	,prg.elev_top_eff
	,prg.reinf_db_id
	,prg.local_reinf_id
	,prg.qty
FROM
	gen.structure_model_xref smx
	,gen.structure_model sm
	,pole.pole_structure pstr
	,pole.pole_reinf_group_xref prgx
	,pole.pole_reinf_group prg
WHERE
	smx.model_id=@ModelID
	AND smx.model_id=sm.ID
	AND sm.pole_structure_id=pstr.ID
	AND prgx.pole_structure_id = pstr.ID
	AND prg.ID = prgx.reinf_group_id
