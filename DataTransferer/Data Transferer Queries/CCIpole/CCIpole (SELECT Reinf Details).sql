[EXISTING MODEL]

SELECT
	sm.ID model_id
	,pstr.ID pole_structure_id
	,prg.ID group_id
	,prd.ID reinf_id
	,prd.local_group_id
	,prd.local_reinf_id
	,prd.pole_flat
	,prd.horizontal_offset
	,prd.rotation
	,prd.note
FROM
	gen.structure_model_xref smx
	,gen.structure_model sm
	,pole.pole_structure pstr
	,pole.pole_reinf_group_xref prgx
	,pole.pole_reinf_group prg
	,pole.pole_reinf_details_xref prdx
	,pole.pole_reinf_details prd
WHERE
	smx.model_id=@ModelID
	AND smx.model_id=sm.ID
	AND sm.pole_structure_id=pstr.ID
	AND prgx.pole_structure_id = pstr.ID
	AND prg.ID = prgx.reinf_group_id
	AND prdx.reinf_group_id = prg.ID
	AND prd.ID = prdx.reinf_id
