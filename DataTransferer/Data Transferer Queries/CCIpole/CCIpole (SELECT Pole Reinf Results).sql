[EXISTING MODEL]

SELECT
	sm.ID model_id
	,pstr.ID pole_structure_id
	,prr.work_order_seq_num
	,prr.ID section_id
	,prr.reinf_group_id
	,prr.local_section_id
	,prr.local_reinf_group_id
	,prr.result_lkup_value
	,prr.rating
FROM
	gen.structure_model_xref smx
	,gen.structure_model sm
	,pole.pole_structure pstr
	,gen.model_work_order_xref mwox
	,pole.pole_reinf_results prr
WHERE
	smx.model_id=@ModelID
	AND smx.model_id=sm.ID
	AND sm.pole_structure_id=pstr.ID
	AND mwox.model_id = sm.ID
	AND prr.work_order_seq_num = mwox.work_order_seq_num
