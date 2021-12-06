SELECT *

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,tnx.tnx_structure ts
	,tnx.tnx_individual_inputs ti
WHERE
    smx.bus_unit=[BU]
    AND smx.structure_id=[STRC ID]
	AND smx.isActive = 1
	AND sm.ID = smx.model_id
    AND ts.ID = sm.tnx_id
    AND ti.ID = ts.tnx_ind_input_id
