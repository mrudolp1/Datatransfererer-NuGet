
    [EXISTING MODEL]

SELECT
    sm.ID model_id
    ,fg.ID foundation_group_id
    ,fd.ID foundation_id
    ,pd.ID pile_id
    ,sl.pile_fnd_id
    ,sl.ID soil_layer_id
    ,sl.bottom_depth
    ,sl.effective_soil_density
    ,sl.cohesion
    ,sl.friction_angle
    ,sl.spt_blow_count
    ,sl.ultimate_skin_friction_comp
    ,sl.ultimate_skin_friction_uplift

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,fnd.foundation_group fg
    ,fnd.foundation_details fd
    ,fnd.pile_details pd
    ,fnd.pile_soil_layer sl
WHERE
    smx.model_id=@ModelID
    AND smx.model_id=sm.ID
    AND sm.foundation_group_id=fg.ID
    AND fg.ID=fd.foundation_group_id
    AND fd.details_id=pd.ID
    AND pd.ID=sl.pile_fnd_id

ORDER BY
    sl.pile_fnd_id
    ,sl.ID
     --Filtering by bottom depth doesn't work if empty rows exist in SQL database
	--,sl.bottom_depth
