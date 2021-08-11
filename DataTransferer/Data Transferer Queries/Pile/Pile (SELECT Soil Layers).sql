[EXISTING MODEL]

SELECT
    sm.id model_id
    ,fd.ID foundation_id
    ,pd.ID pile_id
    ,sl.pile_fnd_id
    ,sl.ID soil_layer_id
    ,sl.bottom_depth
    ,sl.effective_soil_density
    ,sl.cohesion
    ,sl.friction_angle
    --,sl.skin_friction_override_uplift
    ,sl.spt_blow_count
    ,sl.ultimate_skin_friction_comp
    ,sl.ultimate_skin_friction_uplift

FROM
    foundation_details fd
    ,pile_details pd
    ,structure_model sm
    ,pile_soil_layer sl
WHERE
    sl.Pile_fnd_id=pd.ID
    AND pd.foundation_id=fd.ID
    AND fd.model_id=sm.id
    AND sm.ID=@ModelID
ORDER BY
	sl.Pile_fnd_id
	,sl.bottom_depth
