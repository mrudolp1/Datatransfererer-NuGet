[EXISTING MODEL]

SELECT
    sm.ID model_id
    ,fg.ID foundation_group_id
    ,fd.ID foundation_id
    ,pd.ID pile_id
    ,pl.pile_fnd_id
    ,pl.ID location_id
    ,pl.pile_x_coordinate
    ,pl.pile_y_coordinate

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,fnd.foundation_group fg
    ,fnd.foundation_details fd
    ,fnd.pile_details pd
    ,fnd.pile_location pl
WHERE
    smx.model_id=@ModelID
    AND smx.model_id=sm.ID
    AND sm.foundation_group_id=fg.ID
    AND fg.ID=fd.foundation_group_id
    AND fd.details_id=pd.ID
    AND pd.ID=pl.pile_fnd_id

ORDER BY
    pl.pile_fnd_id
    ,pl.ID


