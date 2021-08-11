[EXISTING MODEL]

SELECT
    sm.id model_id
    ,fd.ID foundation_id
    ,pd.ID pile_id
    ,pl.pile_fnd_id
    ,pl.ID location_id
    ,pl.pile_x_coordinate
    ,pl.pile_y_coordinate

FROM
    foundation_details fd
    ,pile_details pd
    ,structure_model sm
    ,pile_location pl
WHERE
    pl.Pile_fnd_id=pd.ID
    AND pd.foundation_id=fd.ID
    AND fd.model_id=sm.id
    AND sm.ID=@ModelID
ORDER BY
	pl.Pile_fnd_id
	,pl.ID