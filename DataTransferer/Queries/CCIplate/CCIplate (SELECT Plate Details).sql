[EXISTING MODEL]

SELECT
    --smx.ID 
    sm.ID model_id
    ,cg.ID connection_group_id
    --,pcx.ID 
    ,pc.ID connection_id
    ,pd.ID plate_id
    ,pd.plate_location
    ,pd.plate_type
    ,pd.plate_diameter
    ,pd.plate_thickness
    ,pd.plate_material
    ,pd.stiffener_configuration
    ,pd.stiffener_clear_space
    ,pd.plate_check
    ,pd.local_id

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,conn.connection_group cg
    ,conn.plate_connections_xref pcx
    ,conn.plate_connections pc
    ,conn.plate_details pd
WHERE
    smx.model_id=@ModelID
    AND smx.model_id=sm.ID
    AND sm.connection_group_id=cg.ID
    AND cg.ID=pcx.connection_group_id
    AND pcx.connection_id=pc.ID
    AND pc.ID=pd.connection_id

ORDER BY
    pd.connection_id
    ,pd.ID
