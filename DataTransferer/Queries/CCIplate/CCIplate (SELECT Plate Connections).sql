[EXISTING MODEL]

SELECT
    sm.ID model_id
    ,cg.ID connection_group_id
    --,pcx.ID
    ,pc.ID connection_id
    ,pc.connection_elevation
    ,pc.connection_type
    ,pc.bolt_configuration
    ,pc.local_id

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,conn.connection_group cg
    ,conn.plate_connections_xref pcx
    ,conn.plate_connections pc
WHERE
    smx.model_id=@ModelID
    AND smx.model_id=sm.ID
    AND sm.connection_group_id=cg.ID
    AND cg.ID=pcx.connection_group_id
    AND pcx.connection_id=pc.ID

ORDER BY
    pc.connection_elevation
    ,pc.ID

