    [EXISTING MODEL]

SELECT

    --,smx.ID 
    sm.ID model_id
    ,cg.ID connection_group_id
    --,pcx.ID 
    ,pc.ID connection_id
    ,bpo.ID base_plate_options_id
    ,bpo.anchor_rod_spacing
    ,bpo.clip_distance
    ,bpo.barb_cl_elevation
    ,bpo.include_pole_reactions
    ,bpo.consider_ar_eccentricity
    ,bpo.leg_mod_eccentricity

FROM
    gen.structure_model_xref smx
    ,gen.structure_model sm
    ,conn.connection_group cg
    ,conn.plate_connections_xref pcx
    ,conn.plate_connections pc
    ,conn.base_plate_options bpo
WHERE
    smx.model_id=@ModelID
    AND smx.model_id=sm.ID
    AND sm.connection_group_id=cg.ID
    AND cg.ID=pcx.connection_group_id
    AND pcx.connection_id=pc.ID
    AND pc.ID=bpo.connection_id

