[EXISTING MODEL]

--WIP
SELECT 
fg.id fnd_group_id --EDIT. place in structure model, lattice structure, Lattice section, then lattice leg detail, lattice custom capacity?
,sm.id model_id 
,fd.id fnd_detail_id 
,lrd.* 
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.anchor_block_details lrd 
WHERE 
smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = abd.id 
--AND fd.foundation_type = @FndType
AND smx.bus_unit = @BU 
AND smx.structure_id = @STR_ID
--WIP