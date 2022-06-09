﻿[EXISTING MODEL]

	--Drilled Pier Section 
SELECT 
fg.id fnd_group_id 
,sm.id model_id 
,fd.id fnd_detail_id 
,dpd.id drilled_pier_id
--,dps.*
,dps.id section_id
,dps.local_drilled_pier_id
,dps.pier_diameter
,dps.clear_cover
,dps.clear_cover_rebar_cage_option
,dps.tie_size
,dps.tie_spacing
,dps.bottom_elevation
,dps.local_section_id
,dps.rho_override
FROM 
gen.structure_model_xref smx 
,gen.structure_model sm 
,fnd.foundation_group fg 
,fnd.foundation_details fd 
,fnd.drilled_pier_details dpd
,fnd.drilled_pier_section dps
WHERE 
smx.model_id = @ModelID
AND smx.model_id = sm.id 
AND sm.foundation_group_id = fg.id 
AND fd.foundation_group_id = fg.id 
AND fd.details_id = dpd.id 
AND dps.drilled_pier_id = dpd.id
--AND fd.foundation_type = @FndType
--AND smx.bus_unit = @BU 
--AND smx.structure_id = @STR_ID




--SELECT
--sm.id model_id
--,fd.id foundation_id
--,s.drilled_pier_id
--,s.ID section_id
--,s.pier_diameter
--,s.clear_cover
--,s.clear_cover_rebar_cage_option
--,s.tie_size
--,s.tie_spacing
--,s.bottom_elevation
----,s.assume_min_steel_rho_override
--,s.local_section_id
----,s.local_drilled_pier_id
--,s.rho_override
--FROM
--drilled_pier_section s
--,foundation_details fd
--,drilled_pier_details dpd
--,structure_model sm
--WHERE
--s.drilled_pier_id=dpd.ID
--AND dpd.foundation_id=fd.ID
--AND fd.model_id=sm.id
--AND sm.ID=@ModelID
--ORDER BY
--s.drilled_pier_id
----,s.top_elevation
--,s.bottom_elevation
