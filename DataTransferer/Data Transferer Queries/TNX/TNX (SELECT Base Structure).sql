SELECT bs.*

FROM
    tnx.tnx tnx
	,tnx.base_structure_sections bs
WHERE
    tnx.bus_unit=[BU]
    AND tnx.structure_id=[STRC ID]
	AND bs.tnx_id = tnx.ID

ORDER BY
	bs.TowerRec ASC