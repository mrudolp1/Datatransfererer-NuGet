SELECT Top 1 [INDEX] As [Index], mats.ID

FROM
    tnx.materials mats
WHERE
    mats.MemberMatFile = [MEMBERMATFILE]
    AND mats.MatName = [MATNAME]
	AND mats.MatValues = [MATVALUES]
	AND mats.IsBolt = [ISBOLT]

