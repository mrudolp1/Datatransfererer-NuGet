SELECT Top 1 [INDEX] As [Index], mems.ID

FROM
    tnx.members mems
WHERE
    mems.[File] = [FILE]
    AND mems.USName = [USNAME]
	AND mems.SIName = [SINAME]
	AND mems.[Values] = [VALUES]

