﻿	BEGIN --[CCIPLATE MATERIAL INSERT BEGIN]
	--[CCIPLATE MATERIAL INSERT]
	END --[CCIPLATE MATERIAL INSERT END]

	INSERT INTO conn.plate_details ([PLATE DETAIL FIELDS]) 
	OUTPUT INSERTED.ID INTO @SubLevel2
	VALUES([PLATE DETAIL VALUES])
	SELECT @SubLevel2ID=ID FROM @SubLevel2

	--BEGIN --[PLATE RESULTS INSERT BEGIN]
	--[PLATE RESULTS INSERT]
	--END --[PLATE RESULTS INSERT END]

	--BEGIN --[STIFFENER GROUP INSERT BEGIN]
	--[STIFFENER GROUP INSERT]
	--END --[STIFFENER GROUP INSERT END]

	--[PLATE DETAIL INSERT]