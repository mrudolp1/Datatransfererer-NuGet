﻿	BEGIN --[CCIPLATE MATERIAL INSERT BEGIN]
	--[CCIPLATE MATERIAL INSERT]
	END --[CCIPLATE MATERIAL INSERT END]	

	INSERT INTO conn.bridge_stiffeners ([BRIDGE STIFFENER DETAIL FIELDS]) 
	OUTPUT INSERTED.ID INTO @SubLevel2
	VALUES([BRIDGE STIFFENER DETAIL VALUES])
	SELECT @SubLevel2ID=ID FROM @SubLevel2

	--[BRIDGE STIFFENER DETAIL INSERT]