﻿	INSERT INTO conn.plates ([CONNECTION FIELDS]) 
	OUTPUT INSERTED.ID INTO @SubLevel1
	VALUES([CONNECTION VALUES])
	SELECT @SubLevel1ID=ID FROM @SubLevel1

	--BEGIN --[PLATE DETAIL INSERT BEGIN]
	--[PLATE DETAIL INSERT]
	--END --[PLATE DETAIL INSERT END]

	--[CONNECTION INSERT]