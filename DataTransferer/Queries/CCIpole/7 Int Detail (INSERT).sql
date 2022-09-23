
BEGIN --Interference Detail SubSubQuery BEGIN
	
	INSERT INTO pole.interference_details ([INT DETAIL FIELDS]) 
		--OUTPUT INSERTED.ID INTO @SubLevel2
		VALUES([INT DETAIL VALUES])
		--SELECT @SubLevel2ID = ID FROM @SubLevel2

END --Interference Detail SubSubQuery END

--[INT DETAIL INSERT]