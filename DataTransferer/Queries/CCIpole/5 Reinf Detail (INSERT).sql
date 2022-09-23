
BEGIN --Reinf Detail SubSubQuery BEGIN
	
	INSERT INTO pole.reinforcement_details ([REINF DETAIL FIELDS]) 
		--OUTPUT INSERTED.ID INTO @SubLevel2
		VALUES([REINF DETAIL VALUES])
		--SELECT @SubLevel2ID = ID FROM @SubLevel2

END --Reinf Detail SubSubQuery END

--[REINF DETAIL INSERT]