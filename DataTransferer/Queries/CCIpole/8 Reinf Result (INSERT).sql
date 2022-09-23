
BEGIN --Result SubQuery BEGIN
	
	INSERT INTO pole.reinforcement_results ([REINF RESULT FIELDS]) 
		--OUTPUT INSERTED.ID INTO @SubLevel2
		VALUES([REINF RESULT VALUES])
		--SELECT @SubLevel1ID = ID FROM @SubLevel1

END --Result SubQuery END

--[RESULT INSERT]