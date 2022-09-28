
BEGIN --Result SubQuery BEGIN
	
	
	SELECT @SubLevel1ID = ID FROM pole.reinforced_sections WHERE pole_id = @TopLevelID AND local_section_id = [local_section_id]
	SELECT @SubLevel2ID = ID FROM pole.reinforcements WHERE pole_id = @TopLevelID AND local_group_id = [local_group_id]

	INSERT INTO pole.reinforcement_results ([REINF RESULT FIELDS]) 
		VALUES([REINF RESULT VALUES])

END --Result SubQuery END

--[RESULT INSERT]