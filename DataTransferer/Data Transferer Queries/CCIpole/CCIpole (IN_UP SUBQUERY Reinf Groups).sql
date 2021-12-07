BEGIN
	DELETE FROM @ReinfGroups
	INSERT INTO pole.pole_reinf_group OUTPUT INSERTED.ID INTO @ReinfGroups VALUES ('[REINF GROUP]')
	SELECT @ReinfGroupID = ReinfGroupID FROM @ReinfGroups
	INSERT INTO pole.pole_reinf_details VALUES ('[REINF DETAILS]') --This would start with (@reinGroupID, value1, value2, value3, value4)
END


--[REINFORCEMENT SUBQUERY]