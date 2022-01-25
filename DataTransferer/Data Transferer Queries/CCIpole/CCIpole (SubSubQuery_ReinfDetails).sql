BEGIN --Reinf Detail SubSubQuery BEGIN

	DELETE FROM @ReinfDetails

	--Add data to Details table
	INSERT INTO pole.pole_reinf_details OUTPUT INSERTED.ID INTO @ReinfDetails VALUES ('[INSERT REINF DETAILS]') --This would start with (@ReinfGroupID, value1, value2, value3, value4)
	SELECT @ReinfDetailID=ReinfDetailID FROM @ReinfDetails

	--Add to xref table
	INSERT INTO pole.pole_reinf_details_xref (reinf_group_id, reinf_id) VALUES (@ReinfGroupID, @ReinfDetailID)


END --Reinf Detail SubSubQuery END

--'[DETAIL SUB-SUBQUERY]'