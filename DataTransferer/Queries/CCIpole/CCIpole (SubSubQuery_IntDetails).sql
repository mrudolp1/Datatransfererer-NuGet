BEGIN --Interference Detail SubSubQuery BEGIN

	DELETE FROM @IntDetails

	--Add data to Details table
	INSERT INTO pole.pole_interference_details OUTPUT INSERTED.ID INTO @IntDetails VALUES ('[INSERT INTERFERENCE DETAILS]') --This would start with (@ReinfGroupID, value1, value2, value3, value4)
	SELECT @IntDetailID=IntDetailID FROM @IntDetails

	--Add to xref table
	INSERT INTO pole.pole_interference_details_xref (interference_group_id, interference_id) VALUES (@IntGroupID, @IntDetailID)


END --Interference Detail SubSubQuery END

--'[DETAIL SUB-SUBQUERY]'