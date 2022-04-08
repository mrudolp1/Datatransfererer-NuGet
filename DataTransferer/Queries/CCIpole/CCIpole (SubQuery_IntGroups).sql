BEGIN --Interference Group SubQuery BEGIN

	DELETE FROM @IntGroups

	--Add row with data to Groups table
	INSERT INTO pole.pole_interference_group OUTPUT INSERTED.ID INTO @IntGroups VALUES ('[INSERT INTERFERENCE GROUP]')
	SELECT @IntGroupID = IntGroupID FROM @IntGroups

	--Add to xref table
	INSERT INTO pole.pole_interference_group_xref (pole_structure_id, interference_group_id) VALUES (@PoleID, @IntGroupID)

	--'[DETAIL SUB-SUBQUERY]'

END --Interference Group SubQuery END

--'[SUBQUERY]'