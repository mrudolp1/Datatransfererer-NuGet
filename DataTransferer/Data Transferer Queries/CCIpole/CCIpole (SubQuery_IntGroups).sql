BEGIN --Interference Group SubQuery BEGIN

	INSERT INTO pole.pole_interference_group OUTPUT INSERTED.ID INTO @IntGroups VALUES ('[INSERT SINGLE INT GROUP]')
	SELECT @IntGroupID = IntGroupID FROM @IntGroups
	--Add to Group XREF table
	INSERT INTO pole.pole_interference_group_xref (pole_structure_id, interference_group_id) VALUES (@PoleID, @IntGroupID)

	--'[DETAIL SUB-SUBQUERY]'

END --Interference Group SubQuery END

--'[SUBQUERY]'