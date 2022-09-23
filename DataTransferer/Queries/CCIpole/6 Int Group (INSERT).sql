
BEGIN --Interference Group SubQuery BEGIN

	INSERT INTO pole.interferences ([INT GROUP FIELDS]) 
		OUTPUT INSERTED.ID INTO @SubLevel1
		VALUES([INT GROUP VALUES])
		SELECT @SubLevel1ID = ID FROM @SubLevel1


    --[INT DETAIL INSERT]


END --Interference Group SubQuery END

--[INT GROUP INSERT]