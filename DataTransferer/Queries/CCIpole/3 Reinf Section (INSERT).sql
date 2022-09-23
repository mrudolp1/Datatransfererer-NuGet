
BEGIN --Reinf Section SubQuery BEGIN

	--Material DB ID
	SET @SubLevel4ID = [MATL ID]

    --[MATL DB INSERT]


	--Section ID	
	INSERT INTO pole.reinforced_sections ([REINF SECTION FIELDS]) 
		--OUTPUT INSERTED.ID INTO @SubLevel1
		VALUES([REINF SECTION VALUES])
		--SELECT @SubLevel1ID = ID FROM @SubLevel1

END --Reinf Section SubQuery END

--[REINF SECTION INSERT]