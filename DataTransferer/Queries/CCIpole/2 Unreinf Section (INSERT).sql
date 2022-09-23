
BEGIN --Unreinf Section SubQuery BEGIN

	--Material DB ID
	SET @SubLevel4ID = [MATL ID]

    --[MATL DB INSERT]


	--Section ID	
	INSERT INTO pole.sections ([UNREINF SECTION FIELDS]) 
		--OUTPUT INSERTED.ID INTO @SubLevel1
		VALUES([UNREINF SECTION VALUES])
		--SELECT @SubLevel1ID = ID FROM @SubLevel1

END --Unreinf Section SubQuery END

--[UNREINF SECTION INSERT]