
BEGIN --Reinf Section SubQuery BEGIN

	--Material DB ID
	SET @SubLevel4ID = [MATL ID]

    --[MATL DB SUBQUERY]


	--Section ID	
	INSERT INTO pole.reinforced_sections ([REINF SECTION FIELDS]) 
		VALUES([REINF SECTION VALUES])

END --Reinf Section SubQuery END

--[REINF SECTION SUBQUERY]