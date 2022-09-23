
BEGIN --Reinf Group SubQuery BEGIN

    --[REINF DB INSERT]


	--Group ID	
	INSERT INTO pole.reinforcements ([REINF GROUP FIELDS]) 
		OUTPUT INSERTED.ID INTO @SubLevel1
		VALUES([REINF GROUP VALUES])
		SELECT @SubLevel1ID = ID FROM @SubLevel1


    --[REINF DETAIL INSERT]


END --Reinf Group SubQuery END

--[REINF GROUP INSERT]