
BEGIN --Custom Reinf DB SubQuery BEGIN

    --Material DB ID
	SET @SubLevel4ID = [MATL ID]

    --[MATL DB INSERT]
    

    --Bolt DB ID
    SET @TopBoltID = [TOP BOLT ID]

	--[TOP BOLT DB INSERT]


    SET @BotBoltID = [BOT BOLT ID]

	--[BOT BOLT DB INSERT]

   
    --Reinforcement DB ID
	SET @SubLevel2ID = [REINF ID]

    IF @SubLevel2ID = NULL
        BEGIN
            IF EXISTS(SELECT * FROM pole.pole_reinforcements WHERE [REINF DB FIELDS AND VALUES])
                SELECT @SubLevel2ID = ID FROM pole.pole_reinforcements WHERE [REINF DB FIELDS AND VALUES]
            ELSE
                INSERT INTO pole.pole_reinforcements ([REINF DB FIELDS])
                OUTPUT INSERTED.ID INTO @SubLevel2
                VALUES([REINF DB VALUES])
                SELECT @SubLevel2ID = ID FROM @SubLevel2
        END

END --Custom Reinf DB SubQuery END