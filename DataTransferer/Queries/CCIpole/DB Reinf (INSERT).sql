
BEGIN --Custom Reinf DB SubQuery BEGIN

    --Material DB ID
	SET @SubLevel4ID = [MATL ID]

    --[MATL DB SUBQUERY]
    

    --Bolt DB ID
    SET @TopBoltID = [TOP BOLT ID]

	--[TOP BOLT DB SUBQUERY]


    SET @BotBoltID = [BOT BOLT ID]

	--[BOT BOLT DB SUBQUERY]

   
    --Reinforcement DB ID
	SET @SubLevel2ID = NULL

    IF @SubLevel2ID IS NULL
        BEGIN
            IF EXISTS(SELECT * FROM gen.pole_reinforcements WHERE [REINF DB FIELDS AND VALUES])
                SELECT @SubLevel2ID = ID FROM gen.pole_reinforcements WHERE [REINF DB FIELDS AND VALUES]
            ELSE
                BEGIN
                    INSERT INTO gen.pole_reinforcements ([REINF DB FIELDS])
                    OUTPUT INSERTED.ID INTO @SubLevel2
                    VALUES([REINF DB VALUES])
                    SELECT @SubLevel2ID = ID FROM @SubLevel2
                END
        END

END --Custom Reinf DB SubQuery END