
BEGIN --Custom Bolt DB SubQuery BEGIN

	IF @BoltID = NULL
        BEGIN
            IF EXISTS(SELECT * FROM pole.pole_bolts WHERE [BOLT DB FIELDS AND VALUES])
                SELECT @BoltID = ID FROM pole.pole_bolts WHERE [BOLT DB FIELDS AND VALUES]
            ELSE
                INSERT INTO pole.pole_bolts ([BOLT DB FIELDS])
                OUTPUT INSERTED.ID INTO @SubLevel3
                VALUES([BOLT DB VALUES])
                SELECT @BoltID = ID FROM @SubLevel3
        END

END --Custom Bolt DB SubQuery END