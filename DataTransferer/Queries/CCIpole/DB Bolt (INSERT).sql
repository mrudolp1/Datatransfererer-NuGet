
BEGIN --Custom Bolt DB SubQuery BEGIN

	IF @BoltID IS NULL
        BEGIN
            IF EXISTS(SELECT * FROM gen.pole_bolts WHERE [BOLT DB FIELDS AND VALUES])
                SELECT @BoltID = ID FROM gen.pole_bolts WHERE [BOLT DB FIELDS AND VALUES]
            ELSE
                BEGIN
                    INSERT INTO gen.pole_bolts ([BOLT DB FIELDS])
                    OUTPUT INSERTED.ID INTO @SubLevel3
                    VALUES([BOLT DB VALUES])
                    SELECT @BoltID = ID FROM @SubLevel3
                END
        END

END --Custom Bolt DB SubQuery END