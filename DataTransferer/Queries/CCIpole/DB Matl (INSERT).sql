
BEGIN --Custom Matl DB SubQuery BEGIN

    IF @SubLevel4ID IS NULL
        BEGIN
            IF EXISTS(SELECT * FROM gen.pole_matls WHERE [MATL DB FIELDS AND VALUES])
                SELECT @SubLevel4ID = ID FROM gen.pole_matls WHERE [MATL DB FIELDS AND VALUES]
            ELSE
                BEGIN
                    INSERT INTO gen.pole_matls ([MATL DB FIELDS])
                    OUTPUT INSERTED.ID INTO @SubLevel4
                    VALUES([MATL DB VALUES])
                    SELECT @SubLevel4ID = ID FROM @SubLevel4
                END
        END

END --Custom Matl DB SubQuery END