
BEGIN --Custom Matl DB SubQuery BEGIN

    IF @SubLevel4ID = NULL
        BEGIN
            IF EXISTS(SELECT * FROM pole.pole_matls WHERE [MATL DB FIELDS AND VALUES])
                SELECT @SubLevel4ID = ID FROM pole.pole_matls WHERE [MATL DB FIELDS AND VALUES]
            ELSE
                INSERT INTO pole.pole_matls ([MATL DB FIELDS])
                OUTPUT INSERTED.ID INTO @SubLevel4
                VALUES([MATL DB VALUES])
                SELECT @SubLevel4ID = ID FROM @SubLevel4
        END

END --Custom Matl DB SubQuery END