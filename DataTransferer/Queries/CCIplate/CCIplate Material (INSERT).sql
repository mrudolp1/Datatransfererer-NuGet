	SET @SubLevel3ID = [MATERIAL PROPERTY ID]
	IF @SubLevel3ID IS NULL
	BEGIN
		IF EXISTS(SELECT * FROM gen.connection_material_properties WHERE [SELECT])
			BEGIN
				SELECT @SubLevel3ID = ID FROM gen.connection_material_properties WHERE [SELECT]
			END
		ELSE
			BEGIN
				INSERT INTO gen.connection_material_properties ([CCIPLATE MATERIAL FIELDS]) 
				OUTPUT INSERTED.ID INTO @SubLevel3
				VALUES([CCIPLATE MATERIAL VALUES])
				SELECT @SubLevel3ID=ID FROM @SubLevel3
			END
	END