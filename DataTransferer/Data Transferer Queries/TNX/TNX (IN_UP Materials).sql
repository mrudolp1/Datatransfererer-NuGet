BEGIN
	SET @materialID = [MATERIAL ID]
	IF @materialID IS NULL 
		BEGIN
			DELETE FROM @material 
			INSERT INTO tnx.materials OUTPUT INSERTED.ID INTO @material VALUES([ALL MATERIAL VALUES])
			SELECT @materialID = materialID FROM @material
		END
	INSERT INTO tnx.materials_xref VALUES(@materialID, @tnxID)
END