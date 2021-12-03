BEGIN
	SET @baseSectionID = [BASE SECTION ID]
	IF @baseSectionID IS NULL 
		BEGIN
			DELETE FROM @baseSection 
			INSERT INTO tnx.base_structure OUTPUT INSERTED.ID INTO @baseSection VALUES([ALL BASE SECTION VALUES])
			SELECT @baseSectionID = baseSectionID FROM @baseSection
		END
	INSERT INTO tnx.base_structure_xref VALUES(@baseSectionID, @tnxID)
END