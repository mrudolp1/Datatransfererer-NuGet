BEGIN
	SET @upperSectionID = [UPPER SECTION ID]
	IF @upperSectionID IS NULL 
		BEGIN
			DELETE FROM @upperSection
			INSERT INTO tnx.upper_structure OUTPUT INSERTED.ID INTO @upperstructure VALUES([ALL UPPER SECTION VALUES])
			SELECT @upperSectionID = upperSectionID FROM @upperSection
		END
	INSERT INTO tnx.upper_structure_xref VALUES(@upperSectionID, @tnxID)
END