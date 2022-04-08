BEGIN --Base Geom SubQuery BEGIN

	--Material ID
	BEGIN
		IF @MatlID IS NULL
			BEGIN
				DELETE FROM @PropMatl
				INSERT INTO pole.matl_prop_flat_plate OUTPUT INSERTED.ID INTO @PropMatl VALUES ('[INSERT MATL PROP]')
				SELECT @MatlID = ID FROM @PropMatl
			END
	END

	--Geometry Section
	BEGIN
		DELETE FROM @PoleSection
	
		--Add row with data to Geometry table
		INSERT INTO pole.pole_section OUTPUT INSERTED.ID INTO @PoleSection VALUES ('[INSERT POLE SECTION]')
		SELECT @PoleSectionID = PoleSectionID FROM @PoleSection

		--Add to xref table
		INSERT INTO pole.pole_section_xref (pole_structure_id, section_id) VALUES (@PoleID, @PoleSectionID)
	END

END --Base Geom SubQuery END

--'[SUBQUERY]'