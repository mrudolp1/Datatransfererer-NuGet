BEGIN --Reinf Geom SubQuery BEGIN

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
		DELETE FROM @PoleReinfSection

		--Add row with data to Geometry table
		INSERT INTO pole.pole_reinf_section OUTPUT INSERTED.ID INTO @PoleReinfSection VALUES ('[INSERT REINF POLE SECTION]')
		SELECT @PoleReinfSectionID = @PoleReinfSectionID FROM @PoleReinfSection

		--Add to xref table
		INSERT INTO pole.pole_reinf_section_xref (pole_structure_id, section_id) VALUES (@PoleID, @PoleReinfSectionID)
	END

END --Reinf Geom SubQuery END

--'[SUBQUERY]'