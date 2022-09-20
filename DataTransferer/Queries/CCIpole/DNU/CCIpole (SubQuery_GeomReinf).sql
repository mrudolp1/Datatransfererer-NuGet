BEGIN --Reinf Geom SubQuery BEGIN
	
	--Material ID	
	SET @MatlID = '[STEEL GRADE ID]'

	--MatlDNU IF @MatlID IS NULL
	--MatlDNU 	BEGIN
	--MatlDNU 		IF EXISTS(SELECT * FROM pole.matl_prop_flat_plate WHERE local_id = '[local_id]' AND name = '[name]' AND fy = '[fy]' AND fu = '[fu]')
	--MatlDNU 			SELECT @MatlID = ID FROM pole.matl_prop_flat_plate WHERE local_id = '[local_id]' AND name = '[name]' AND fy = '[fy]' AND fu = '[fu]'
	--MatlDNU 		ELSE
	--MatlDNU 			BEGIN
	--MatlDNU 				INSERT INTO pole.matl_prop_flat_plate OUTPUT INSERTED.ID INTO @PropMatl VALUES ('[INSERT MATL PROP]')
	--MatlDNU 				SELECT @MatlID = MatlID FROM @PropMatl
	--MatlDNU 			END
	--MatlDNU 	END

	--Add to Matl XREF table
	IF NOT EXISTS (SELECT * FROM pole.matl_prop_flat_plate_xref WHERE pole_structure_id = @PoleID AND matl_id = @MatlID)
		INSERT INTO pole.matl_prop_flat_plate_xref (pole_structure_id, matl_id) VALUES (@PoleID, @MatlID)
	

	--Geometry Section
	INSERT INTO pole.pole_reinf_section OUTPUT INSERTED.ID INTO @PoleReinfSection VALUES ('[INSERT SINGLE POLE SECTION]')
	SELECT @PoleReinfSectionID = PoleReinfSectionID FROM @PoleReinfSection

	--Add to Geom XREF table
	INSERT INTO pole.pole_reinf_section_xref (pole_structure_id, section_id) VALUES (@PoleID, @PoleReinfSectionID)

END --Reinf Geom SubQuery END

--'[SUBQUERY]'