BEGIN --Reinf Group SubQuery BEGIN

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


	--'[BOLT SUB-SUBQUERY]'

	--Add to Bolts to XREF table
	IF NOT EXISTS (SELECT * FROM pole.bolt_prop_flat_plate_xref WHERE pole_structure_id = @PoleID AND bolt_id = @BotBoltID)
		INSERT INTO pole.bolt_prop_flat_plate_xref (pole_structure_id, bolt_id) VALUES (@PoleID, @BotBoltID)
	IF NOT EXISTS (SELECT * FROM pole.bolt_prop_flat_plate_xref WHERE pole_structure_id = @PoleID AND bolt_id = @TopBoltID)
		INSERT INTO pole.bolt_prop_flat_plate_xref (pole_structure_id, bolt_id) VALUES (@PoleID, @TopBoltID)


	--Reinforcement ID
	SET @ReinfID = '[REINFORCEMENT ID]'

	--ReinfDNU IF @ReinfID IS NULL
	--ReinfDNU 	BEGIN
	--ReinfDNU 		INSERT INTO pole.memb_prop_flat_plate OUTPUT INSERTED.ID INTO @PropReinf VALUES ('[INSERT REINF PROP]')
	--ReinfDNU 		SELECT @ReinfID = ReinfID FROM @PropReinf
	--ReinfDNU 	END

	--Add to Reinf XREF table
	INSERT INTO pole.memb_prop_flat_plate_xref (pole_structure_id, reinf_id) VALUES (@PoleID, @ReinfID)
	

	--Group ID
	INSERT INTO pole.pole_reinf_group OUTPUT INSERTED.ID INTO @ReinfGroups VALUES ('[INSERT SINGLE REINF GROUP]')
	SELECT @ReinfGroupID = ReinfGroupID FROM @ReinfGroups
	--Add to Group XREF table
	INSERT INTO pole.pole_reinf_group_xref (pole_structure_id, reinf_group_id) VALUES (@PoleID, @ReinfGroupID)

	--'[DETAIL SUB-SUBQUERY]'


END --Reinf Group SubQuery END

--'[SUBQUERY]'