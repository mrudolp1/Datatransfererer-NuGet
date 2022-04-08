BEGIN --Reinf Group SubQuery BEGIN

	--Material ID
	BEGIN
		IF @MatlID IS NULL
			BEGIN
				DELETE FROM @PropMatl
				INSERT INTO pole.matl_prop_flat_plate OUTPUT INSERTED.ID INTO @PropMatl VALUES ('[INSERT MATL PROP]')
				SELECT @MatlID = ID FROM @PropMatl
			END
	END

	--Bolt ID
	BEGIN
		IF @BoltID IS NULL
			BEGIN
				DELETE FROM @PropBolt
				INSERT INTO pole.bolt_prop_flat_plate OUTPUT INSERTED.ID INTO @PropBolt VALUES ('[INSERT BOLT PROP]')
				SELECT @BoltID = ID FROM @PropBolt
			END
	END

	--Reinforcement ID
	BEGIN
		IF @ReinfID IS NULL
			BEGIN
				DELETE FROM @PropReinf
				INSERT INTO pole.memb_prop_flat_plate OUTPUT INSERTED.ID INTO @PropReinf VALUES ('[INSERT REINF PROP]')
				SELECT @ReinfID = ID FROM @PropReinf
			END
	END

	--Group ID
	BEGIN
		DELETE FROM @ReinfGroups

		--Add row with data to Groups table
		INSERT INTO pole.pole_reinf_group OUTPUT INSERTED.ID INTO @ReinfGroups VALUES ('[INSERT REINF GROUP]')
		SELECT @ReinfGroupID = ReinfGroupID FROM @ReinfGroups

		--Add to xref table
		INSERT INTO pole.pole_reinf_group_xref (pole_structure_id, reinf_group_id) VALUES (@PoleID, @ReinfGroupID)

		--'[DETAIL SUB-SUBQUERY]'

	END


END --Reinf Group SubQuery END

--'[SUBQUERY]'