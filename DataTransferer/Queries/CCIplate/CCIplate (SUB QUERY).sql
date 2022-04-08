	DELETE FROM @Connection
	DELETE FROM @PlateDetail
	--Add row with data to plate_connections table
	INSERT INTO conn.plate_connections OUTPUT INSERTED.ID INTO @Connection VALUES ('[INSERT PLATE CONNECTIONS]')
	SELECT @CID=CID FROM @Connection
	--Add to plate_connections_xref
	INSERT INTO conn.plate_connections_xref (connection_group_id,connection_id) VALUES (@CongrpID,@CID)
	--Add row with data to plate_details table
	INSERT INTO conn.plate_details OUTPUT INSERTED.ID INTO @PlateDetail VALUES ('[INSERT PLATE DETAILS]')
	SELECT @PDID=PDID FROM @PlateDetail
	--Add row with data to base_plate_options table (Baseplate only)
	INSERT INTO conn.base_plate_options VALUES ('[INSERT BASE PLATE OPTIONS]')

	--Add additional row with data to plate_details for second flange connection
	--(Flange 2)INSERT INTO conn.plate_details OUTPUT INSERTED.ID INTO @PlateDetail VALUES ('[INSERT PLATE DETAILS 2]')
	--(Flange 2)SELECT @PDID=PDID FROM @PlateDetail

--SUBQUERY GOES HERE