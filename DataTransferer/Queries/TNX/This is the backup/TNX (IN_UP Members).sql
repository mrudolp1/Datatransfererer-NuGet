BEGIN
	SELECT Top 1 @SubLevel1ID = tbl.ID 
	FROM tnx.members tbl
	WHERE tbl.File = Me.File.ToString.FormatDBValue
	AND tbl.USName = Me.USName.ToString.FormatDBValue
	AND tbl.SIName = Me.SIName.ToString.FormatDBValue
	AND tbl.Values = Me.Values.ToString.FormatDBValue

	IF @SubLevel1ID IS NULL
		INSERT INTO tnx.members OUTPUT INSERTED.ID INTO @TopLevel VALUES([ALL MEMBER VALUES])
		SELECT @SubLevel1ID = ID FROM @SubLevel1
		END

	INSERT INTO tnx.members_xref VALUES(@SubLevel1ID, [TNXID])

	Delete FROM @SubLevel1
	Set @SubLevel1ID = NULL
END